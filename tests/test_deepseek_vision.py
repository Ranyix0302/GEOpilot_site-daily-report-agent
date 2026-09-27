import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from PIL import Image, ImageDraw

from agent.deepseek_vision import DeepSeekVision, prepare_candidates


class DeepSeekVisionTests(unittest.TestCase):
    def make_image(self, draw_callback=None, size=(1200, 800)):
        temp = tempfile.TemporaryDirectory()
        path = Path(temp.name) / "evidence.png"
        image = Image.new("RGB", size, (45, 48, 50))
        if draw_callback:
            draw_callback(ImageDraw.Draw(image))
        image.save(path)
        self.addCleanup(temp.cleanup)
        return str(path)

    def test_operation_tries_upper_right_then_original(self):
        path = self.make_image()
        candidates = prepare_candidates(path, "operation")
        self.assertEqual([item.strategy for item in candidates], ["operation-upper-right", "original"])
        self.assertIsNotNone(candidates[0].crop_box)

    def test_silo_finds_red_led_candidate(self):
        def draw_led(draw):
            draw.rectangle((120, 210, 360, 390), fill=(10, 10, 10))
            draw.rectangle((215, 270, 270, 325), fill=(245, 15, 20))

        path = self.make_image(draw_led)
        candidates = prepare_candidates(path, "silo")
        self.assertEqual(candidates[0].strategy, "silo-active-led")
        left, top, right, bottom = candidates[0].crop_box
        self.assertLess(left, 240)
        self.assertGreater(right, 240)
        self.assertLess(top, 300)
        self.assertGreater(bottom, 300)
        self.assertEqual(candidates[-1].strategy, "original")

    def test_silo_ignores_two_dim_eights_before_bright_33(self):
        def draw_segments(draw, left, color, digit_three=False):
            segments = {
                "top": (left + 10, 70, left + 65, 82),
                "middle": (left + 10, 140, left + 65, 152),
                "bottom": (left + 10, 210, left + 65, 222),
                "upper_left": (left, 82, left + 12, 140),
                "upper_right": (left + 63, 82, left + 75, 140),
                "lower_left": (left, 152, left + 12, 210),
                "lower_right": (left + 63, 152, left + 75, 210),
            }
            enabled = {"top", "middle", "bottom", "upper_right", "lower_right"} if digit_three else set(segments)
            for name, box in segments.items():
                if name in enabled:
                    draw.rectangle(box, fill=color)

        def draw_display(draw):
            draw.rectangle((60, 35, 560, 270), fill=(12, 12, 12))
            draw_segments(draw, 120, (96, 24, 25))
            draw_segments(draw, 220, (96, 24, 25))
            draw_segments(draw, 320, (250, 12, 20), digit_three=True)
            draw_segments(draw, 420, (250, 12, 20), digit_three=True)

        path = self.make_image(draw_display, size=(640, 320))
        candidates = prepare_candidates(path, "silo")
        self.assertEqual(candidates[0].strategy, "silo-active-led")
        self.assertEqual(candidates[0].active_digits, 2)

    def test_invalid_crop_result_falls_back_to_original(self):
        path = self.make_image()
        with patch.dict(os.environ, {"DEEPSEEK_API_KEY": "test-key"}, clear=False):
            vision = DeepSeekVision()
        with patch.object(vision, "_request", side_effect=[
            {"is_target": False, "pile_no": "", "confidence": 0.2},
            {"is_target": True, "pile_no": "G12-R1-16", "confidence": 0.96},
        ]) as request:
            result = vision.extract_stage_json("operation", "prompt", path)
        self.assertEqual(request.call_count, 2)
        self.assertEqual(result["pile_no"], "G12-R1-16")
        self.assertEqual(result["_vision_strategy"], "original")
        self.assertEqual(result["_vision_attempts"], 2)

    def test_request_uses_deepseek_image_payload_and_json_mode(self):
        path = self.make_image(size=(400, 300))
        prepared = prepare_candidates(path, "other")[0]
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "choices": [{"message": {"content": '{"is_target":true,"pile_no":"A-1"}'}}]
        }
        with patch.dict(os.environ, {
            "DEEPSEEK_API_KEY": "test-key",
            "DEEPSEEK_MODEL": "deepseek-flash",
            "DEEPSEEK_BASE_URL": "https://api.deepseek.test",
        }, clear=False):
            vision = DeepSeekVision()
        with patch("agent.deepseek_vision.requests.post", return_value=response) as post:
            result = vision._request("prompt", prepared)
        self.assertEqual(result["pile_no"], "A-1")
        args = post.call_args.kwargs
        self.assertEqual(args["json"]["model"], "deepseek-flash")
        self.assertEqual(args["json"]["response_format"], {"type": "json_object"})
        self.assertEqual(args["json"]["thinking"], {"type": "disabled"})
        url = args["json"]["messages"][0]["content"][1]["image_url"]["url"]
        self.assertTrue(url.startswith("data:image/jpeg;base64,"))
        self.assertNotIn("test-key", str(args["json"]))


if __name__ == "__main__":
    unittest.main()
