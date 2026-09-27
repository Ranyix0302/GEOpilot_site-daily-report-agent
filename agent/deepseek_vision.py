"""DeepSeek image recognition with automatic, local region preparation.

The engineer uploads an ordinary WhatsApp export.  This module works only on
the extracted media files: it never requires a user to crop an image manually.
Likely target regions are tried first and the original image is retained as a
fallback so a conservative crop cannot hide useful evidence.
"""
from __future__ import annotations

import base64
import io
import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests
from PIL import Image, ImageEnhance, ImageFilter, ImageOps


class DeepSeekError(RuntimeError):
    pass


@dataclass(frozen=True)
class PreparedImage:
    data: bytes
    mime_type: str
    strategy: str
    crop_box: tuple[int, int, int, int] | None = None
    active_digits: int | None = None


def _clamp_box(
    box: tuple[float, float, float, float], width: int, height: int
) -> tuple[int, int, int, int]:
    left = max(0, min(width - 1, int(box[0])))
    top = max(0, min(height - 1, int(box[1])))
    right = max(left + 1, min(width, int(box[2])))
    bottom = max(top + 1, min(height, int(box[3])))
    return left, top, right, bottom


def _encode_image(image: Image.Image, strategy: str, crop_box=None) -> PreparedImage:
    """Enhance a region without changing the evidence image stored by the app."""
    image = image.convert("RGB")
    max_side = max(image.size)
    if max_side > 4096:
        scale = 4096 / max_side
        image = image.resize(
            (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
            Image.Resampling.LANCZOS,
        )
    # Small display crops benefit from local upscaling before API-side resizing.
    if max(image.size) < 1400:
        scale = min(3.0, 1400 / max(image.size))
        image = image.resize(
            (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
            Image.Resampling.LANCZOS,
        )
    image = ImageOps.autocontrast(image, cutoff=1)
    image = ImageEnhance.Contrast(image).enhance(1.18)
    image = image.filter(ImageFilter.UnsharpMask(radius=1.4, percent=145, threshold=3))
    output = io.BytesIO()
    image.save(output, format="JPEG", quality=94, optimize=True)
    return PreparedImage(output.getvalue(), "image/jpeg", strategy, crop_box)


def _otsu_threshold(values: list[int]) -> int:
    """Return an adaptive split between dim red outlines and lit LED segments."""
    if not values:
        return 255
    if max(values) - min(values) < 24:
        return min(values)
    histogram = [0] * 256
    for value in values:
        histogram[value] += 1
    total = len(values)
    weighted_total = sum(index * count for index, count in enumerate(histogram))
    background_weight = 0
    background_sum = 0
    best_variance = -1.0
    best_threshold = min(values)
    for threshold, count in enumerate(histogram):
        background_weight += count
        if background_weight == 0:
            continue
        foreground_weight = total - background_weight
        if foreground_weight == 0:
            break
        background_sum += threshold * count
        background_mean = background_sum / background_weight
        foreground_mean = (weighted_total - background_sum) / foreground_weight
        variance = background_weight * foreground_weight * (background_mean - foreground_mean) ** 2
        if variance > best_variance:
            best_variance = variance
            best_threshold = threshold
    return best_threshold


def _is_red_led_pixel(red: int, green: int, blue: int) -> bool:
    # The ratio excludes orange/yellow equipment paint while retaining both
    # bright and dim red LED segments.
    return (
        red >= 70
        and red >= green * 1.8
        and red >= blue * 1.45
        and red - max(green, blue) >= 35
    )


def _active_digit_count(points: list[tuple[int, int]], width: int) -> int | None:
    if not points:
        return None
    counts: dict[int, int] = {}
    for x, _ in points:
        counts[x] = counts.get(x, 0) + 1
    active_columns = sorted(x for x, count in counts.items() if count >= 2)
    if not active_columns:
        return None
    groups: list[list[int]] = [[active_columns[0]]]
    split_gap = max(5, round(width * 0.012))
    for x in active_columns[1:]:
        if x - groups[-1][-1] > split_gap:
            groups.append([x])
        else:
            groups[-1].append(x)
    substantial = [group for group in groups if len(group) >= 3]
    return len(substantial) or None


def _encode_active_led(
    image: Image.Image,
    crop_box: tuple[int, int, int, int],
    active_digits: int | None,
) -> PreparedImage:
    """Suppress unlit seven-segment outlines instead of enhancing their ghost image."""
    image = image.convert("RGB")
    source_pixels = (
        image.get_flattened_data() if hasattr(image, "get_flattened_data") else image.getdata()
    )
    red_values = [
        red
        for red, green, blue in source_pixels
        if _is_red_led_pixel(red, green, blue)
    ]
    threshold = max(100, _otsu_threshold(red_values))
    pixels = []
    source_pixels = (
        image.get_flattened_data() if hasattr(image, "get_flattened_data") else image.getdata()
    )
    for red, green, blue in source_pixels:
        if _is_red_led_pixel(red, green, blue) and red > threshold:
            pixels.append((255, min(65, green), min(65, blue)))
        else:
            gray = int((red * 0.21 + green * 0.72 + blue * 0.07) * 0.22)
            pixels.append((gray, gray, gray))
    masked = Image.new("RGB", image.size)
    masked.putdata(pixels)
    if max(masked.size) < 1400:
        scale = min(3.0, 1400 / max(masked.size))
        masked = masked.resize(
            (max(1, round(masked.width * scale)), max(1, round(masked.height * scale))),
            Image.Resampling.LANCZOS,
        )
    masked = masked.filter(ImageFilter.UnsharpMask(radius=1.2, percent=120, threshold=3))
    output = io.BytesIO()
    masked.save(output, format="JPEG", quality=94, optimize=True)
    return PreparedImage(output.getvalue(), "image/jpeg", "silo-active-led", crop_box, active_digits)


def _operation_candidates(image: Image.Image) -> list[PreparedImage]:
    """Try the usual upper-right control-screen input region before the full photo."""
    width, height = image.size
    candidates: list[PreparedImage] = []
    if width >= 500 and height >= 350:
        # The D.S.M overview layout places Cluster No./No. in the upper-right.
        # Relative coordinates tolerate different camera resolutions and framing.
        box = _clamp_box((width * 0.46, height * 0.02, width * 0.99, height * 0.45), width, height)
        candidates.append(_encode_image(image.crop(box), "operation-upper-right", box))
    candidates.append(_encode_image(image, "original"))
    return candidates


def _red_display_box(image: Image.Image) -> tuple[tuple[int, int, int, int], int | None] | None:
    """Find a compact cluster of bright red pixels typical of an LED display.

    Work on a small proxy image, use coarse spatial voting instead of a heavy CV
    dependency, and return coordinates in the original image.
    """
    proxy = image.convert("RGB")
    scale = min(1.0, 800 / max(proxy.size))
    if scale < 1:
        proxy = proxy.resize(
            (max(1, round(proxy.width * scale)), max(1, round(proxy.height * scale))),
            Image.Resampling.BILINEAR,
        )
    width, height = proxy.size
    cols = rows = 12
    cell_w = max(1, (width + cols - 1) // cols)
    cell_h = max(1, (height + rows - 1) // rows)
    counts = [[0 for _ in range(cols)] for _ in range(rows)]
    red_pixels: list[tuple[int, int, int]] = []
    for y in range(height):
        for x in range(width):
            red, green, blue = proxy.getpixel((x, y))
            if _is_red_led_pixel(red, green, blue):
                red_pixels.append((x, y, red))
    if len(red_pixels) < 4:
        return None

    threshold = max(100, _otsu_threshold([red for _, _, red in red_pixels]))
    red_points = [(x, y) for x, y, red in red_pixels if red > threshold]
    if len(red_points) < 4:
        red_points = [(x, y) for x, y, red in red_pixels if red >= threshold]
    if len(red_points) < 4:
        return None
    for x, y in red_points:
        counts[min(rows - 1, y // cell_h)][min(cols - 1, x // cell_w)] += 1

    best_score = 0
    best_cell: tuple[int, int] | None = None
    for row in range(rows):
        for col in range(cols):
            score = sum(
                counts[near_row][near_col]
                for near_row in range(max(0, row - 1), min(rows, row + 2))
                for near_col in range(max(0, col - 1), min(cols, col + 2))
            )
            if score > best_score:
                best_score, best_cell = score, (col, row)
    if not best_cell or best_score < 4:
        return None

    col, row = best_cell
    search = (
        max(0, (col - 1) * cell_w),
        max(0, (row - 1) * cell_h),
        min(width, (col + 2) * cell_w),
        min(height, (row + 2) * cell_h),
    )
    local = [
        (x, y)
        for x, y in red_points
        if search[0] <= x < search[2] and search[1] <= y < search[3]
    ]
    if len(local) < 4:
        return None
    min_x = min(x for x, _ in local)
    max_x = max(x for x, _ in local) + 1
    min_y = min(y for _, y in local)
    max_y = max(y for _, y in local) + 1
    center_x = (min_x + max_x) / 2
    center_y = (min_y + max_y) / 2
    # Keep enough dark panel around the digits for the model to distinguish an
    # LED counter from a red pointer, painted mark, or timestamp.
    crop_w = max(width * 0.24, (max_x - min_x) * 5.0)
    crop_h = max(height * 0.20, (max_y - min_y) * 5.0)
    proxy_box = _clamp_box(
        (center_x - crop_w / 2, center_y - crop_h / 2,
         center_x + crop_w / 2, center_y + crop_h / 2),
        width,
        height,
    )
    inverse = 1 / scale
    original_box = _clamp_box(tuple(value * inverse for value in proxy_box), image.width, image.height)
    return original_box, _active_digit_count(local, width)


def prepare_candidates(image_path: str, stage: str) -> list[PreparedImage]:
    try:
        with Image.open(image_path) as opened:
            image = ImageOps.exif_transpose(opened).convert("RGB")
    except (OSError, ValueError) as exc:
        raise DeepSeekError(f"无法读取图片 {Path(image_path).name}。") from exc
    if stage == "operation":
        return _operation_candidates(image)
    if stage == "silo":
        candidates: list[PreparedImage] = []
        detected = _red_display_box(image)
        if detected:
            box, active_digits = detected
            candidates.append(_encode_active_led(image.crop(box), box, active_digits))
        candidates.append(_encode_image(image, "original"))
        return candidates
    return [_encode_image(image, "original")]


class DeepSeekVision:
    def __init__(self, timeout: int = 90):
        gateway_url = os.getenv("LLM_GATEWAY_URL", "").strip()
        self.gateway_mode = bool(gateway_url)
        if self.gateway_mode:
            self.api_key = os.getenv("LLM_GATEWAY_API_KEY", "").strip()
            self.model = (
                os.getenv("LLM_MODEL", "global.anthropic.claude-sonnet-4-5-20250929-v1:0").strip()
                or "global.anthropic.claude-sonnet-4-5-20250929-v1:0"
            )
            gateway_url = gateway_url.rstrip("/")
            self.base_url = gateway_url if gateway_url.endswith("/v1") else f"{gateway_url}/v1"
        else:
            self.api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
            self.model = os.getenv("DEEPSEEK_MODEL", "deepseek-flash").strip() or "deepseek-flash"
            self.base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/")
        self.timeout = int(os.getenv("DEEPSEEK_TIMEOUT_SECONDS", str(timeout)))

    def available(self) -> bool:
        return bool(self.api_key)

    @staticmethod
    def _extract_json_text(text: str) -> dict[str, Any]:
        cleaned = text.replace("```json", "").replace("```", "").strip()
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start < 0 or end <= start:
            raise DeepSeekError("DeepSeek 返回结果不是 JSON。")
        try:
            parsed = json.loads(cleaned[start:end + 1])
        except json.JSONDecodeError as exc:
            raise DeepSeekError("DeepSeek 返回了格式不正确的 JSON。") from exc
        if not isinstance(parsed, dict):
            raise DeepSeekError("DeepSeek 返回结果格式不是 JSON 对象。")
        return parsed

    def _request(self, prompt: str, prepared: PreparedImage) -> dict[str, Any]:
        if not self.api_key:
            variable = "LLM_GATEWAY_API_KEY" if self.gateway_mode else "DEEPSEEK_API_KEY"
            raise DeepSeekError(f"尚未设置 {variable}。请把密钥放入项目 .env 文件后重启。")
        encoded = base64.b64encode(prepared.data).decode("ascii")
        payload = {
            "model": self.model,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {
                        "url": f"data:{prepared.mime_type};base64,{encoded}"
                    }},
                ],
            }],
            "temperature": 0,
            "max_tokens": 300,
        }
        if not self.gateway_mode:
            payload.update({
                "response_format": {"type": "json_object"},
                # OCR/classification does not benefit from billed chain-of-thought.
                "thinking": {"type": "disabled"},
            })
        try:
            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers={"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"},
                json=payload,
                timeout=self.timeout,
            )
            response.raise_for_status()
            body = response.json()
            content = body.get("choices", [{}])[0].get("message", {}).get("content", "")
            if isinstance(content, list):
                content = "".join(str(item.get("text", "")) for item in content if isinstance(item, dict))
            if not isinstance(content, str) or not content.strip():
                raise DeepSeekError("DeepSeek 完成了请求，但没有返回识别结果。")
            return self._extract_json_text(content)
        except DeepSeekError:
            raise
        except requests.Timeout as exc:
            raise DeepSeekError(f"单张图片识别超过 {self.timeout} 秒，已停止本次请求。") from exc
        except requests.RequestException as exc:
            detail = ""
            response = getattr(exc, "response", None)
            if response is not None:
                try:
                    detail = str(response.json().get("error", {}).get("message", ""))
                except Exception:
                    detail = ""
            suffix = f"：{detail[:200]}" if detail else ""
            raise DeepSeekError(f"DeepSeek API 请求失败{suffix}") from exc
        except (KeyError, IndexError, TypeError, ValueError) as exc:
            raise DeepSeekError("DeepSeek API 返回了无法解析的响应。") from exc

    @staticmethod
    def _usable(stage: str, result: dict[str, Any]) -> bool:
        if not result.get("is_target"):
            return False
        if stage == "operation":
            return bool(str(result.get("pile_no") or "").strip())
        if stage == "silo":
            return result.get("red_number") is not None
        return True

    def extract_stage_json(self, stage: str, prompt: str, image_path: str) -> dict[str, Any]:
        candidates = prepare_candidates(image_path, stage)
        last_result: dict[str, Any] | None = None
        errors: list[DeepSeekError] = []
        for attempt, prepared in enumerate(candidates, start=1):
            try:
                candidate_prompt = prompt
                if stage == "silo" and prepared.active_digits:
                    candidate_prompt += (
                        f" 本地图像分析检测到约 {prepared.active_digits} 个真正发光的数字位；"
                        "暗红色但未发光的七段轮廓不属于读数，绝不能按数字 8 计入。"
                    )
                result = self._request(candidate_prompt, prepared)
                result["_vision_strategy"] = prepared.strategy
                result["_vision_attempts"] = attempt
                last_result = result
                if self._usable(stage, result):
                    return result
            except DeepSeekError as exc:
                errors.append(exc)
        if last_result is not None:
            return last_result
        if errors:
            raise errors[-1]
        raise DeepSeekError("没有可用于识别的图片区域。")
