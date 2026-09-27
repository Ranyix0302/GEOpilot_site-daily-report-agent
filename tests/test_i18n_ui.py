import os
import unittest
from unittest.mock import patch

from agent.app import app


class BilingualUiTests(unittest.TestCase):
    def test_home_defaults_to_english_and_loads_language_switcher(self):
        response = app.test_client().get("/")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('<html lang="en">', html)
        self.assertIn('data-lang="en"', html)
        self.assertIn('data-lang="zh"', html)
        self.assertIn('Checking model connection', html)
        self.assertLess(html.index("i18n.js"), html.index("app.js"))

    def test_upload_controls_use_shared_bilingual_layout(self):
        response = app.test_client().get("/")
        html = response.get_data(as_text=True)
        self.assertIn('class="roster-upload-button"', html)
        self.assertIn('data-i18n="rosterTitle"', html)
        self.assertIn('data-i18n="rosterHint"', html)
        self.assertIn('data-i18n="rosterAction"', html)
        self.assertIn('viewBox="0 0 20 20"', html)

        with open("static/app.js", encoding="utf-8") as app_js:
            script = app_js.read()
        self.assertIn('class="upload-icon" aria-hidden="true"', script)
        self.assertNotIn('class="upload-icon">↑</span>', script)

    def test_status_names_competition_gateway_provider(self):
        with patch.dict(os.environ, {
            "LLM_GATEWAY_URL": "https://api.softwaresystems.app",
            "LLM_GATEWAY_API_KEY": "test-key",
            "LLM_MODEL": "competition-model",
        }, clear=True):
            response = app.test_client().get("/api/status")
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertEqual(payload["model_provider"], "Competition LLM Gateway")
        self.assertEqual(payload["model"], "competition-model")
        self.assertTrue(payload["ready"])


if __name__ == "__main__":
    unittest.main()
