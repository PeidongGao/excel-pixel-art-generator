import unittest
from unittest.mock import patch

from excel_pixel_art.analytics import CLARITY_PROJECT_ID, render_clarity_analytics


class AnalyticsTest(unittest.TestCase):
    def test_clarity_script_uses_configured_project_id(self):
        with patch("excel_pixel_art.analytics.components.html") as html:
            render_clarity_analytics()

        script = html.call_args.args[0]
        self.assertIn("https://www.clarity.ms/tag/", script)
        self.assertIn(CLARITY_PROJECT_ID, script)
        html.assert_called_once()


if __name__ == "__main__":
    unittest.main()
