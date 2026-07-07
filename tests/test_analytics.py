import unittest
from unittest.mock import patch

from excel_pixel_art.analytics import (
    CLARITY_PROJECT_ID,
    GA4_MEASUREMENT_ID,
    render_analytics,
    render_clarity_analytics,
)


class AnalyticsTest(unittest.TestCase):
    def test_analytics_script_uses_configured_tracking_ids(self):
        with patch("excel_pixel_art.analytics.components.html") as html:
            render_analytics()

        script = html.call_args.args[0]
        self.assertIn("https://www.googletagmanager.com/gtag/js", script)
        self.assertIn(GA4_MEASUREMENT_ID, script)
        self.assertIn("https://www.clarity.ms/tag/", script)
        self.assertIn(CLARITY_PROJECT_ID, script)
        html.assert_called_once_with(script, height=0, width=0)

    def test_legacy_clarity_helper_renders_combined_analytics(self):
        with patch("excel_pixel_art.analytics.components.html") as html:
            render_clarity_analytics()

        script = html.call_args.args[0]
        self.assertIn("https://www.googletagmanager.com/gtag/js", script)
        self.assertIn(GA4_MEASUREMENT_ID, script)
        self.assertIn("https://www.clarity.ms/tag/", script)
        self.assertIn(CLARITY_PROJECT_ID, script)
        html.assert_called_once()


if __name__ == "__main__":
    unittest.main()
