import io
import unittest
from contextlib import redirect_stderr, redirect_stdout

from excel_pixel_art.cli import build_parser, main, parse_resolution


class CliTest(unittest.TestCase):
    def test_help_can_be_built_without_loading_converter_dependencies(self):
        parser = build_parser()

        self.assertEqual(parser.prog, "excel-pixel-art")

    def test_canvas_arguments_are_parsed(self):
        args = build_parser().parse_args(
            [
                "image/photo.jpg",
                "--canvas-size",
                "letter",
                "--orientation",
                "landscape",
                "--fit",
                "cover",
                "--colors",
                "32",
                "--background-color",
                "F8F3E8",
            ]
        )

        self.assertEqual(args.canvas_size, "letter")
        self.assertEqual(args.resolution, None)
        self.assertEqual(args.orientation, "landscape")
        self.assertEqual(args.fit, "cover")
        self.assertEqual(args.colors, 32)
        self.assertEqual(args.background_color, "F8F3E8")

    def test_resolution_argument_is_parsed(self):
        args = build_parser().parse_args(["image/photo.jpg", "--resolution", "160x100"])

        self.assertEqual(args.resolution, (160, 100))

    def test_parse_resolution_rejects_invalid_format(self):
        with redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit):
                build_parser().parse_args(["image/photo.jpg", "--resolution", "160"])

    def test_missing_file_prints_clean_error(self):
        output = io.StringIO()

        with redirect_stdout(output):
            exit_code = main(["missing.png"])

        self.assertEqual(exit_code, 1)
        self.assertIn("Error: Input image not found: missing.png", output.getvalue())


if __name__ == "__main__":
    unittest.main()
