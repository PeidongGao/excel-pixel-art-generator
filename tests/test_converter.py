import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image

from excel_pixel_art.converter import image_to_excel


class ConverterTest(unittest.TestCase):
    def test_image_to_excel_writes_pixel_colors(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            image_path = directory_path / "source.png"
            output_path = directory_path / "pixel_art.xlsx"

            image = Image.new("RGBA", (2, 1))
            image.putpixel((0, 0), (255, 0, 0, 255))
            image.putpixel((1, 0), (0, 128, 255, 255))
            image.save(image_path)

            result = image_to_excel(image_path, output_path, max_size=10)

            self.assertEqual(result, output_path)
            workbook = load_workbook(output_path)
            sheet = workbook.active
            self.assertEqual(workbook.sheetnames, ["Reference", "Template", "Color Index"])
            self.assertEqual(sheet["A1"].fill.fgColor.rgb, "00FF0000")
            self.assertEqual(sheet["B1"].fill.fgColor.rgb, "000080FF")

    def test_workbook_includes_template_and_color_index(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            image_path = directory_path / "source.png"
            output_path = directory_path / "paint_by_number.xlsx"

            image = Image.new("RGBA", (3, 1))
            image.putpixel((0, 0), (255, 0, 0, 255))
            image.putpixel((1, 0), (0, 0, 255, 255))
            image.putpixel((2, 0), (255, 0, 0, 255))
            image.save(image_path)

            image_to_excel(image_path, output_path, max_size=10)

            workbook = load_workbook(output_path)
            template = workbook["Template"]
            color_index = workbook["Color Index"]

            self.assertEqual(color_index["A1"].value, "Number")
            self.assertEqual(color_index["B1"].value, "Hex Code")
            self.assertEqual(color_index["D1"].value, "Cells")
            self.assertEqual(color_index["A2"].value, 1)
            self.assertEqual(color_index["B2"].value, "#FF0000")
            self.assertEqual(color_index["D2"].value, 2)
            self.assertEqual(color_index["A3"].value, 2)
            self.assertEqual(color_index["B3"].value, "#0000FF")
            self.assertEqual(color_index["D3"].value, 1)
            self.assertEqual(template["A1"].value, 1)
            self.assertEqual(template["B1"].value, 2)
            self.assertEqual(template["C1"].value, 1)
            self.assertEqual(template["A1"].fill.fill_type, None)
            self.assertEqual(template["A1"].font.color.rgb, "00A6A6A6")
            self.assertEqual(template["A1"].border.left.style, None)
            self.assertEqual(template.print_options.gridLines, False)
            self.assertEqual(template.sheet_view.showGridLines, True)

    def test_color_count_limits_indexed_palette(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            image_path = directory_path / "source.png"
            output_path = directory_path / "limited_palette.xlsx"

            image = Image.new("RGBA", (4, 1))
            for column, color in enumerate(
                [
                    (255, 0, 0, 255),
                    (0, 255, 0, 255),
                    (0, 0, 255, 255),
                    (255, 255, 0, 255),
                ]
            ):
                image.putpixel((column, 0), color)
            image.save(image_path)

            image_to_excel(image_path, output_path, color_count=2)

            workbook = load_workbook(output_path)
            color_index = workbook["Color Index"]
            self.assertLessEqual(color_index.max_row - 1, 2)

    def test_transparent_pixels_are_left_unfilled(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            image_path = directory_path / "transparent.png"
            output_path = directory_path / "pixel_art.xlsx"

            image = Image.new("RGBA", (1, 1), (255, 0, 0, 0))
            image.save(image_path)

            image_to_excel(image_path, output_path)

            workbook = load_workbook(output_path)
            sheet = workbook.active
            self.assertEqual(sheet["A1"].fill.fill_type, None)

    def test_canvas_preset_creates_fixed_a4_landscape_grid(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            image_path = directory_path / "source.png"
            output_path = directory_path / "a4.xlsx"

            Image.new("RGBA", (200, 100), (20, 40, 60, 255)).save(image_path)

            image_to_excel(
                image_path,
                output_path,
                max_size=20,
                canvas_size="a4",
                orientation="landscape",
            )

            workbook = load_workbook(output_path)
            sheet = workbook.active
            self.assertEqual(sheet.max_column, 20)
            self.assertEqual(sheet.max_row, 14)
            self.assertEqual(sheet.page_setup.orientation, "landscape")
            self.assertEqual(sheet.page_setup.paperSize, 9)

    def test_resolution_creates_exact_grid(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            image_path = directory_path / "source.png"
            output_path = directory_path / "resolution.xlsx"

            Image.new("RGBA", (20, 10), (20, 40, 60, 255)).save(image_path)

            image_to_excel(image_path, output_path, resolution=(12, 7))

            workbook = load_workbook(output_path)
            sheet = workbook.active
            self.assertEqual(sheet.max_column, 12)
            self.assertEqual(sheet.max_row, 7)

    def test_rejects_invalid_size_options(self):
        with self.assertRaises(ValueError):
            image_to_excel("image.png", "output.xlsx", max_size=0)

        with self.assertRaises(ValueError):
            image_to_excel("image.png", "output.xlsx", cell_size=0)

        with self.assertRaises(ValueError):
            image_to_excel("image.png", "output.xlsx", canvas_size="poster")

        with self.assertRaises(ValueError):
            image_to_excel("image.png", "output.xlsx", fit="stretch")

        with self.assertRaises(ValueError):
            image_to_excel("image.png", "output.xlsx", resolution=(0, 10))

        with self.assertRaises(ValueError):
            image_to_excel("image.png", "output.xlsx", color_count=1)


if __name__ == "__main__":
    unittest.main()
