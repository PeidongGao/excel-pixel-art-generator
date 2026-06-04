import io
import unittest

from openpyxl import load_workbook
from PIL import Image

from excel_pixel_art.streamlit_app import _build_digital_workbook, _build_physical_workbook


class UploadedImage:
    name = "source.png"

    def __init__(self):
        image = Image.new("RGBA", (4, 2), (20, 40, 60, 255))
        content = io.BytesIO()
        image.save(content, format="PNG")
        self._content = content.getvalue()

    def getvalue(self):
        return self._content


class StreamlitAppTest(unittest.TestCase):
    def test_digital_and_physical_builders_create_separate_outputs(self):
        uploaded_file = UploadedImage()

        digital_bytes, digital_name = _build_digital_workbook(
            uploaded_file=uploaded_file,
            max_size=128,
            cell_size=3.0,
            color_count=48,
            canvas_size="a4",
            resolution=(8, 6),
            orientation="landscape",
            fit="contain",
            background_color="FFFFFF",
        )
        physical_bytes, physical_name = _build_physical_workbook(
            uploaded_file=uploaded_file,
            max_size=128,
            cell_size=3.0,
            material_color_count=24,
            canvas_size="letter",
            resolution=(5, 7),
            orientation="portrait",
            fit="cover",
            background_color="EEEEEE",
            poster_pages=(2, 2),
            generate_color_masks=True,
            max_color_masks=1,
        )

        digital = load_workbook(io.BytesIO(digital_bytes))
        physical = load_workbook(io.BytesIO(physical_bytes))

        self.assertEqual(digital_name, "source_digital.xlsx")
        self.assertEqual(physical_name, "source_physical.xlsx")
        self.assertEqual(digital.sheetnames, ["Reference", "Template", "Color Index"])
        self.assertEqual(
            physical.sheetnames,
            ["Print Reference", "Print Template", "Material Palette", "Material Mask 001"],
        )
        self.assertEqual((digital["Template"].max_column, digital["Template"].max_row), (8, 6))
        self.assertEqual((physical["Print Template"].max_column, physical["Print Template"].max_row), (5, 7))


if __name__ == "__main__":
    unittest.main()
