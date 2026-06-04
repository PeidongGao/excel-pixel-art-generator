import re
import tempfile
import unittest
from pathlib import Path

from PIL import Image

from excel_pixel_art.poster_print import (
    A4_PAGE_SIZE,
    split_poster_tiles,
    validate_poster_split,
    write_vector_poster_pdf,
)


class PosterPrintTest(unittest.TestCase):
    def test_split_tiles_keep_global_canvas_bounds(self):
        image = Image.new("RGB", (8, 6), "white")

        tiles = split_poster_tiles(image, (2, 2))

        self.assertEqual(len(tiles), 4)
        self.assertEqual(
            [(tile.left, tile.top, tile.right, tile.bottom) for tile in tiles],
            [(0, 0, 4, 3), (4, 0, 8, 3), (0, 3, 4, 6), (4, 3, 8, 6)],
        )

    def test_vector_poster_pdf_contains_three_sets_without_embedded_images(self):
        image = Image.new("RGB", (4, 4), "white")
        for row in range(4):
            for column in range(2):
                image.putpixel((column, row), (220, 20, 20))
        for row in range(4):
            for column in range(2, 4):
                image.putpixel((column, row), (20, 40, 220))

        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "poster.pdf"
            write_vector_poster_pdf(image, output_path, (2, 2))
            pdf_bytes = output_path.read_bytes()

        self.assertEqual(len(re.findall(rb"/Type /Page\b", pdf_bytes)), 12)
        self.assertEqual(pdf_bytes.count(b"/MediaBox"), 12)
        self.assertIn(b"/Subtype /Type1", pdf_bytes)
        self.assertNotIn(b"/Subtype /Image", pdf_bytes)

    def test_rejects_invalid_poster_split(self):
        with self.assertRaises(ValueError):
            split_poster_tiles(Image.new("RGB", (4, 4)), (0, 2))
        with self.assertRaisesRegex(ValueError, "Width 4 is not divisible by 3"):
            split_poster_tiles(Image.new("RGB", (4, 4)), (3, 2))

    def test_poster_split_summary_reports_valid_and_invalid_dimensions(self):
        valid = validate_poster_split((128, 128), (2, 2))
        invalid = validate_poster_split((128, 128), (3, 2))

        self.assertTrue(valid.valid)
        self.assertEqual(valid.cells_per_page, (64, 64))
        self.assertIsNone(valid.reason)
        self.assertFalse(invalid.valid)
        self.assertIsNone(invalid.cells_per_page)
        self.assertEqual(invalid.reason, "Width 128 is not divisible by 3.")

    def test_b5_master_style_256_square_four_by_four_builds_48_a4_pages(self):
        image = Image.new("RGB", (256, 256), (20, 40, 60))

        tiles = split_poster_tiles(image, (4, 4))
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "poster.pdf"
            write_vector_poster_pdf(image, output_path, (4, 4))
            pdf_bytes = output_path.read_bytes()

        self.assertEqual(len(tiles), 16)
        self.assertTrue(all(tile.image.size == (64, 64) for tile in tiles))
        self.assertEqual((tiles[0].left, tiles[0].top, tiles[0].right, tiles[0].bottom), (0, 0, 64, 64))
        self.assertEqual(
            (tiles[-1].left, tiles[-1].top, tiles[-1].right, tiles[-1].bottom),
            (192, 192, 256, 256),
        )
        self.assertEqual(len(re.findall(rb"/Type /Page\b", pdf_bytes)), 48)
        self.assertEqual(pdf_bytes.count(b"/MediaBox"), 48)
        self.assertNotIn(b"/Subtype /Image", pdf_bytes)
        self.assertAlmostEqual(A4_PAGE_SIZE[0], 595.28, places=1)
        self.assertAlmostEqual(A4_PAGE_SIZE[1], 841.89, places=1)


if __name__ == "__main__":
    unittest.main()
