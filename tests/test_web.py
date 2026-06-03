import unittest

from excel_pixel_art.web import _parse_multipart_form, _parse_optional_resolution, _safe_stem


class WebTest(unittest.TestCase):
    def test_safe_stem_normalizes_uploaded_filename(self):
        self.assertEqual(_safe_stem("Under The Wave!.jpg"), "under-the-wave")
        self.assertEqual(_safe_stem("..."), "image")

    def test_parse_multipart_form_extracts_fields_and_file(self):
        boundary = "boundary-123"
        body = (
            b"--boundary-123\r\n"
            b'Content-Disposition: form-data; name="max_size"\r\n\r\n'
            b"64\r\n"
            b"--boundary-123\r\n"
            b'Content-Disposition: form-data; name="image"; filename="photo.jpg"\r\n'
            b"Content-Type: image/jpeg\r\n\r\n"
            b"fake-image-bytes\r\n"
            b"--boundary-123--\r\n"
        )

        form = _parse_multipart_form(f"multipart/form-data; boundary={boundary}", body)

        self.assertEqual(form.fields["max_size"], "64")
        self.assertEqual(form.files["image"].filename, "photo.jpg")
        self.assertEqual(form.files["image"].content, b"fake-image-bytes")

    def test_parse_optional_resolution(self):
        form = _parse_multipart_form(
            "multipart/form-data; boundary=boundary-123",
            (
                b"--boundary-123\r\n"
                b'Content-Disposition: form-data; name="resolution_width"\r\n\r\n'
                b"160\r\n"
                b"--boundary-123\r\n"
                b'Content-Disposition: form-data; name="resolution_height"\r\n\r\n'
                b"100\r\n"
                b"--boundary-123--\r\n"
            ),
        )

        self.assertEqual(_parse_optional_resolution(form), (160, 100))


if __name__ == "__main__":
    unittest.main()
