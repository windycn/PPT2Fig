from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from PIL import Image
from pypdf import PdfReader
from reportlab.pdfgen.canvas import Canvas

from ppt2fig.core import ExportError, ExportOptions, export, output_filename
from ppt2fig.gui import parse_page_spec


def sample_pdf(path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(720, 405))
    canvas.setStrokeColorRGB(0.1, 0.2, 0.8)
    canvas.rect(0, 0, 720, 405, stroke=1, fill=0)
    canvas.setFont("Helvetica", 30)
    canvas.drawString(50, 300, "PPT2Fig Vector Text")
    canvas.showPage()
    canvas.setFillColorRGB(0.9, 0.1, 0.1)
    canvas.rect(100, 90, 250, 130, stroke=0, fill=1)
    canvas.showPage()
    canvas.save()


class ExportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / "input.pdf"
        sample_pdf(self.source)
        self.output = self.root / "output"

    def tearDown(self) -> None:
        self.temp.cleanup()

    def options(self, fmt: str, **kwargs) -> ExportOptions:
        return ExportOptions(self.source, self.output, "论文图片", fmt,
                             export_stamp="20260927-101530-123", **kwargs)

    def test_png_at_dpi_and_tiff_by_pixel_width(self) -> None:
        result = export(self.options("png", dpi=300, page_numbers=(2, 5)))
        self.assertEqual([p.name for p in result], [
            "论文图片_20260927-101530-123_slide-002_300dpi.png",
            "论文图片_20260927-101530-123_slide-005_300dpi.png",
        ])
        with Image.open(result[0]) as image:
            self.assertEqual(image.size, (3000, 1688))
        result = export(self.options("tiff", pixel_width=1800))
        with Image.open(result[0]) as image:
            self.assertEqual(image.size, (1800, 1012))

    def test_vector_and_raster_pdf_exact_page_size(self) -> None:
        vector = export(self.options("pdf", pdf_mode="vector"))[0]
        vector_reader = PdfReader(str(vector))
        self.assertEqual(len(vector_reader.pages), 2)
        self.assertEqual(float(vector_reader.pages[0].mediabox.width), 720)
        self.assertEqual(float(vector_reader.pages[0].mediabox.height), 405)
        self.assertIn("PPT2Fig Vector Text", vector_reader.pages[0].extract_text())
        raster = export(self.options("pdf", pdf_mode="raster", dpi=150, pdf_packaging="separate", page_numbers=(2, 5)))
        self.assertEqual(len(raster), 2)
        page = PdfReader(str(raster[0])).pages[0]
        self.assertAlmostEqual(float(page.mediabox.width), 720, places=4)
        self.assertAlmostEqual(float(page.mediabox.height), 405, places=4)
        self.assertEqual(page.extract_text(), "")

    def test_eps_and_selection_crop(self) -> None:
        eps = export(self.options("eps", dpi=600))
        self.assertEqual(len(eps), 2)
        self.assertIn("%%BoundingBox:", eps[0].read_text("latin-1", errors="ignore"))
        one_page = self.root / "one-page.pdf"
        writer = __import__("pypdf").PdfWriter()
        writer.add_page(PdfReader(str(self.source)).pages[0])
        with one_page.open("wb") as stream:
            writer.write(stream)
        cropped = export(ExportOptions(one_page, self.output, "selection", "pdf", selection=True, selection_crop=(40, 60, 300, 180)))
        page = PdfReader(str(cropped[0])).pages[0]
        self.assertEqual(float(page.mediabox.width), 300)
        self.assertEqual(float(page.mediabox.height), 180)

    def test_selection_png_removes_page_fill_but_keeps_white_artwork(self) -> None:
        source = self.root / "white-page.pdf"
        canvas = Canvas(str(source), pagesize=(720, 405))
        canvas.setFillColorRGB(1, 1, 1)
        canvas.rect(0, 0, 720, 405, stroke=0, fill=1)
        canvas.setFillColorRGB(0, 0.4, 0.8)
        canvas.rect(100, 100, 220, 160, stroke=0, fill=1)
        canvas.setFillColorRGB(1, 1, 1)
        canvas.rect(160, 120, 20, 20, stroke=0, fill=1)
        canvas.save()
        result = export(ExportOptions(source, self.output, "selection", "png",
                                      dpi=72, selection=True,
                                      selection_crop=(0, 0, 720, 405)))[0]
        with Image.open(result) as image:
            self.assertEqual(image.getpixel((0, 0))[3], 0)
            self.assertEqual(image.getpixel((170, 275)), (255, 255, 255, 255))

    def test_existing_file_is_not_overwritten(self) -> None:
        result = export(self.options("pdf"))[0]
        original = result.read_bytes()
        with self.assertRaises(ExportError):
            export(self.options("pdf"))
        self.assertEqual(result.read_bytes(), original)

    def test_page_range(self) -> None:
        self.assertEqual(parse_page_spec("1,3-5,3", 6), [1, 3, 4, 5])
        with self.assertRaises(ExportError):
            parse_page_spec("2-8", 6)

    def test_filename_includes_time_scope_and_quality(self) -> None:
        pdf = self.options("pdf", pdf_mode="vector")
        self.assertEqual(output_filename(pdf, (1, 3, 5), pdf.export_stamp),
                         "论文图片_20260927-101530-123_pages-001-005-n3_vector.pdf")
        eps = self.options("eps", dpi=1200, selection=True)
        self.assertEqual(output_filename(eps, (4,), eps.export_stamp),
                         "论文图片_20260927-101530-123_object-004_fx1200dpi.eps")
        short = self.options("png", dpi=600, naming_mode="simple")
        self.assertEqual(output_filename(short, (2,), short.export_stamp),
                         "论文图片_slide-002_600dpi.png")


if __name__ == "__main__":
    unittest.main()
