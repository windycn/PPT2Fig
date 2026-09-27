from __future__ import annotations

import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Literal

from PIL import Image
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream, NameObject
from .ini import read_ini


RasterFormat = Literal["png", "jpg", "tiff"]
OutputFormat = Literal["png", "jpg", "tiff", "pdf", "eps"]


class ExportError(RuntimeError):
    pass


@dataclass(frozen=True)
class ExportOptions:
    source_pdf: Path
    output_dir: Path
    base_name: str
    format: OutputFormat
    dpi: float = 300
    pixel_width: int | None = None
    pixel_height: int | None = None
    pdf_mode: Literal["vector", "raster"] = "vector"
    pdf_packaging: Literal["combined", "separate"] = "combined"
    selection: bool = False
    first_slide_number: int = 1
    page_numbers: tuple[int, ...] = ()
    selection_crop: tuple[float, float, float, float] | None = None
    naming_mode: Literal["detailed", "simple"] = "detailed"
    export_stamp: str | None = None

    @classmethod
    def from_file(cls, path: Path) -> "ExportOptions":
        if path.suffix.lower() == ".ini":
            parser = read_ini(path)
            data = dict(parser["export"])
        else:
            data = json.loads(path.read_text(encoding="utf-8"))
        page_numbers = data.get("page_numbers", [])
        if isinstance(page_numbers, str):
            page_numbers = [int(p) for p in page_numbers.split(",") if p.strip()]
        crop = data.get("selection_crop")
        if isinstance(crop, str):
            crop = tuple(float(p) for p in crop.split(",")) if crop else None
        return cls(
            source_pdf=Path(data["source_pdf"]),
            output_dir=Path(data["output_dir"]),
            base_name=data.get("base_name", "presentation"),
            format=data["format"],
            dpi=float(data.get("dpi", 300)),
            pixel_width=int(data["pixel_width"]) if data.get("pixel_width") else None,
            pixel_height=int(data["pixel_height"]) if data.get("pixel_height") else None,
            pdf_mode=data.get("pdf_mode", "vector"),
            pdf_packaging=data.get("pdf_packaging", "combined"),
            selection=str(data.get("selection", "false")).lower() in {"true", "1", "yes"},
            first_slide_number=int(data.get("first_slide_number", 1)),
            page_numbers=tuple(int(p) for p in page_numbers),
            selection_crop=tuple(crop) if crop else None,
            naming_mode=data.get("naming_mode", "detailed"),
            export_stamp=data.get("export_stamp") or None,
        )

    def validate(self) -> None:
        if not self.source_pdf.is_file():
            raise ExportError(f"中间 PDF 不存在：{self.source_pdf}")
        if self.format not in {"png", "jpg", "tiff", "pdf", "eps"}:
            raise ExportError(f"不支持的格式：{self.format}")
        if (not self.base_name or self.base_name.rstrip(" .") != self.base_name
                or any(ord(c) < 32 or c in '<>:"/\\|?*' for c in self.base_name)):
            raise ExportError("文件名前缀包含无效字符")
        if self.naming_mode not in {"detailed", "simple"}:
            raise ExportError("文件命名方式无效")
        if self.export_stamp and not re.fullmatch(r"\d{8}-\d{6}-\d{3}", self.export_stamp):
            raise ExportError("导出时间标识无效")
        if not math.isfinite(self.dpi) or not 36 <= self.dpi <= 2400:
            raise ExportError("DPI 必须介于 36 和 2400 之间")
        if self.pixel_width is not None and not 64 <= self.pixel_width <= 30000:
            raise ExportError("宽度像素必须介于 64 和 30000 之间")
        if self.pixel_height is not None and not 64 <= self.pixel_height <= 30000:
            raise ExportError("高度像素必须介于 64 和 30000 之间")
        if self.pixel_width is not None and self.pixel_height is not None:
            raise ExportError("只需指定宽度或高度像素；另一边自动保持比例")
        if self.pdf_mode not in {"vector", "raster"}:
            raise ExportError("PDF 模式无效")
        if self.pdf_packaging not in {"combined", "separate"}:
            raise ExportError("PDF 输出方式无效")
        if self.page_numbers and any(p < 1 for p in self.page_numbers):
            raise ExportError("页码必须从 1 开始")
        if self.selection_crop and (len(self.selection_crop) != 4 or self.selection_crop[2] <= 0 or self.selection_crop[3] <= 0):
            raise ExportError("选区裁切尺寸无效")


def check_poppler() -> dict[str, str]:
    candidates = []
    configured = os.environ.get("PPT2FIG_POPPLER_DIR")
    if configured:
        candidates.append(Path(configured))
    candidates.append(Path(sys.executable).resolve().parent / "poppler" / "bin")
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent / "poppler")
    candidates += [Path("/opt/homebrew/bin"), Path("/usr/local/bin")]
    found = {}
    for command in ("pdftocairo", "pdftops", "pdfinfo"):
        binary = command + (".exe" if os.name == "nt" else "")
        path = next((str(folder / binary) for folder in candidates if (folder / binary).is_file()), None)
        path = path or shutil.which(command)
        if not path:
            raise ExportError(f"缺少 Poppler 工具 {command}；请安装 Poppler，或设置 PPT2FIG_POPPLER_DIR")
        found[command] = path
    return found


def page_size_points(page) -> tuple[float, float]:
    box = page.mediabox
    return float(box.width), float(box.height)


def output_dimensions(width_pt: float, height_pt: float, options: ExportOptions) -> tuple[int, int, float]:
    if width_pt <= 0 or height_pt <= 0:
        raise ExportError("PDF 页面尺寸无效")
    if options.pixel_width:
        width = options.pixel_width
        height = max(1, round(width * height_pt / width_pt))
        dpi = width * 72 / width_pt
    elif options.pixel_height:
        height = options.pixel_height
        width = max(1, round(height * width_pt / height_pt))
        dpi = height * 72 / height_pt
    else:
        dpi = options.dpi
        width = max(1, round(width_pt / 72 * dpi))
        height = max(1, round(height_pt / 72 * dpi))
    if width * height > 200_000_000:
        raise ExportError(f"目标图像 {width}×{height} 超过 2 亿像素限制")
    return width, height, dpi


def _run(args: list[str], timeout: int = 600) -> None:
    try:
        result = subprocess.run(args, capture_output=True, text=True, timeout=timeout, check=False)
    except subprocess.TimeoutExpired as exc:
        raise ExportError(f"转换超时：{Path(args[0]).name}") from exc
    if result.returncode:
        detail = (result.stderr or result.stdout).strip()[-1500:]
        raise ExportError(f"{Path(args[0]).name} 失败：{detail}")


def _render_png(poppler: dict[str, str], source: Path, page_no: int, dpi: float, path: Path, transparent: bool) -> None:
    prefix = path.with_suffix("")
    args = [poppler["pdftocairo"], "-f", str(page_no), "-l", str(page_no), "-singlefile", "-r", f"{dpi:.6f}", "-png"]
    if transparent:
        args.append("-transp")
    args.extend([str(source), str(prefix)])
    _run(args)
    rendered = prefix.with_suffix(".png")
    if rendered != path:
        rendered.replace(path)


def _save_raster(rendered: Path, destination: Path, fmt: RasterFormat, dimensions: tuple[int, int], dpi: float, transparent: bool) -> None:
    with Image.open(rendered) as image:
        if image.size != dimensions:
            image = image.resize(dimensions, Image.Resampling.LANCZOS)
        if fmt == "png":
            image = image.convert("RGBA" if transparent else "RGB")
            image.save(destination, "PNG", dpi=(dpi, dpi), optimize=True)
        elif fmt == "jpg":
            if image.mode == "RGBA":
                background = Image.new("RGB", image.size, "white")
                background.paste(image, mask=image.getchannel("A"))
                image = background
            else:
                image = image.convert("RGB")
            image.save(destination, "JPEG", quality=95, subsampling=0, dpi=(dpi, dpi), optimize=True)
        else:
            image = image.convert("RGBA" if transparent else "RGB")
            image.save(destination, "TIFF", compression="tiff_lzw", dpi=(dpi, dpi))


def _single_page_pdf(source_page, image: Image.Image, dpi: float) -> bytes:
    rgb = Image.new("RGB", image.size, "white")
    if image.mode == "RGBA":
        rgb.paste(image, mask=image.getchannel("A"))
    else:
        rgb.paste(image.convert("RGB"))
    buffer = io.BytesIO()
    rgb.save(buffer, "PDF", resolution=dpi)
    reader = PdfReader(io.BytesIO(buffer.getvalue()))
    writer = PdfWriter()
    page = writer.add_page(reader.pages[0])
    width, height = page_size_points(source_page)
    page.scale_to(width, height)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def _write_pdf_pages(pages, destination: Path) -> None:
    writer = PdfWriter()
    for page in pages:
        writer.add_page(page)
    with destination.open("wb") as stream:
        writer.write(stream)


def _remove_white_page_background(source: Path, destination: Path) -> bool:
    """Remove PowerPoint's opaque white page fill from a one-page selection PDF.

    PowerPoint paints the page first, even if the temporary slide background is
    hidden.  Removing that paint operation preserves white text and white shapes.
    """
    reader = PdfReader(str(source))
    page = reader.pages[0]
    stream = ContentStream(page.get_contents(), page.pdf)
    page_width, page_height = page_size_points(page)
    white = False
    points: list[tuple[float, float]] = []
    changed = False
    for index, (args, operator) in enumerate(stream.operations[:80]):
        if operator in {b"sc", b"rg"} and len(args) == 3:
            white = all(abs(float(component) - 1) < 0.001 for component in args)
        elif operator == b"m" and len(args) == 2:
            points = [(float(args[0]), float(args[1]))]
        elif operator == b"l" and len(args) == 2 and points:
            points.append((float(args[0]), float(args[1])))
        elif operator == b"re" and len(args) == 4:
            x, y, width, height = map(float, args)
            points = [(x, y), (x + width, y), (x + width, y + height), (x, y + height)]
        elif operator == b"n":
            points = []
        elif operator in {b"f", b"F", b"f*"}:
            if white and len(points) >= 4:
                xs = [point[0] for point in points]
                ys = [point[1] for point in points]
                if (min(xs) <= 0.1 and min(ys) <= 0.1
                        and max(xs) >= page_width - 0.1
                        and max(ys) >= page_height - 0.1):
                    stream.operations[index] = ([], b"n")
                    changed = True
            break  # Never remove any content after the first visible paint.
    if changed:
        page[NameObject("/Contents")] = stream
        _write_pdf_pages([page], destination)
    return changed


def _safe_destination(directory: Path, name: str) -> Path:
    destination = directory / name
    if destination.exists():
        raise ExportError(f"输出文件已存在，为避免覆盖已停止：{destination}")
    return destination


def new_export_stamp() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]


def _short_prefix(value: str) -> str:
    """Leave room for the timestamp and page metadata in a 255-byte filename."""
    result = ""
    for character in value:
        if len((result + character).encode("utf-8")) > 120:
            break
        result += character
    return result


def _quality_tag(options: ExportOptions) -> str:
    if options.format == "pdf" and options.pdf_mode == "vector":
        return "vector"
    if options.format == "eps":
        return f"fx{options.dpi:g}dpi".replace(".", "p")
    if options.pixel_width:
        return f"w{options.pixel_width}px"
    if options.pixel_height:
        return f"h{options.pixel_height}px"
    return f"{options.dpi:g}dpi".replace(".", "p")


def output_filename(options: ExportOptions, page_numbers: tuple[int, ...], stamp: str) -> str:
    if not page_numbers:
        raise ExportError("输出文件缺少页码")
    pad = max(3, len(str(max(page_numbers))))
    if options.selection:
        scope = f"object-{page_numbers[0]:0{pad}d}"
    elif len(page_numbers) == 1:
        scope = f"slide-{page_numbers[0]:0{pad}d}"
    else:
        scope = f"pages-{page_numbers[0]:0{pad}d}-{page_numbers[-1]:0{pad}d}-n{len(page_numbers)}"
    parts = [_short_prefix(options.base_name)]
    if options.naming_mode == "detailed":
        parts.append(stamp)
    parts.append(scope)
    parts.append(_quality_tag(options))
    return "_".join(parts) + "." + options.format


def export(options: ExportOptions) -> list[Path]:
    options.validate()
    poppler = check_poppler()
    options.output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ppt2fig-") as temp_name:
        temp = Path(temp_name)
        reader = PdfReader(str(options.source_pdf))
        if reader.is_encrypted or not reader.pages:
            raise ExportError("中间 PDF 已加密或没有页面")
        source_pdf = options.source_pdf
        if options.selection_crop:
            if len(reader.pages) != 1:
                raise ExportError("元素裁切只支持单页中间 PDF")
            x, y, width, height = options.selection_crop
            source_page = reader.pages[0]
            page_width, page_height = page_size_points(source_page)
            if x < 0 or y < 0 or x + width > page_width + 0.1 or y + height > page_height + 0.1:
                raise ExportError("选区超出幻灯片页面，无法无损裁切")
            for box in (source_page.mediabox, source_page.cropbox, source_page.trimbox, source_page.bleedbox, source_page.artbox):
                box.lower_left = (x, y)
                box.upper_right = (x + width, y + height)
            source_pdf = temp / "cropped.pdf"
            _write_pdf_pages([source_page], source_pdf)
            transparent_pdf = temp / "transparent.pdf"
            if _remove_white_page_background(source_pdf, transparent_pdf):
                source_pdf = transparent_pdf
            reader = PdfReader(str(source_pdf))
        count = len(reader.pages)
        if options.page_numbers and len(options.page_numbers) != count:
            raise ExportError("页码数量与中间 PDF 页数不一致")
        def slide_number(index: int) -> int:
            return options.page_numbers[index - 1] if options.page_numbers else options.first_slide_number + index - 1
        stamp = options.export_stamp or new_export_stamp()
        page_numbers = tuple(slide_number(index) for index in range(1, count + 1))
        staged: list[tuple[Path, Path]] = []
        if options.format == "pdf" and options.pdf_mode == "vector":
            if options.pdf_packaging == "combined":
                name = output_filename(options, page_numbers, stamp)
                target = _safe_destination(options.output_dir, name)
                staged_file = temp / name
                _write_pdf_pages(reader.pages, staged_file)
                staged.append((staged_file, target))
            else:
                for idx, page in enumerate(reader.pages, 1):
                    name = output_filename(options, (slide_number(idx),), stamp)
                    target = _safe_destination(options.output_dir, name)
                    staged_file = temp / name
                    _write_pdf_pages([page], staged_file)
                    staged.append((staged_file, target))
        elif options.format == "pdf" and options.pdf_mode == "raster":
            page_bytes = []
            for idx, page in enumerate(reader.pages, 1):
                width_pt, height_pt = page_size_points(page)
                width, height, dpi = output_dimensions(width_pt, height_pt, options)
                rendered = temp / f"render-{idx}.png"
                _render_png(poppler, source_pdf, idx, dpi, rendered, False)
                with Image.open(rendered) as image:
                    if image.size != (width, height):
                        image = image.resize((width, height), Image.Resampling.LANCZOS)
                    page_bytes.append(_single_page_pdf(page, image, dpi))
            if options.pdf_packaging == "combined":
                name = output_filename(options, page_numbers, stamp)
                target = _safe_destination(options.output_dir, name)
                writer = PdfWriter()
                for data in page_bytes:
                    writer.add_page(PdfReader(io.BytesIO(data)).pages[0])
                staged_file = temp / name
                with staged_file.open("wb") as stream:
                    writer.write(stream)
                staged.append((staged_file, target))
            else:
                for idx, data in enumerate(page_bytes, 1):
                    name = output_filename(options, (slide_number(idx),), stamp)
                    target = _safe_destination(options.output_dir, name)
                    staged_file = temp / name
                    staged_file.write_bytes(data)
                    staged.append((staged_file, target))
        elif options.format == "eps":
            for idx in range(1, count + 1):
                name = output_filename(options, (slide_number(idx),), stamp)
                target = _safe_destination(options.output_dir, name)
                staged_file = temp / name
                _run([poppler["pdftops"], "-f", str(idx), "-l", str(idx), "-eps", "-level3", "-r", str(round(options.dpi)), str(source_pdf), str(staged_file)])
                staged.append((staged_file, target))
        else:
            fmt = options.format
            for idx, page in enumerate(reader.pages, 1):
                width_pt, height_pt = page_size_points(page)
                width, height, dpi = output_dimensions(width_pt, height_pt, options)
                rendered = temp / f"render-{idx}.png"
                transparent = options.selection and fmt in {"png", "tiff"}
                _render_png(poppler, source_pdf, idx, dpi, rendered, transparent)
                name = output_filename(options, (slide_number(idx),), stamp)
                target = _safe_destination(options.output_dir, name)
                staged_file = temp / name
                _save_raster(rendered, staged_file, fmt, (width, height), dpi, transparent)
                staged.append((staged_file, target))

        # Commit only after every page has converted successfully.
        for staged_file, target in staged:
            if not staged_file.is_file() or staged_file.stat().st_size == 0:
                raise ExportError(f"转换结果为空：{staged_file.name}")
        for staged_file, target in staged:
            os.replace(staged_file, target)
    return [target for _, target in staged]
