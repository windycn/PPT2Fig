from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from .core import ExportError, ExportOptions, check_poppler, new_export_stamp, output_filename
from .ini import read_ini, write_ini_utf16


def parse_page_spec(spec: str, slide_count: int) -> list[int]:
    pages: set[int] = set()
    if not spec.strip():
        raise ExportError("请输入页码，例如 1,3-5")
    for part in spec.replace("，", ",").split(","):
        part = part.strip()
        match = re.fullmatch(r"(\d+)(?:\s*-\s*(\d+))?", part)
        if not match:
            raise ExportError(f"页码格式错误：{part}")
        start = int(match.group(1))
        end = int(match.group(2) or start)
        if not 1 <= start <= end <= slide_count:
            raise ExportError(f"页码超出 1–{slide_count} 范围：{part}")
        pages.update(range(start, end + 1))
    return sorted(pages)


def show_dialog(request_path: Path, response_path: Path) -> None:
    try:
        from PySide6.QtCore import QSettings, Qt
        from PySide6.QtWidgets import (
            QApplication, QComboBox, QDialog, QDialogButtonBox, QFileDialog,
            QFormLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox,
            QPushButton, QSpinBox, QVBoxLayout,
        )
    except ImportError as exc:
        raise ExportError("缺少 PySide6；请安装 ppt2fig[gui]") from exc

    if request_path.suffix.lower() == ".ini":
        parser = read_ini(request_path)
        request = dict(parser["request"])
        request["slide_count"] = int(request["slide_count"])
        request["current_slide"] = int(request["current_slide"])
        request["selected_slides"] = [int(p) for p in request.get("selected_slides", "").split(",") if p]
        request["selected_shape_ids"] = [int(p) for p in request.get("selected_shape_ids", "").split(",") if p]
        if request.get("selection_bounds"):
            request["selection_bounds"] = [float(p) for p in request["selection_bounds"].split(",")]
    else:
        request = json.loads(request_path.read_text(encoding="utf-8"))
    check_poppler()
    app = QApplication.instance() or QApplication(sys.argv)

    class ExportDialog(QDialog):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("PPT2Fig · 导出")
            self.setMinimumWidth(620)
            self.settings = QSettings("PPT2Fig", "PPT2Fig")
            self.export_stamp = new_export_stamp()
            self.preset_resolution = request.get("preset_dpi", "300")
            root = QVBoxLayout(self)
            title = QLabel(f"<b>{request.get('presentation_name', 'PowerPoint')}</b> · {request['slide_count']} 页")
            root.addWidget(title)
            form = QFormLayout()
            form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
            form.setHorizontalSpacing(18)
            form.setVerticalSpacing(12)
            self.form = form
            root.addLayout(form)

            self.scope = QComboBox()
            self.scope.addItem("当前页", "current")
            if request.get("selected_slides"):
                self.scope.addItem("PowerPoint 已选页", "selected")
            self.scope.addItem("指定页码", "range")
            self.scope.addItem("全部页面", "all")
            if request.get("selected_shape_ids"):
                self.scope.addItem("当前选中元素", "shapes")
                self.scope.setCurrentIndex(self.scope.count() - 1)
            preferred_scope = request.get("preset_scope")
            if preferred_scope:
                index = self.scope.findData(preferred_scope)
                if index < 0:
                    raise ExportError("当前没有可用的已选页或元素，请先在 PowerPoint 中选择")
                self.scope.setCurrentIndex(index)
            form.addRow("导出内容", self.scope)

            self.pages = QLineEdit(str(request.get("current_slide", 1)))
            self.pages.setPlaceholderText("例如 1,3-5,8")
            form.addRow("指定页码", self.pages)

            self.format = QComboBox()
            for label, value in (("PNG", "png"), ("JPG", "jpg"), ("TIFF", "tiff"), ("PDF", "pdf"), ("EPS", "eps")):
                self.format.addItem(label, value)
            preferred_format = request.get("preset_format")
            if preferred_format:
                index = self.format.findData(preferred_format)
                if index >= 0:
                    self.format.setCurrentIndex(index)
            form.addRow("格式", self.format)

            self.pdf_mode = QComboBox()
            self.pdf_mode.addItem("保留矢量", "vector")
            self.pdf_mode.addItem("整页按 DPI 栅格化", "raster")
            if request.get("preset_pdf_mode") == "raster":
                self.pdf_mode.setCurrentIndex(1)
            form.addRow("PDF 模式", self.pdf_mode)

            self.pdf_packaging = QComboBox()
            self.pdf_packaging.addItem("合并为一个 PDF", "combined")
            self.pdf_packaging.addItem("每页一个 PDF", "separate")
            if request.get("preset_packaging") == "separate":
                self.pdf_packaging.setCurrentIndex(1)
            form.addRow("PDF 分页", self.pdf_packaging)

            self.resolution = QComboBox()
            form.addRow("分辨率", self.resolution)

            self.value = QSpinBox()
            self.value.setRange(36, 2400)
            self.value.setValue(300)
            form.addRow("自定义数值", self.value)

            self.prefix = QLineEdit(Path(str(request.get("presentation_name", "presentation"))).stem)
            self.prefix.setMinimumWidth(320)
            form.addRow("文件名前缀", self.prefix)

            self.naming = QComboBox()
            self.naming.addItem("工程命名 · 时间 + 页码 + 参数", "detailed")
            self.naming.addItem("简洁命名 · 页码 + 参数", "simple")
            form.addRow("命名方式", self.naming)

            folder_row = QHBoxLayout()
            default_folder = Path.home() / "Documents"
            if not default_folder.is_dir():
                default_folder = Path.home()
            last_folder = Path(str(self.settings.value("last_export_dir", default_folder)))
            if not last_folder.is_dir():
                last_folder = default_folder
            self.folder = QLineEdit(str(last_folder))
            self.folder.setMinimumWidth(320)
            folder_row.addWidget(self.folder)
            browse = QPushButton("选择…")
            browse.clicked.connect(self.choose_folder)
            folder_row.addWidget(browse)
            form.addRow("保存到", folder_row)

            self.preview = QLabel()
            self.preview.setWordWrap(True)
            self.preview.setTextInteractionFlags(Qt.TextSelectableByMouse)
            root.addWidget(self.preview)
            self.filename_preview = QLabel()
            self.filename_preview.setWordWrap(True)
            self.filename_preview.setTextInteractionFlags(Qt.TextSelectableByMouse)
            root.addWidget(self.filename_preview)
            self.note = QLabel("矢量 PDF 保留 PowerPoint 原有图像质量；EPS 不支持透明效果，必要区域可能被栅格化。")
            self.note.setWordWrap(True)
            root.addWidget(self.note)
            self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            self.buttons.button(QDialogButtonBox.Ok).setText("导出")
            self.buttons.button(QDialogButtonBox.Cancel).setText("取消")
            self.buttons.accepted.connect(self.accept)
            self.buttons.rejected.connect(self.reject)
            root.addWidget(self.buttons)

            self.scope.currentIndexChanged.connect(self.refresh)
            self.format.currentIndexChanged.connect(self.refresh)
            self.pdf_mode.currentIndexChanged.connect(self.refresh)
            self.pdf_packaging.currentIndexChanged.connect(self.refresh)
            self.resolution.currentIndexChanged.connect(self.refresh)
            self.value.valueChanged.connect(self.refresh)
            self.pages.textChanged.connect(self.refresh)
            self.prefix.textChanged.connect(self.refresh)
            self.naming.currentIndexChanged.connect(self.refresh)
            self.refresh()

        def choose_folder(self) -> None:
            selected = QFileDialog.getExistingDirectory(self, "选择导出文件夹", self.folder.text())
            if selected:
                self.folder.setText(selected)

        def refresh(self) -> None:
            fmt = self.format.currentData()
            is_pdf = fmt == "pdf"
            vector_pdf = is_pdf and self.pdf_mode.currentData() == "vector"
            self.form.setRowVisible(self.pdf_mode, is_pdf)
            self.form.setRowVisible(self.pdf_packaging, is_pdf)
            if vector_pdf:
                self.note.setText("矢量 PDF 不设置 DPI；PowerPoint 原生导出的位图和效果保持其原有质量。")
            elif is_pdf:
                self.note.setText("栅格 PDF 将整页按指定 DPI 绘制；更高 DPI 会显著增加文件大小和内存占用。")
            elif fmt == "eps":
                self.note.setText("EPS 的 DPI 只影响转换时必须栅格化的透明等效果；文字和图形仍尽量保留矢量。")
            else:
                self.note.setText("选中元素的 PNG/TIFF 使用透明背景；JPG 使用白色背景。")
            self.form.setRowVisible(self.pages, self.scope.currentData() == "range")

            if vector_pdf:
                choices = [("不适用 · 原生矢量", "vector")]
            elif is_pdf or fmt == "eps":
                choices = [(f"{dpi} DPI", dpi) for dpi in (300, 600, 1200, 2400)]
                choices.append(("自定义 DPI · 36–2400", "custom"))
            else:
                choices = [(f"{dpi} DPI", dpi) for dpi in (150, 300, 600, 1200, 2400)]
                choices += [("自定义 DPI · 36–2400", "custom"),
                            ("指定宽度像素", "width"), ("指定高度像素", "height")]
            current = self.resolution.currentData()
            if current not in (None, "vector"):
                self.previous_resolution = current
            preferred = self.preset_resolution
            if preferred is not None:
                preferred = int(preferred) if str(preferred).isdigit() else preferred
                self.preset_resolution = None
            else:
                preferred = getattr(self, "previous_resolution", 300)
            if vector_pdf:
                preferred = "vector"
            if preferred not in [data for _, data in choices]:
                preferred = 300
            if [self.resolution.itemData(i) for i in range(self.resolution.count())] != [data for _, data in choices]:
                self.resolution.blockSignals(True)
                self.resolution.clear()
                for label, data in choices:
                    self.resolution.addItem(label, data)
                self.resolution.setCurrentIndex(self.resolution.findData(preferred))
                self.resolution.blockSignals(False)
            self.resolution.setEnabled(not vector_pdf)
            mode = self.resolution.currentData()
            custom = mode in ("custom", "width", "height")
            self.form.setRowVisible(self.value, custom and not vector_pdf)
            self.form.labelForField(self.value).setText(
                "宽度像素" if mode == "width" else "高度像素" if mode == "height" else "自定义 DPI"
            )
            self.value.blockSignals(True)
            self.value.setRange(64 if mode in ("width", "height") else 36, 30000 if mode in ("width", "height") else 2400)
            self.value.blockSignals(False)
            width_pt = float(request.get("slide_width_pt", 960))
            height_pt = float(request.get("slide_height_pt", 540))
            if self.scope.currentData() == "shapes" and request.get("selection_bounds"):
                bounds = request["selection_bounds"]
                width_pt, height_pt = float(bounds[2]), float(bounds[3])
            if vector_pdf:
                width = height = 0
                effective_dpi = 300
                self.preview.setText(f"原生矢量页面：{width_pt / 72:.2f} × {height_pt / 72:.2f} 英寸；无 DPI 上限。")
            elif mode == "width":
                width = self.value.value()
                height = round(width * height_pt / width_pt)
                effective_dpi = width * 72 / width_pt
            elif mode == "height":
                height = self.value.value()
                width = round(height * width_pt / height_pt)
                effective_dpi = height * 72 / height_pt
            else:
                effective_dpi = self.value.value() if mode == "custom" else int(mode)
                width = round(width_pt / 72 * effective_dpi)
                height = round(height_pt / 72 * effective_dpi)
            if not vector_pdf:
                description = "效果栅格化参考尺寸" if fmt == "eps" else "预计画布"
                warning = "；超出 2 亿像素限制，请降低 DPI" if width * height > 200_000_000 and fmt != "eps" else ""
                self.preview.setText(f"{description}：{width} × {height} 像素（页面 {width_pt / 72:.2f} × {height_pt / 72:.2f} 英寸）{warning}")
                if fmt == "eps" and width * height > 200_000_000:
                    self.note.setText(self.note.text() + " 大范围透明效果在此 DPI 下可能占用大量内存。")
            self.buttons.button(QDialogButtonBox.Ok).setEnabled(vector_pdf or fmt == "eps" or width * height <= 200_000_000)

            current_page = int(request.get("current_slide", 1))
            scope = self.scope.currentData()
            try:
                if scope == "range":
                    pages = tuple(parse_page_spec(self.pages.text(), request["slide_count"]))
                elif scope == "selected":
                    pages = tuple(request["selected_slides"])
                elif scope == "all":
                    pages = tuple(range(1, request["slide_count"] + 1))
                else:
                    pages = (current_page,)
                preview_pages = pages if is_pdf and self.pdf_packaging.currentData() == "combined" else pages[:1]
                options = ExportOptions(Path("preview.pdf"), Path("."), self.prefix.text().strip() or "presentation",
                                        fmt, dpi=effective_dpi,
                                        pixel_width=self.value.value() if mode == "width" else None,
                                        pixel_height=self.value.value() if mode == "height" else None,
                                        pdf_mode=self.pdf_mode.currentData(), selection=scope == "shapes",
                                        naming_mode=self.naming.currentData())
                self.filename_preview.setText("文件名示例：" + output_filename(options, preview_pages, self.export_stamp))
            except ExportError:
                self.filename_preview.setText("文件名示例：请完成页码设置")

        def accept(self) -> None:
            try:
                mode = self.scope.currentData()
                if mode == "range":
                    pages = parse_page_spec(self.pages.text(), request["slide_count"])
                elif mode == "selected":
                    pages = [int(p) for p in request["selected_slides"]]
                elif mode == "all":
                    pages = list(range(1, int(request["slide_count"]) + 1))
                else:
                    pages = [int(request["current_slide"])]
                folder = Path(self.folder.text()).expanduser()
                if not folder.is_dir():
                    raise ExportError("保存文件夹不存在")
                prefix = self.prefix.text().strip()
                if (not prefix or prefix.rstrip(" .") != prefix
                        or any(ord(character) < 32 or character in '<>:"/\\|?*' for character in prefix)):
                    raise ExportError("文件名前缀为空或包含无效字符")
                resolution = self.resolution.currentData()
                value = self.value.value()
                dpi = value if resolution == "custom" else resolution if isinstance(resolution, int) else 300
                if self.format.currentData() not in {"pdf", "eps"} or self.pdf_mode.currentData() == "raster":
                    width_pt = float(request.get("slide_width_pt", 960))
                    height_pt = float(request.get("slide_height_pt", 540))
                    if mode == "shapes" and request.get("selection_bounds"):
                        width_pt, height_pt = map(float, request["selection_bounds"][2:])
                    width = value if resolution == "width" else round(width_pt / 72 * dpi) if resolution != "height" else round(value * width_pt / height_pt)
                    height = value if resolution == "height" else round(height_pt / 72 * dpi) if resolution != "width" else round(value * height_pt / width_pt)
                    if width * height > 200_000_000:
                        raise ExportError("目标画布超过 2 亿像素，请降低 DPI 或像素尺寸")
                result = {
                    "scope": mode,
                    "pages": pages,
                    "format": self.format.currentData(),
                    "pdf_mode": self.pdf_mode.currentData(),
                    "pdf_packaging": self.pdf_packaging.currentData(),
                    "dpi": dpi,
                    "pixel_width": value if resolution == "width" else None,
                    "pixel_height": value if resolution == "height" else None,
                    "output_dir": str(folder),
                    "base_name": prefix,
                    "naming_mode": self.naming.currentData(),
                    "export_stamp": self.export_stamp,
                }
                if response_path.suffix.lower() == ".ini":
                    write_ini_utf16(response_path, "options", result)
                else:
                    response_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
                self.settings.setValue("last_export_dir", str(folder))
                self.settings.sync()
                super().accept()
            except ExportError as exc:
                QMessageBox.warning(self, "无法导出", str(exc))

    window = ExportDialog()
    if window.exec() != QDialog.Accepted:
        response_path.unlink(missing_ok=True)
    del window
    if QApplication.instance() is app:
        app.quit()
