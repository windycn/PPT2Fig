# 第三方组件

1.0.0 发行包中的 `.ppam` 不打包 Poppler 可执行文件。辅助程序使用用户本机安装的 `pdftocairo`、`pdftops`、`pdfinfo`。若后续发行包内置 Poppler，必须按实际内置版本保留其许可文本、修改记录、对应源码和获取方式；先完成许可证核查再分发。官方源码入口：[Poppler](https://poppler.freedesktop.org/)。

| 组件 | 用途 | 官方许可信息 |
| --- | --- | --- |
| Poppler | PDF → EPS/栅格图像 | [源码和项目主页](https://poppler.freedesktop.org/)；具体文件许可以所发版本为准 |
| Qt for Python / PySide6 | 本机设置窗口 | [Qt for Python 许可说明](https://doc.qt.io/qtforpython-6.10/commercial/index.html) |
| pypdf | PDF 页面、裁切与合并 | [BSD-3-Clause 说明](https://pypdf.readthedocs.io/en/stable/meta/faq.html) |
| Pillow | 图像输出与尺寸调整 | [Pillow 项目许可](https://github.com/python-pillow/Pillow/blob/main/LICENSE) |
| PyInstaller（可选） | 单文件辅助程序构建 | [项目许可](https://pyinstaller.org/en/stable/license.html) |

如果将 PySide6、Pillow、pypdf 或 PyInstaller 打入单文件辅助程序，需要随发行物保留与**实际打包版本**相符的许可文本和通知，并按所选 Qt 许可履行对应义务。当前源码安装方式通过包管理器安装这些组件；本文件是依赖索引，不替代正式发行的逐版本许可清单。
