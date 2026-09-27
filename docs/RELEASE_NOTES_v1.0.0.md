# PPT2Fig v1.0.0

首次公开发布。PPT2Fig 在 PowerPoint 功能区提供页面、选中元素、PNG/JPG/TIFF、PDF、EPS 的独立入口及下拉预设，配合本机设置窗口导出科研图片和工程文档素材。

- 图片：DPI 预设、自定义 36–2400 DPI、指定宽度或高度像素，导出前预览画布大小。
- PDF：原生矢量或按 DPI 整页栅格化；多页合并或逐页；保留原幻灯片页面尺寸。
- EPS：逐页或逐组选中元素导出；无法保留为矢量的效果可能栅格化。
- 范围：当前页、已选页、指定页码、全部页、单个/组合/临时多选元素。
- 文件名可包含时间、页码和参数；默认记忆上次确认导出的目录；转换过程使用临时副本。

**验证状态：** macOS PowerPoint 16.90 已完成部分真机验证。Windows 与 PowerPoint 2016/2021/2024/其他 Microsoft 365 构建为目标环境，尚未实机验收。安装需要 Python 3.10+ 和 Poppler；发行包未捆绑 Poppler。详见 [中文 README](../README.md) / [English README](../README.en.md)。

## English summary

Initial public release with separate ribbon controls for slide/object scope and PNG, JPG, TIFF, PDF, and EPS exports. Includes DPI and pixel sizing, vector or raster PDF, page combining, timestamped filenames, and remembered output location. Tested partially on macOS PowerPoint 16.90; Windows and other target PowerPoint builds, including 2016, remain unverified. See the README for installation and test status.
