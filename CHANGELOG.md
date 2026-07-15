# Changelog

All notable changes to this project will be documented in this file.

Format based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Changed

- README：中文为主 + 锚点切换 English；定位「AI `$LaTeX$` → WPS/Word 原生公式」（口语钩子注明偏 Gemini）；论文/专利仅作场景说明；exe 另发
- `pyproject.toml`：description / keywords 对齐 AI×docx×WPS（去掉 patent keyword）
- 新增 `assets/before-after.svg` 供 README 展示（中文标注）

### Notes

- 仓库不再跟踪本地 `scripts/`、`dist/`；桌面 exe 若有发布，走站外渠道，不作为本仓库主推路径

## [1.1.2] - 2026-05-29

### Fixed

- GUI 拖放：用队列在主线程处理，修复 GIL 崩溃

## [1.1.0] - 2026-05-29

### Added

- Desktop GUI (`gui/app.py`, CustomTkinter): file picker, drag-and-drop (Windows), font size, XSL detection
- Console entry: `wps-latex2equation-gui`
- `convert_docx()` returns list of failed formula texts for GUI display
- Optional install: `pip install -e ".[gui]"`

## [1.0.0] - 2026-05-29

### Added

- Single-file CLI `convert_latex_docx.py` for `$...$` / `$$...$$` → native OMML
- Inline vs display equation detection; display equations centered with `(1)(2)...` tab stops
- Font size override (`--size`), tab layout modes (`--tab-mode page|chars`)
- `pyproject.toml` with `patent-math` / `docx-latex-math` console entry points
- Example docx: `examples/demo_before.docx`, `examples/demo_after.docx`
- GitHub Actions: unit tests (all platforms) + optional Windows conversion smoke test
- Legacy WPS/Word VBA macro moved to `legacy/`

### Notes

- Requires `MML2OMML.XSL` from Microsoft Office install (no Word COM / OMath API at runtime)
- Output opens in **WPS** and Microsoft Word as native editable equations

[1.0.0]: https://github.com/cgflag/WPS-LaTex2Equation-tool/releases/tag/v1.0.0
