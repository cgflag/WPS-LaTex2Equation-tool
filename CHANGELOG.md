# Changelog

All notable changes to this project will be documented in this file.

Format based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

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
