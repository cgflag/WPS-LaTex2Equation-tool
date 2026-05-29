#!/usr/bin/env python3
"""
patent-math-tool — Convert $...$ / $$...$$ LaTeX in .docx to native Word OMML equations.

Works with Microsoft Word and WPS (no MathType). Block formulas get centered tab stops
and sequential numbering (1), (2), ...

Usage:
  python convert_latex_docx.py input.docx [output.docx] [--size 小四|四号|五号|24]

Environment:
  PATENT_MATH_MML2OMML — path to MML2OMML.XSL (from Microsoft Office install)
"""

from __future__ import annotations

__version__ = "1.1.2"

import argparse
import os
import re
import sys
import tempfile
import zipfile
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path

from lxml import etree
from latex2mathml.converter import convert as latex_to_mathml

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NSMAP = {"w": WNS, "m": MNS}

FORMULA_RE = re.compile(r"\$\$([^$]+)\$\$|\$([^$]+)\$")

# Word w:sz is half-points
CN_FONT_SIZE = {
    "初号": 84,
    "小初": 72,
    "一号": 52,
    "小一": 48,
    "二号": 44,
    "小二": 36,
    "三号": 32,
    "小三": 30,
    "四号": 28,
    "小四": 24,
    "五号": 21,
    "小五": 18,
}

# Brackets and punctuation -> upright (non-italic)
ROMAN_MATH_CHARS = set("()[]{},;:|=<>+-*/\\^& ")

# Default A4 twips (dxa): 21cm page width, 3.17cm margins
DEFAULT_PAGE_WIDTH_DXA = 11906
DEFAULT_MARGIN_DXA = 1800


@dataclass
class TabLayoutConfig:
    """Block equation tab stops: center + right."""

    mode: str = "page"  # "page" = from paper size; "chars" = from character count
    center_chars: float = 21.0
    right_chars: float = 44.0

XSL_CANDIDATES = [
    Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"),
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL"),
    Path(r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL"),
    Path(r"C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL"),
    Path("/Applications/Microsoft Word.app/Contents/Resources/MML2OMML.XSL"),
    Path("/usr/lib/libreoffice/share/xslt/import/ml/omml2mathml/MML2OMML.XSL"),
]


def find_mml2omml_xsl(explicit: Path | None = None) -> Path:
    if explicit is not None:
        if explicit.exists():
            return explicit.resolve()
        raise FileNotFoundError(f"MML2OMML.XSL not found: {explicit}")

    env = os.environ.get("PATENT_MATH_MML2OMML", "").strip()
    if env:
        p = Path(env)
        if p.exists():
            return p.resolve()

    for p in XSL_CANDIDATES:
        if p.exists():
            return p

    raise FileNotFoundError(
        "MML2OMML.XSL not found. Install Microsoft Office (or WPS with Office XSL), "
        "set PATENT_MATH_MML2OMML, or pass --xsl PATH."
    )


def parse_size_arg(value: str | None) -> int | None:
    if value is None:
        return None
    if value in CN_FONT_SIZE:
        return CN_FONT_SIZE[value]
    if value.isdigit():
        return int(value)
    raise ValueError(f"Unknown font size: {value}. Use 小四/四号/五号 or half-points like 24.")


def char_is_roman(ch: str) -> bool:
    if ch in ROMAN_MATH_CHARS:
        return True
    if ch.isdigit() or ch in "0123456789":
        return True
    if ord(ch) > 127:
        return True
    return False


def split_style_segments(text: str) -> list[tuple[str, bool]]:
    if not text:
        return []
    segments: list[tuple[str, bool]] = []
    i = 0
    while i < len(text):
        roman = char_is_roman(text[i])
        j = i + 1
        while j < len(text) and char_is_roman(text[j]) == roman:
            j += 1
        segments.append((text[i:j], not roman))  # True = italic
        i = j
    return segments


def get_paragraph_wrpr(p: etree._Element) -> etree._Element | None:
    for r in p.findall("w:r", NSMAP):
        wrpr = r.find("w:rPr", NSMAP)
        if wrpr is not None:
            return deepcopy(wrpr)
    ppr = p.find("w:pPr", NSMAP)
    if ppr is not None:
        wrpr = ppr.find("w:rPr", NSMAP)
        if wrpr is not None:
            return deepcopy(wrpr)
    return None


def set_wrpr_font_size(wrpr: etree._Element, half_points: int) -> None:
    sz = wrpr.find("w:sz", NSMAP)
    if sz is None:
        sz = etree.SubElement(wrpr, f"{{{WNS}}}sz")
    sz.set(f"{{{WNS}}}val", str(half_points))

    sz_cs = wrpr.find("w:szCs", NSMAP)
    if sz_cs is None:
        sz_cs = etree.SubElement(wrpr, f"{{{WNS}}}szCs")
    sz_cs.set(f"{{{WNS}}}val", str(half_points))


def ensure_times_new_roman(wrpr: etree._Element) -> None:
    fonts = wrpr.find("w:rFonts", NSMAP)
    if fonts is None:
        fonts = etree.SubElement(wrpr, f"{{{WNS}}}rFonts")
    fonts.set(f"{{{WNS}}}ascii", "Times New Roman")
    fonts.set(f"{{{WNS}}}hAnsi", "Times New Roman")
    fonts.set(f"{{{WNS}}}cs", "Times New Roman")
    if f"{{{WNS}}}eastAsia" in fonts.attrib or fonts.get(f"{{{WNS}}}eastAsia"):
        fonts.set(f"{{{WNS}}}eastAsia", "宋体")


def prepare_wrpr(base: etree._Element | None, override_sz: int | None) -> etree._Element:
    wrpr = deepcopy(base) if base is not None else etree.Element(f"{{{WNS}}}rPr")
    ensure_times_new_roman(wrpr)
    if override_sz is not None:
        set_wrpr_font_size(wrpr, override_sz)
    return wrpr


def set_mrpr_style(mr: etree._Element, italic: bool) -> None:
    mrpr = mr.find("m:rPr", NSMAP)
    if mrpr is None:
        mrpr = etree.SubElement(mr, f"{{{MNS}}}rPr")
    sty = mrpr.find("m:sty", NSMAP)
    if sty is None:
        sty = etree.SubElement(mrpr, f"{{{MNS}}}sty")
    sty.set(f"{{{MNS}}}val", "i" if italic else "p")


def apply_wrpr_to_omath(omath: etree._Element, wrpr: etree._Element) -> None:
    wrpr_copy = deepcopy(wrpr)
    for node in omath.iter():
        tag = node.tag
        if tag == f"{{{MNS}}}r" or tag.endswith("ctrlPr"):
            old = node.find("w:rPr", NSMAP)
            if old is not None:
                node.remove(old)
            node.append(deepcopy(wrpr_copy))


def fix_omath_delimiters(omath: etree._Element) -> None:
    for mr in list(omath.iter(f"{{{MNS}}}r")):
        mt = mr.find(f"{{{MNS}}}t")
        if mt is None or not mt.text:
            continue

        segments = split_style_segments(mt.text)
        if len(segments) <= 1:
            if segments:
                set_mrpr_style(mr, segments[0][1])
            continue

        parent = mr.getparent()
        if parent is None:
            continue
        idx = list(parent).index(mr)
        old_wrpr = mr.find(f"{{{WNS}}}rPr", NSMAP)
        parent.remove(mr)

        for seg_text, is_italic in segments:
            if not seg_text:
                continue
            new_mr = etree.Element(f"{{{MNS}}}r")
            set_mrpr_style(new_mr, is_italic)
            if old_wrpr is not None:
                new_mr.append(deepcopy(old_wrpr))
            new_mt = etree.SubElement(new_mr, f"{{{MNS}}}t")
            if seg_text.startswith(" ") or seg_text.endswith(" "):
                new_mt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            new_mt.text = seg_text
            parent.insert(idx, new_mr)
            idx += 1


def normalize_omath(omath: etree._Element, wrpr: etree._Element) -> None:
    fix_omath_delimiters(omath)
    apply_wrpr_to_omath(omath, wrpr)


def latex_to_omath(latex: str, xslt, wrpr: etree._Element) -> etree._Element:
    mathml = latex_to_mathml(latex.strip())
    mml_tree = etree.fromstring(mathml.encode("utf-8"))
    omml = xslt(mml_tree)
    root = omml.getroot() if hasattr(omml, "getroot") else omml
    if root.tag != f"{{{MNS}}}oMath":
        wrapper = etree.Element(f"{{{MNS}}}oMath")
        wrapper.append(root)
        root = wrapper
    normalize_omath(root, wrpr)
    return root


def paragraph_text(p: etree._Element) -> str:
    parts: list[str] = []
    for t in p.findall(".//w:t", NSMAP):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def is_display_paragraph(text: str, full_match: str, latex: str, double: bool) -> bool:
    trimmed = text.strip()
    if double:
        return trimmed == full_match
    return trimmed == full_match or trimmed == f"${latex}$"


def make_text_run(text: str, wrpr: etree._Element) -> etree._Element:
    r = etree.Element(f"{{{WNS}}}r")
    r.append(deepcopy(wrpr))
    t = etree.SubElement(r, f"{{{WNS}}}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def make_inline_math_run(omath: etree._Element, wrpr: etree._Element) -> etree._Element:
    r = etree.Element(f"{{{WNS}}}r")
    r.append(deepcopy(wrpr))
    r.append(omath)
    return r


def wrpr_sz_half_points(wrpr: etree._Element) -> int:
    sz = wrpr.find("w:sz", NSMAP)
    if sz is not None and sz.get(f"{{{WNS}}}val"):
        return int(sz.get(f"{{{WNS}}}val"))
    return 21


def char_width_dxa(sz_half_points: int) -> int:
    """Approximate one CJK character width in twips (for tab-by-char layout)."""
    return max(sz_half_points * 10, 180)


def read_page_layout(root: etree._Element) -> tuple[int, int, int]:
    """Return (page_width, left_margin, right_margin) in twips."""
    body = root.find("w:body", NSMAP)
    sect = body.find("w:sectPr", NSMAP) if body is not None else None
    if sect is None:
        return DEFAULT_PAGE_WIDTH_DXA, DEFAULT_MARGIN_DXA, DEFAULT_MARGIN_DXA

    pg_sz = sect.find("w:pgSz", NSMAP)
    pg_mar = sect.find("w:pgMar", NSMAP)
    width = int(pg_sz.get(f"{{{WNS}}}w", DEFAULT_PAGE_WIDTH_DXA)) if pg_sz is not None else DEFAULT_PAGE_WIDTH_DXA
    left = int(pg_mar.get(f"{{{WNS}}}left", DEFAULT_MARGIN_DXA)) if pg_mar is not None else DEFAULT_MARGIN_DXA
    right = int(pg_mar.get(f"{{{WNS}}}right", DEFAULT_MARGIN_DXA)) if pg_mar is not None else DEFAULT_MARGIN_DXA
    return width, left, right


def compute_tab_positions(
    page_width: int,
    left_margin: int,
    right_margin: int,
    wrpr: etree._Element,
    tab_cfg: TabLayoutConfig,
) -> tuple[int, int]:
    """Tab w:pos is relative to the paragraph left edge (not the physical page)."""
    content_width = page_width - left_margin - right_margin
    cw = char_width_dxa(wrpr_sz_half_points(wrpr))

    if tab_cfg.mode == "chars":
        center = int(tab_cfg.center_chars * cw)
        right = int(tab_cfg.right_chars * cw)
    else:
        center = int(content_width / 2)
        right = int(content_width)

    # Clamp to printable content width
    center = max(0, min(center, content_width))
    right = max(center, min(right, content_width))
    return center, right


def clear_block_formula_indent(ppr: etree._Element) -> None:
    """Block equations must not inherit body first-line / hanging indent."""
    ind = ppr.find("w:ind", NSMAP)
    if ind is None:
        return
    for attr in ("firstLine", "firstLineChars", "hanging", "hangingChars", "left", "leftChars"):
        key = f"{{{WNS}}}{attr}"
        if key in ind.attrib:
            del ind.attrib[key]
    if len(ind.attrib) == 0:
        ppr.remove(ind)


def set_paragraph_tab_stops(ppr: etree._Element, center_dxa: int, right_dxa: int) -> None:
    tabs = ppr.find("w:tabs", NSMAP)
    if tabs is None:
        tabs = etree.Element(f"{{{WNS}}}tabs")
        ppr.insert(0, tabs)
    else:
        for old in list(tabs.findall("w:tab", NSMAP)):
            tabs.remove(old)

    t_center = etree.SubElement(tabs, f"{{{WNS}}}tab")
    t_center.set(f"{{{WNS}}}val", "center")
    t_center.set(f"{{{WNS}}}pos", str(center_dxa))

    t_right = etree.SubElement(tabs, f"{{{WNS}}}tab")
    t_right.set(f"{{{WNS}}}val", "right")
    t_right.set(f"{{{WNS}}}pos", str(right_dxa))

    jc = ppr.find("w:jc", NSMAP)
    if jc is not None:
        ppr.remove(jc)


def make_tab_run(wrpr: etree._Element) -> etree._Element:
    r = etree.Element(f"{{{WNS}}}r")
    r.append(deepcopy(wrpr))
    etree.SubElement(r, f"{{{WNS}}}tab")
    return r


def make_display_paragraph(
    omath: etree._Element,
    eq_num: int | None,
    source_ppr: etree._Element | None,
    wrpr: etree._Element,
    tab_cfg: TabLayoutConfig,
    page_width: int,
    left_margin: int,
    right_margin: int,
) -> etree._Element:
    p = etree.Element(f"{{{WNS}}}p")

    if source_ppr is not None:
        ppr = deepcopy(source_ppr)
    else:
        ppr = etree.SubElement(p, f"{{{WNS}}}pPr")

    center_dxa, right_dxa = compute_tab_positions(
        page_width, left_margin, right_margin, wrpr, tab_cfg
    )
    set_paragraph_tab_stops(ppr, center_dxa, right_dxa)
    clear_block_formula_indent(ppr)

    spacing = ppr.find("w:spacing", NSMAP)
    if spacing is None:
        spacing = etree.SubElement(ppr, f"{{{WNS}}}spacing")
    spacing.set(f"{{{WNS}}}before", "120")
    spacing.set(f"{{{WNS}}}after", "120")

    ppr_rpr = ppr.find("w:rPr", NSMAP)
    if ppr_rpr is None:
        ppr.insert(0, deepcopy(wrpr))
    else:
        ensure_times_new_roman(ppr_rpr)
        sz_val = wrpr.find("w:sz", NSMAP)
        if sz_val is not None and sz_val.get(f"{{{WNS}}}val"):
            set_wrpr_font_size(ppr_rpr, int(sz_val.get(f"{{{WNS}}}val")))

    p.append(ppr)

    # \t + formula + \t + (n)
    p.append(make_tab_run(wrpr))

    omath_para = etree.SubElement(p, f"{{{MNS}}}oMathPara")
    omath_para.append(omath)

    if eq_num is not None:
        p.append(make_tab_run(wrpr))
        p.append(make_text_run(f"({eq_num})", wrpr))
    return p


def split_paragraph(text: str) -> list[tuple[str, str | None]]:
    pieces: list[tuple[str, str | None]] = []
    last = 0
    for m in FORMULA_RE.finditer(text):
        if m.start() > last:
            pieces.append(("text", text[last : m.start()]))
        latex = m.group(1) if m.group(1) is not None else m.group(2)
        pieces.append(("math", latex))
        last = m.end()
    if last < len(text):
        pieces.append(("text", text[last:]))
    return pieces


def process_document_xml(
    xml_bytes: bytes,
    xslt,
    override_sz: int | None,
    tab_cfg: TabLayoutConfig,
) -> tuple[bytes, int, int, list[str]]:
    root = etree.fromstring(xml_bytes)
    body = root.find("w:body", NSMAP)
    if body is None:
        raise ValueError("Invalid docx: missing w:body")

    page_width, left_margin, right_margin = read_page_layout(root)

    ok = 0
    fail = 0
    failed_formulas: list[str] = []
    display_num = 0
    new_children: list[etree._Element] = []

    for p in list(body.findall("w:p", NSMAP)):
        text = paragraph_text(p)
        if "$" not in text:
            new_children.append(p)
            body.remove(p)
            continue

        pieces = split_paragraph(text)
        has_math = any(k == "math" for k, _ in pieces)
        if not has_math:
            new_children.append(p)
            body.remove(p)
            continue

        base_wrpr = get_paragraph_wrpr(p)
        wrpr = prepare_wrpr(base_wrpr, override_sz)
        source_ppr = p.find("w:pPr", NSMAP)

        only_math = len(pieces) == 1 and pieces[0][0] == "math"
        trimmed = text.strip()
        m = FORMULA_RE.search(trimmed)
        is_display = False
        if m and only_math:
            latex = m.group(1) or m.group(2)
            is_display = is_display_paragraph(text, m.group(0), latex, m.group(1) is not None)

        if is_display:
            latex = pieces[0][1]
            try:
                omath = latex_to_omath(latex, xslt, wrpr)
                display_num += 1
                src_ppr = deepcopy(source_ppr) if source_ppr is not None else None
                new_p = make_display_paragraph(
                    omath,
                    display_num,
                    src_ppr,
                    wrpr,
                    tab_cfg,
                    page_width,
                    left_margin,
                    right_margin,
                )
                new_children.append(new_p)
                ok += 1
            except Exception:
                new_children.append(p)
                fail += 1
                if latex:
                    failed_formulas.append(latex)
            body.remove(p)
            continue

        rebuilt: list[tuple[str, etree._Element | None]] = []
        for kind, value in pieces:
            if kind == "text":
                if value:
                    rebuilt.append(("text", make_text_run(value, wrpr)))
            else:
                try:
                    omath = latex_to_omath(value, xslt, wrpr)
                    rebuilt.append(("math", make_inline_math_run(omath, wrpr)))
                    ok += 1
                except Exception:
                    rebuilt.append(("text", make_text_run(f"${value}$", wrpr)))
                    fail += 1
                    failed_formulas.append(value)

        ppr = source_ppr
        for child in list(p):
            if child is not ppr:
                p.remove(child)

        if ppr is not None:
            ppr_wrpr = ppr.find("w:rPr", NSMAP)
            if ppr_wrpr is None:
                ppr.append(deepcopy(wrpr))
            else:
                ensure_times_new_roman(ppr_wrpr)
                sz = wrpr.find("w:sz", NSMAP)
                if sz is not None and sz.get(f"{{{WNS}}}val"):
                    set_wrpr_font_size(ppr_wrpr, int(sz.get(f"{{{WNS}}}val")))

        for _, node in rebuilt:
            if node is not None:
                p.append(node)
        new_children.append(p)
        body.remove(p)

    for p in new_children:
        body.append(p)

    return (
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"),
        ok,
        fail,
        failed_formulas,
    )


def convert_docx(
    input_path: Path,
    output_path: Path,
    override_sz: int | None,
    tab_cfg: TabLayoutConfig,
    xsl_path: Path | None = None,
) -> tuple[int, int, list[str]]:
    xsl_path = find_mml2omml_xsl(xsl_path)
    xslt = etree.XSLT(etree.parse(str(xsl_path)))

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        with zipfile.ZipFile(input_path, "r") as zin:
            zin.extractall(tmp_dir)

        doc_xml = tmp_dir / "word" / "document.xml"
        new_xml, ok, fail, failed_formulas = process_document_xml(
            doc_xml.read_bytes(), xslt, override_sz, tab_cfg
        )
        doc_xml.write_bytes(new_xml)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for file in tmp_dir.rglob("*"):
                if file.is_file():
                    zout.write(file, file.relative_to(tmp_dir).as_posix())

    return ok, fail, failed_formulas


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="patent-math-tool",
        description="Convert $...$ / $$...$$ LaTeX in .docx to native Word/WPS OMML equations.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("input", help="Input .docx")
    parser.add_argument("output", nargs="?", help="Output .docx (default: <input>_converted.docx)")
    parser.add_argument(
        "--size",
        dest="size",
        default=None,
        help="Override font size: 小四, 四号, 五号, or half-points (e.g. 24)",
    )
    parser.add_argument(
        "--xsl",
        dest="xsl",
        default=None,
        help="Path to MML2OMML.XSL (default: Office install or PATENT_MATH_MML2OMML)",
    )
    parser.add_argument(
        "--tab-mode",
        choices=("page", "chars"),
        default="page",
        help="Block formula tab stops: page=paper-adaptive (default), chars=character count",
    )
    parser.add_argument(
        "--tab-center-chars",
        type=float,
        default=21.0,
        help="Center tab at N chars from left margin (used when --tab-mode chars)",
    )
    parser.add_argument(
        "--tab-right-chars",
        type=float,
        default=44.0,
        help="Right tab at N chars from left margin (used when --tab-mode chars)",
    )
    parser.add_argument("-q", "--quiet", action="store_true", help="Only print errors and output path")
    args = parser.parse_args(argv)

    input_path = Path(args.input).resolve()
    if args.output:
        output_path = Path(args.output).resolve()
    else:
        output_path = input_path.with_name(input_path.stem + "_converted.docx")

    if not input_path.exists():
        print(f"File not found: {input_path}", file=sys.stderr)
        return 1

    try:
        override_sz = parse_size_arg(args.size)
    except ValueError as e:
        print(str(e), file=sys.stderr)
        return 1

    tab_cfg = TabLayoutConfig(
        mode=args.tab_mode,
        center_chars=args.tab_center_chars,
        right_chars=args.tab_right_chars,
    )

    xsl = Path(args.xsl).resolve() if args.xsl else None

    try:
        ok, fail, _ = convert_docx(input_path, output_path, override_sz, tab_cfg, xsl)
    except FileNotFoundError as e:
        print(str(e), file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Conversion failed: {e}", file=sys.stderr)
        return 1

    if not args.quiet:
        print(f"Done: {ok} converted, {fail} kept as original text")
        if override_sz:
            print(f"Font size (half-points): {override_sz}")
        if tab_cfg.mode == "page":
            print("Block formula tabs: content-area adaptive (center=content/2, right=content width)")
        else:
            print(
                f"Block formula tabs: center={tab_cfg.center_chars} chars, "
                f"right={tab_cfg.right_chars} chars"
            )
    print(f"Output: {output_path}")
    return 0 if fail == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
