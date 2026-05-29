#!/usr/bin/env python3
"""Build examples/demo_before.docx and run converter for demo_after.docx."""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from convert_latex_docx import convert_docx, TabLayoutConfig  # noqa: E402

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""


def esc(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def paragraph(text: str) -> str:
    return f"<w:p><w:r><w:t xml:space=\"preserve\">{esc(text)}</w:t></w:r></w:p>"


def build_document_xml() -> str:
    lines = [
        "专利文档公式示例（脱敏演示）",
        "",
        "行内公式：根据关系式 $E = mc^2$，可得能量密度。",
        "",
        "另一行内：设集合 $S = \\{1, 2, 3\\}$，则 $|S| = 3$。",
        "",
        "$$\\int_0^1 x^2 \\, dx = \\frac{1}{3}$$",
        "",
        "$$\\sigma = \\sqrt{\\frac{1}{n}\\sum_{i=1}^{n}(x_i - \\bar{x})^2}$$",
        "",
        "块级（整段仅一个 $...$ 也会居中编号）：",
        "",
        "$V_{audit} = \\{v \\in V \\mid f(v) = 1\\}$",
    ]
    body = "".join(paragraph(t) for t in lines)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"""


def write_demo_before(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", build_document_xml())


def main() -> None:
    examples = ROOT / "examples"
    before = examples / "demo_before.docx"
    after = examples / "demo_after.docx"

    write_demo_before(before)
    print(f"Created: {before}")

    ok, fail, _ = convert_docx(before, after, override_sz=None, tab_cfg=TabLayoutConfig())
    print(f"Converted: {ok} ok, {fail} fail -> {after}")


if __name__ == "__main__":
    main()
