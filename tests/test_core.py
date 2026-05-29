"""Core logic tests — no MML2OMML.XSL required."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from convert_latex_docx import (  # noqa: E402
    CN_FONT_SIZE,
    TabLayoutConfig,
    is_display_paragraph,
    parse_size_arg,
    split_paragraph,
)


class TestSplitParagraph:
    def test_inline_and_display_markers(self):
        pieces = split_paragraph("设 $x=1$ 且 $$y=2$$ 成立")
        kinds = [k for k, _ in pieces]
        assert kinds == ["text", "math", "text", "math", "text"]
        assert pieces[1][1] == "x=1"
        assert pieces[3][1] == "y=2"

    def test_single_inline(self):
        pieces = split_paragraph("能量 $E=mc^2$ 守恒")
        assert len(pieces) == 3
        assert pieces[1] == ("math", "E=mc^2")


class TestDisplayDetection:
    def test_double_dollar_full_paragraph(self):
        assert is_display_paragraph("$$a+b$$", "$$a+b$$", "a+b", True)

    def test_single_dollar_own_line(self):
        assert is_display_paragraph("$x^2$", "$x^2$", "x^2", False)

    def test_inline_in_sentence(self):
        assert not is_display_paragraph("见式 $x^2$ 可知", "$x^2$", "x^2", False)


class TestParseSize:
    def test_chinese_size(self):
        assert parse_size_arg("小四") == CN_FONT_SIZE["小四"]

    def test_half_points(self):
        assert parse_size_arg("24") == 24

    def test_invalid(self):
        with pytest.raises(ValueError):
            parse_size_arg("超大")


class TestExamplesExist:
    def test_demo_before_present(self):
        p = ROOT / "examples" / "demo_before.docx"
        assert p.is_file() and p.stat().st_size > 500
