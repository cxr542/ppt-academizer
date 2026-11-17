"""A cover with an eyebrow, a subtitle and a source line is still a cover, and keeps all lines."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Emu, Inches, Pt

from scripts.build_cmp_academy import build_slide_plan
from scripts.llm_classifier import analyze_slide_with_llm
from scripts.slide_classifier_v2 import cover_lines


def _prs():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs


def _text(slide, top, text, size, left=0.9, width=11):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(0.6))
    box.text_frame.text = text
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(size)
    return box


def _cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _text(s, 2.4, "영업·프리세일즈 교육", 14)
    _text(s, 3.0, "VIOLA CI/CD vs TROMBONE", 48)
    _text(s, 4.3, "통제력과 자율성: CI/CD 기능 차이와 고객 제안 기준", 24)
    _text(s, 6.6, "자료 기준: VIOLA CI/CD 3.0.6, TROMBONE 3.0.5.1 매뉴얼·기능명세서 | 2026.11.17", 12)
    return s


def test_cover_with_meta_lines_is_planned_as_cover() -> None:
    prs = _prs()
    _cover(prs)
    info = analyze_slide_with_llm(prs.slides[0])
    assert info["body_heavy"] and info["cover_meta_only"]
    assert build_slide_plan(prs)[0] == (0, "cover")


def test_cover_lines_put_the_title_first_and_keep_every_line() -> None:
    prs = _prs()
    _cover(prs)
    lines = cover_lines(analyze_slide_with_llm(prs.slides[0])["blocks"])
    assert lines[0] == "VIOLA CI/CD vs TROMBONE"
    assert len(lines) == 4
    assert any(l.startswith("자료 기준") for l in lines)
    assert "영업·프리세일즈 교육" in lines


def test_a_slide_with_a_long_paragraph_is_still_content() -> None:
    prs = _prs()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _text(s, 3.0, "개요", 40)
    _text(s, 4.2, "긴 본문 " * 30, 14)
    assert build_slide_plan(prs)[0] == (0, "content")
