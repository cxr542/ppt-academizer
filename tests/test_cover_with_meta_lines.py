"""A cover with an eyebrow, a subtitle and a source line is still a cover, and keeps all lines."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Emu, Inches, Pt

from scripts.build_cmp_academy import build_slide_plan
from scripts.llm_classifier import analyze_slide_with_llm
from scripts.slide_classifier_v2 import split_cover_roles


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


def test_roles_split_into_eyebrow_title_subtitle_footer() -> None:
    prs = _prs()
    _cover(prs)
    roles = split_cover_roles(
        analyze_slide_with_llm(prs.slides[0])["blocks"], prs.slide_height / 914400
    )
    assert roles["title"] == "VIOLA CI/CD vs TROMBONE"
    assert roles["eyebrow"] == ["영업·프리세일즈 교육"]
    assert roles["subtitle"] == ["통제력과 자율성: CI/CD 기능 차이와 고객 제안 기준"]
    assert len(roles["footer"]) == 1 and roles["footer"][0].startswith("자료 기준")


def test_filled_cover_places_eyebrow_icons_and_source_line() -> None:
    from scripts import build_cmp_academy as bca
    from scripts.academy_deck_build_lib import _non_placeholder_text_shapes

    prs = _prs()
    src = _cover(prs)
    icon = Inches(0.85)
    import io
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (20, 20), (0, 200, 0)).save(buf, "PNG")
    for i in range(2):
        buf.seek(0)
        src.shapes.add_picture(buf, Inches(0.9 + i * 1.05), Inches(5.2), icon, icon)

    dst = prs.slides.add_slide(prs.slide_layouts[6])
    box = dst.shapes.add_textbox(Inches(0.57), Inches(2.08), Inches(10.8), Inches(1.5))
    box.text_frame.text = "x"
    info = analyze_slide_with_llm(src)
    bca._fill_cover_with_meta(dst, src, info, bca.config_for_kind(bca.GENERIC_KIND))

    texts = [p.text for p in box.text_frame.paragraphs]
    assert texts[0] == "영업·프리세일즈 교육"
    assert texts[1] == "VIOLA CI/CD vs TROMBONE"
    assert texts[2].startswith("통제력과")
    pics = [s for s in dst.shapes if s.shape_type == 13]
    assert len(pics) == 2 and pics[0].left < pics[1].left
    assert all(p.top + p.height <= box.top + box.height for p in pics)
    assert all(p.top >= box.top + Inches(1.4) for p in pics)  # under the text, not over it
    note = [s for s in dst.shapes if s.has_text_frame and s.text_frame.text.startswith("자료 기준")]
    assert note and note[0].top >= box.top + box.height


def test_a_slide_with_a_long_paragraph_is_still_content() -> None:
    prs = _prs()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _text(s, 3.0, "개요", 40)
    _text(s, 4.2, "긴 본문 " * 30, 14)
    assert build_slide_plan(prs)[0] == (0, "content")
