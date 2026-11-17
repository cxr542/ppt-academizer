"""A flowchart keeps its arrows, and a card's bullet list is never promoted to the governing line.

Regression for a deck whose slide 6 ("고객 상황별 제안 기준") lost all eight connectors and whose
slide 8 ("정리") moved the card's three bullets into the one-line governing placeholder.
"""

from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from scripts.build_cmp_academy import copy_body_shapes_filtered, is_radial_diagram_slide
from scripts.slide_classifier_v2 import analyze_slide_layout


def _prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs


def _text(slide, left, top, width, height, text, size):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.text = text
    for para in box.text_frame.paragraphs:
        for run in para.runs:
            run.font.size = Pt(size)
    return box


def _arrow(slide, x1, y1, x2, y2):
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    conn.line.width = Pt(1.5)
    conn.line.color.rgb = RGBColor(0x5F, 0x6B, 0x86)
    ln = conn._element.spPr.find(qn("a:ln"))
    tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)
    return conn


def _flowchart_slide(prs):
    """Five question boxes in a row with 'no' arrows between them and 'yes' arrows below."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _text(slide, 0.9, 0.9, 8, 0.5, "고객 상황별 제안 기준", 30)
    xs = [0.9, 3.5, 6.1, 8.7, 11.3]
    for i, x in enumerate(xs, 1):
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.8), Inches(1.9), Inches(1.5)
        )
        box.text_frame.text = f"Q{i}\n질문 {i}번을\n확인하는가?"
    for i, x in enumerate(xs[:-1]):
        _arrow(slide, x + 1.9, 2.55, xs[i + 1], 2.55)
        _text(slide, x + 1.9, 2.1, 0.7, 0.3, "아니오", 12)
        _arrow(slide, x + 0.95, 3.3, x + 0.95, 3.8)
        _text(slide, x + 1.0, 3.35, 0.5, 0.3, "예", 12)
    _text(slide, 1.2, 4.3, 6.5, 1.3, "TROMBONE 제안\n티켓·워크플로우로 변경 이력과 승인을 통제", 14)
    _text(slide, 8.9, 4.3, 3.4, 1.5, "VIOLA CI/CD 제안\n표준 파이프라인으로 빠르게 시작", 14)
    return slide


def test_flowchart_is_taken_for_a_radial_slide() -> None:
    """The premise of the bug: this layout really does trip the radial heuristic."""
    prs = _prs()
    slide = _flowchart_slide(prs)
    assert is_radial_diagram_slide(slide, prs.slide_width, prs.slide_height)


def test_flowchart_arrows_are_copied_with_their_heads() -> None:
    prs = _prs()
    src = _flowchart_slide(prs)
    dst = prs.slides.add_slide(prs.slide_layouts[6])

    copy_body_shapes_filtered(
        src, dst, "고객 상황별 제안 기준", None, prs.slide_width, prs.slide_height,
        skip_connectors=False,
    )

    def lines(slide):
        return [
            s for s in slide.shapes
            if s.shape_type == MSO_SHAPE_TYPE.LINE or s._element.tag.endswith("}cxnSp")
        ]

    assert len(lines(src)) == 8
    copied = lines(dst)
    assert len(copied) == 8
    for conn in copied:
        tail = conn._element.find(".//" + qn("a:tailEnd"))
        assert tail is not None and tail.get("type") == "triangle"


def test_flipped_arrow_keeps_its_flip() -> None:
    prs = _prs()
    src = prs.slides.add_slide(prs.slide_layouts[6])
    conn = _arrow(src, 3, 3, 1, 1)  # drawn right-to-left, bottom-to-top => flipped source
    dst = prs.slides.add_slide(prs.slide_layouts[6])

    copy_body_shapes_filtered(src, dst, None, None, prs.slide_width, prs.slide_height)

    xfrm_src = conn._element.find(".//" + qn("a:xfrm"))
    out = [s for s in dst.shapes if s._element.tag.endswith("}cxnSp")][0]
    xfrm_dst = out._element.find(".//" + qn("a:xfrm"))
    for attr in ("flipH", "flipV"):
        assert xfrm_dst.get(attr) == xfrm_src.get(attr)


def test_card_bullets_are_not_the_governing_message() -> None:
    prs = _prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _text(slide, 0.7, 0.7, 8, 0.5, "정리", 30)
    _text(slide, 0.9, 2.0, 5, 0.4, "핵심 정리", 20)
    _text(
        slide, 1.0, 2.5, 5, 1.7,
        "VIOLA CI/CD는 K8s 앱의 빌드·테스트·배포를 표준 파이프라인으로 구성\n"
        "TROMBONE은 티켓·워크플로우 기반 변경관리, 결재와 이력 추적\n"
        "선택 기준은 배포 환경, 승인 절차, 감사 요구",
        15,
    )
    info = analyze_slide_layout(slide)
    assert info["title"] == "정리"
    assert not info["governing"]


def test_a_short_sentence_under_the_title_is_still_governing() -> None:
    prs = _prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _text(slide, 0.7, 0.7, 8, 0.5, "분기 실적", 30)
    _text(slide, 0.7, 1.4, 8, 0.4, "매출은 늘었지만 이익률은 소폭 하락했습니다.", 16)
    _text(slide, 0.7, 3.0, 8, 1.0, "세부 내용", 14)
    info = analyze_slide_layout(slide)
    assert info["governing"] == "매출은 늘었지만 이익률은 소폭 하락했습니다."
