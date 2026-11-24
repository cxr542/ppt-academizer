"""A cover with a small diagram, a divider with a top-left title, and a card header bar.

Regression for a deck whose cover (two version boxes and an arrow) was rebuilt as a content slide,
whose two divider slides (one title line each) became empty content slides, and whose slide of
two cards lost both blue header bars and their labels.
"""

from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt

from scripts import shape_migrate_lib as bac
from scripts.build_cmp_academy import build_slide_plan
from scripts.llm_classifier import analyze_slide_with_llm


def _prs():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs


def _text(slide, left, top, width, height, text, size):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.text = text
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(size)
    return box


def _filled(slide, shape, left, top, width, height):
    sh = slide.shapes.add_shape(shape, Inches(left), Inches(top), Inches(width), Inches(height))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor(0x14, 0x21, 0x3D)
    return sh


def _diagram_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _text(s, 0.8, 1.0, 6, 0.3, "WHAT'S NEW", 14)
    _text(s, 0.8, 1.35, 11, 0.9, "OKESTRO CMP v3.0.5", 48)
    _text(s, 0.8, 2.4, 11, 0.5, "업무 라벨 기준으로 한 콘솔에서 조회", 20)
    _filled(s, MSO_SHAPE.RECTANGLE, 0.8, 3.7, 4.8, 1.3)
    _text(s, 0.8, 3.8, 4.8, 0.3, "v3.0.4", 14)
    _text(s, 0.8, 4.2, 4.8, 0.55, "테넌트 연결", 24)
    _filled(s, MSO_SHAPE.RIGHT_ARROW, 5.95, 4.1, 1.5, 0.5)
    _filled(s, MSO_SHAPE.RECTANGLE, 7.8, 3.7, 4.8, 1.3)
    _text(s, 7.8, 3.8, 4.8, 0.3, "v3.0.5", 14)
    _text(s, 7.8, 4.2, 4.8, 0.55, "업무 라벨 중심", 24)
    _text(s, 0.8, 6.05, 11, 0.4, "Business-driven Cloud Governance", 16)
    return s


def test_cover_with_a_small_diagram_is_planned_as_cover() -> None:
    prs = _prs()
    _diagram_cover(prs)
    info = analyze_slide_with_llm(prs.slides[0])
    assert info["body_heavy"] and not info["cover_meta_only"]
    assert info["cover_diagram"]
    assert build_slide_plan(prs)[0] == (0, "cover")


def test_cover_with_a_long_paragraph_beside_boxes_is_still_content() -> None:
    prs = _prs()
    s = _diagram_cover(prs)
    _text(s, 0.8, 5.3, 11, 0.6, "긴 본문 " * 40, 14)
    assert not analyze_slide_with_llm(prs.slides[0])["cover_diagram"]
    assert build_slide_plan(prs)[0] == (0, "content")


def test_filled_cover_keeps_the_diagram_inside_the_card() -> None:
    from scripts import build_cmp_academy as bca

    prs = _prs()
    src = _diagram_cover(prs)
    dst = prs.slides.add_slide(prs.slide_layouts[6])
    box = dst.shapes.add_textbox(Inches(0.57), Inches(2.08), Inches(10.8), Inches(1.5))
    box.text_frame.text = "x"
    info = analyze_slide_with_llm(src)
    bca._fill_cover_with_meta(dst, src, info, bca.config_for_kind(bca.GENERIC_KIND))

    labels = [s for s in dst.shapes if s.has_text_frame and s.text_frame.text in ("v3.0.4", "v3.0.5")]
    assert len(labels) == 2 and labels[0].left < labels[1].left
    boxes = [s for s in dst.shapes if s.shape_type == 1 and not s.text_frame.text.strip()]
    assert len(boxes) == 3  # two boxes and the arrow
    for sh in boxes:
        assert sh.left >= box.left and sh.left + sh.width <= box.left + box.width
        assert sh.top >= box.top + Inches(1.4)  # under the text
        assert sh.top + sh.height <= box.top + box.height
    # the diagram's labels are not folded into the subtitle
    assert [p.text for p in box.text_frame.paragraphs][-1].startswith("업무 라벨 기준")


def _dark_background(slide, color="051247"):
    from lxml import etree

    ns = (
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    )
    bg = etree.fromstring(
        f'<p:bg {ns}><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="{color}"/></a:gs>'
        f'<a:gs pos="100000"><a:schemeClr val="accent4"/></a:gs></a:gsLst></a:gradFill>'
        f"<a:effectLst/></p:bgPr></p:bg>"
    )
    slide._element.cSld.insert(0, bg)


def test_a_dark_slide_with_only_a_top_left_title_is_a_divider() -> None:
    prs = _prs()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_background(s)
    _text(s, 0.75, 0.55, 5, 0.8, "3.0.4 vs 3.0.5", 32)
    info = analyze_slide_with_llm(s)
    assert info["layout"] == "section"
    assert info["title"] == "3.0.4 vs 3.0.5"


def test_a_white_slide_with_only_a_header_title_stays_content() -> None:
    prs = _prs()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _text(s, 0.75, 0.55, 5, 0.8, "3.0.4 vs 3.0.5", 32)
    assert analyze_slide_with_llm(s)["layout"] != "section"


def test_a_title_over_a_picture_is_not_a_divider() -> None:
    import io

    from PIL import Image

    prs = _prs()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_background(s)
    _text(s, 0.75, 0.55, 5, 0.8, "구조", 32)
    buf = io.BytesIO()
    Image.new("RGB", (20, 20), (0, 0, 200)).save(buf, "PNG")
    buf.seek(0)
    s.shapes.add_picture(buf, Inches(1), Inches(2), Inches(4), Inches(3))
    assert analyze_slide_with_llm(s)["layout"] != "section"


def test_card_header_bar_and_label_under_the_title_row_are_kept() -> None:
    prs = _prs()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _text(s, 0.55, 0.3, 8, 0.6, "정리", 30)
    top_bar = _filled(s, MSO_SHAPE.RECTANGLE, 0, 0, 13.33, 0.14)
    card = _filled(s, MSO_SHAPE.RECTANGLE, 0.55, 1.1, 5.98, 3.95)
    bar = _filled(s, MSO_SHAPE.RECTANGLE, 0.55, 1.1, 5.98, 0.5)
    label = _text(s, 0.75, 1.1, 5.6, 0.5, "3.0.4 → 3.0.5", 14)

    assert bac.is_body_shape(card, "정리", None)
    assert bac.is_body_shape(bar, "정리", None)
    assert bac.is_body_shape(label, "정리", None)
    assert not bac.is_body_shape(top_bar, "정리", None)  # the source's own top accent bar
    title = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text == "정리"][0]
    assert not bac.is_body_shape(title, "정리", None)
