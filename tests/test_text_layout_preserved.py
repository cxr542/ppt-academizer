"""Copying a text shape keeps its box anchor, paragraph alignment, bullets and corner radius."""

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

import scripts.academy_elements_style as style
import scripts.shape_migrate_lib as lib

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _src_slide():
    prs = Presentation()
    return prs, prs.slides.add_slide(prs.slide_layouts[6])


def _set_ppr(paragraph, xml_children: str, attrs: str = ""):
    ppr = etree.fromstring(f'<a:pPr xmlns:a="{A}" {attrs}>{xml_children}</a:pPr>')
    paragraph._p.insert(0, ppr)


def test_centered_label_keeps_anchor_insets_and_alignment():
    prs, src = _src_slide()
    box = src.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    box.text_frame.text = "Business Label"
    bpr = box.text_frame._txBody.find(f"{{{A}}}bodyPr")
    bpr.set("anchor", "ctr")
    bpr.set("lIns", "0")
    _set_ppr(box.text_frame.paragraphs[0], "<a:buNone/>", 'algn="ctr"')
    dst = prs.slides.add_slide(prs.slide_layouts[6])
    lib.copy_textbox_native(dst, box)
    out = dst.shapes[-1].text_frame._txBody
    assert out.find(f"{{{A}}}bodyPr").get("anchor") == "ctr"
    assert out.find(f"{{{A}}}bodyPr").get("lIns") == "0"
    assert out.find(f".//{{{A}}}pPr").get("algn") == "ctr"


def test_short_bullet_lines_keep_their_bullets():
    prs, src = _src_slide()
    box = src.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    box.text_frame.text = "3D 랙\n개인화 Home"  # short lines: the length guess would never bullet them
    for p in box.text_frame.paragraphs:
        _set_ppr(p, '<a:buChar char="&#x2022;"/>', 'marL="139700" indent="-139700"')
    dst = prs.slides.add_slide(prs.slide_layouts[6])
    lib.copy_textbox_native(dst, box)
    ppr = dst.shapes[-1].text_frame._txBody.findall(f".//{{{A}}}pPr")
    assert len(ppr) == 2
    for p in ppr:
        assert p.find(f"{{{A}}}buChar") is not None
        assert p.get("marL") == style.BULLET_MARGIN  # academy bullet look
        assert p.find(f"{{{A}}}buClr/{{{A}}}schemeClr").get("val") == "accent6"


def test_rounded_card_keeps_its_corner_radius():
    prs, src = _src_slide()
    card = src.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
    card.adjustments[0] = 0.05
    dst = prs.slides.add_slide(prs.slide_layouts[6])
    lib.copy_autoshape_native(dst, card, card.left, card.top)
    assert abs(dst.shapes[-1].adjustments[0] - 0.05) < 1e-6


def test_an_explicit_no_bullet_is_not_overridden_by_the_list_guess():
    prs, src = _src_slide()
    box = src.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    box.text_frame.text = "첫 번째 긴 설명 문장입니다.\n두 번째 긴 설명 문장입니다."
    for p in box.text_frame.paragraphs:
        _set_ppr(p, "<a:buNone/>")
    assert style._bulletize(box) is False
    assert all(p.find(f".//{{{A}}}buChar") is None for p in box.text_frame._txBody.findall(f"{{{A}}}p"))
