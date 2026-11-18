"""A group whose child space is scaled keeps its size when flattened, and is not copied twice.

Regression for a deck whose slide 7 held a 2.4in icon group (children 3.6in in child space):
the copy came out 50% too big and the raster sync, finding no matching box, copied it again.
"""

from __future__ import annotations

import io

from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches

from scripts.build_cmp_academy import copy_body_shapes_filtered, sync_raster_diagram_assets


def _prs():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs


def _group_slide(prs, ext_in=2.4, child_in=3.6):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    grp = slide.shapes.add_group_shape()
    buf = io.BytesIO()
    Image.new("RGB", (30, 30), (0, 180, 0)).save(buf, "PNG")
    buf.seek(0)
    grp.shapes.add_picture(buf, Inches(4.9), Inches(4.3), Inches(child_in), Inches(child_in))
    label = grp.shapes.add_textbox(Inches(5.3), Inches(4.6), Inches(1.5), Inches(1.0))
    label.text_frame.text = "VIOLA"
    xfrm = grp._element.grpSpPr.get_or_add_xfrm()
    xfrm.off.x, xfrm.off.y = Inches(5.5), Inches(4.4)
    xfrm.ext.cx = xfrm.ext.cy = Inches(ext_in)
    xfrm.chOff.x, xfrm.chOff.y = Inches(4.9), Inches(4.3)
    xfrm.chExt.cx = xfrm.chExt.cy = Inches(child_in)
    return slide


def test_scaled_group_children_are_scaled_to_the_group() -> None:
    prs = _prs()
    src = _group_slide(prs)
    dst = prs.slides.add_slide(prs.slide_layouts[6])

    copy_body_shapes_filtered(src, dst, None, None, prs.slide_width, prs.slide_height)

    pic = [s for s in dst.shapes if s.shape_type == 13][0]
    assert abs(Emu(pic.width).inches - 2.4) < 0.02
    assert abs(Emu(pic.height).inches - 2.4) < 0.02
    assert abs(Emu(pic.left).inches - 5.5) < 0.02
    label = [s for s in dst.shapes if s.has_text_frame and s.text_frame.text == "VIOLA"][0]
    assert abs(Emu(label.width).inches - 1.0) < 0.02  # 1.5in * (2.4 / 3.6)


def test_raster_sync_does_not_copy_the_group_a_second_time() -> None:
    prs = _prs()
    src = _group_slide(prs)
    dst = prs.slides.add_slide(prs.slide_layouts[6])
    copy_body_shapes_filtered(src, dst, None, None, prs.slide_width, prs.slide_height)
    before = len(dst.shapes)

    sync_raster_diagram_assets(
        src, dst, None, None, prs.slide_width, prs.slide_height, skip_connectors=False
    )

    assert len(dst.shapes) == before == 2


def test_an_unscaled_group_is_flattened_as_before() -> None:
    prs = _prs()
    src = _group_slide(prs, ext_in=3.6, child_in=3.6)
    dst = prs.slides.add_slide(prs.slide_layouts[6])
    copy_body_shapes_filtered(src, dst, None, None, prs.slide_width, prs.slide_height)
    pic = [s for s in dst.shapes if s.shape_type == 13][0]
    assert abs(Emu(pic.width).inches - 3.6) < 0.02
