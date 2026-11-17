"""Body shape colors move onto the academy palette; structure, text and pictures do not."""

from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches

from scripts.academy_palette import (
    DEFAULT_PALETTE,
    map_color,
    map_slide_colors,
    palette_from_pack,
)

PALETTE_VALUES = {v for v in DEFAULT_PALETTE.values()}


def _luma(h: str) -> float:
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


def test_palette_colors_and_white_are_unchanged() -> None:
    for hex_v in [*PALETTE_VALUES, "FFFFFF", "FBFBFD", "F5F6F9"]:
        assert map_color(hex_v) == hex_v


def test_source_navy_slate_and_orange_move_to_the_theme() -> None:
    assert map_color("14213D") in PALETTE_VALUES  # dark navy
    assert map_color("3F4A63") == "44546A"  # slate -> dk2
    assert map_color("5F6B86") == "78788C"  # gray -> accent6
    assert map_color("B34A0A") == "F15928"  # orange -> accent5


def test_light_tint_keeps_its_lightness() -> None:
    mapped = map_color("E8ECF4")
    assert mapped not in PALETTE_VALUES
    assert abs(_luma(mapped) - _luma("E8ECF4")) < 0.03


def test_a_distinct_hue_is_kept() -> None:
    """Status colors (a green "OK") have no stand-in in the academy palette."""
    assert map_color("2E9E5B") == "2E9E5B"
    assert map_color("7B1FA2") == "7B1FA2"


def test_palette_from_pack_overrides_and_falls_back() -> None:
    pal = palette_from_pack({"theme": {"accent1": "#112233", "accent2": "bad"}})
    assert pal["accent1"] == "112233"
    assert pal["accent2"] == DEFAULT_PALETTE["accent2"]
    assert palette_from_pack(None) == DEFAULT_PALETTE


def test_map_slide_colors_recolors_shapes_but_not_text_or_geometry() -> None:
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
    )
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor.from_string("14213D")
    box.line.color.rgb = RGBColor.from_string("5F6B86")
    box.text_frame.text = "Q1 질문"
    box.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string("3F4A63")
    geometry = (box.left, box.top, box.width, box.height)

    changed = map_slide_colors(slide)

    assert changed == 3
    assert str(box.fill.fore_color.rgb) in PALETTE_VALUES
    assert str(box.line.color.rgb) == "78788C"
    assert str(box.text_frame.paragraphs[0].runs[0].font.color.rgb) == "44546A"
    assert box.text_frame.text == "Q1 질문"
    assert (box.left, box.top, box.width, box.height) == geometry
    # idempotent: a second pass finds nothing left to move
    assert map_slide_colors(slide) == 0


def test_connector_keeps_its_own_color_when_copied() -> None:
    from pptx.enum.shapes import MSO_CONNECTOR
    from scripts import shape_migrate_lib as bac

    prs = Presentation()
    src = prs.slides.add_slide(prs.slide_layouts[6])
    dst = prs.slides.add_slide(prs.slide_layouts[6])
    conn = src.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(1), Inches(3), Inches(1))
    conn.line.color.rgb = RGBColor.from_string("5F6B86")

    bac.copy_line_native(dst, conn)

    copied = [s for s in dst.shapes if s._element.tag.endswith("}cxnSp")][0]
    assert str(copied.line.color.rgb) == "5F6B86"
