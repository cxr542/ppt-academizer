"""SVG rasterize must fix [Content_Types] for .png parts."""

from __future__ import annotations

from scripts.sanitize_pptx_package import _fix_png_parts_declared_as_svg


def test_fix_png_parts_declared_as_svg() -> None:
    raw = (
        '<Override PartName="/ppt/media/image3.png" ContentType="image/svg+xml"/>'
        '<Override PartName="/ppt/media/image5.png" ContentType="image/svg+xml"/>'
    )
    fixed = _fix_png_parts_declared_as_svg(raw)
    assert "image/svg+xml" not in fixed
    assert fixed.count("image/png") == 2
