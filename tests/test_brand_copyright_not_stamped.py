"""The brand pack must not stamp its copyright picture: its year is baked into the pixels."""

from __future__ import annotations

from pathlib import Path

from PIL import Image
from pptx import Presentation

from scripts.academy_brand import ensure_brand_pictures


def _pack(tmp_path: Path) -> dict:
    logos = tmp_path / "logos"
    logos.mkdir()
    for name in ("wordmark.png", "copyright.png", "copyright-content.png"):
        Image.new("RGBA", (40, 4), (255, 255, 255, 255)).save(logos / name)
    spec = lambda f, left, top: {"file": f"logos/{f}", "left": left, "top": top, "width": 900000, "height": 90000}
    return {
        "dir": tmp_path,
        "placements": {
            "cover": {
                "academy_wordmark": spec("wordmark.png", 500000, 900000),
                "copyright": spec("copyright.png", 8300000, 6300000),
            },
            "content": {
                "header_wordmark": spec("wordmark.png", 10700000, 370000),
                "footer_copyright": spec("copyright-content.png", 8800000, 6500000),
            },
        },
    }


def _bottom_pictures(slide) -> list:
    return [s for s in slide.shapes if s.shape_type == 13 and s.top > 6_000_000]


def test_copyright_picture_is_not_stamped_but_other_logos_are(tmp_path: Path) -> None:
    pack = _pack(tmp_path)
    prs = Presentation()
    for kind in ("cover", "content"):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        added = ensure_brand_pictures(slide, pack, kind)
        assert added == 1
        assert _bottom_pictures(slide) == []
        assert len([s for s in slide.shapes if s.shape_type == 13]) == 1
