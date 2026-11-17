"""Map the colors of copied body shapes onto the academy theme palette.

The body of a slide is copied in place so the original structure survives, but its colors are
the source deck's own (navy ``14213D``, slate ``3F4A63`` ...). Only the shape *colors* are
mapped here: positions, sizes, text and structure are untouched.

Rules (see ``map_color``):
* white / near white and colors already in the palette stay as they are;
* light tints (card backgrounds) become a tint of the nearest palette color with the same
  lightness, so the light/dark hierarchy of the source survives;
* other colors move to the nearest palette color, unless they are clearly a different hue
  (a green "OK" or a red "error"), which the academy palette cannot express, so they stay.
"""
from __future__ import annotations

import colorsys

from pptx.enum.shapes import MSO_SHAPE_TYPE

# Theme of the academy template (OKESTRO). The brand pack overrides these when present.
DEFAULT_PALETTE: dict[str, str] = {
    "dk1": "000000",
    "dk2": "44546A",
    "accent1": "006DFF",
    "accent2": "00187A",
    "accent3": "1F286F",
    "accent4": "1139ED",
    "accent5": "F15928",
    "accent6": "78788C",
}

NEAR_WHITE_LUMA = 0.96
LIGHT_TINT_LUMA = 0.78
FOREIGN_HUE_DISTANCE = 110.0
NEUTRAL_SATURATION = 0.10
FOREIGN_MIN_SATURATION = 0.30
FOREIGN_HUE_GAP = 0.07  # of the 0..1 hue circle (about 25 degrees)

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_SKIP_ANCESTORS = frozenset({"blip", "graphicFrame", "pic"})


def palette_from_pack(pack: dict | None) -> dict[str, str]:
    """Theme colors from the brand pack, falling back to the template defaults."""
    palette = dict(DEFAULT_PALETTE)
    theme = (pack or {}).get("theme") or {}
    for key in palette:
        value = theme.get(key)
        if isinstance(value, str):
            hex_v = value.lstrip("#").upper()
            if len(hex_v) == 6 and all(c in "0123456789ABCDEF" for c in hex_v):
                palette[key] = hex_v
    return palette


def _rgb(hex_v: str) -> tuple[int, int, int]:
    return int(hex_v[0:2], 16), int(hex_v[2:4], 16), int(hex_v[4:6], 16)


def _hex(rgb: tuple[float, float, float]) -> str:
    return "".join(f"{max(0, min(255, round(c))):02X}" for c in rgb)


def _luma(rgb: tuple[int, int, int]) -> float:
    r, g, b = rgb
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


def _saturation(rgb: tuple[int, int, int]) -> float:
    return colorsys.rgb_to_hsv(*(c / 255.0 for c in rgb))[1]


def _distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5


def _hue_distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    ha = colorsys.rgb_to_hsv(*(c / 255.0 for c in a))[0]
    hb = colorsys.rgb_to_hsv(*(c / 255.0 for c in b))[0]
    d = abs(ha - hb)
    return min(d, 1.0 - d)


def map_color(src_hex: str, palette: dict[str, str] | None = None) -> str:
    """The academy palette color for ``src_hex`` (``RRGGBB``), or ``src_hex`` itself."""
    pal = palette or DEFAULT_PALETTE
    src_hex = src_hex.upper()
    src = _rgb(src_hex)
    luma = _luma(src)
    values = {v.upper() for v in pal.values()}
    if src_hex in values or luma >= NEAR_WHITE_LUMA:
        return src_hex

    if luma >= LIGHT_TINT_LUMA:
        # A light tint: keep its lightness, take the hue of the nearest palette color.
        if _saturation(src) < NEUTRAL_SATURATION:
            base_hex = pal["dk2"]
        else:
            candidates = [pal[k] for k in pal if k not in ("dk1",)]
            base_hex = min(candidates, key=lambda h: _hue_distance(src, _rgb(h)))
        base = _rgb(base_hex)
        base_luma = _luma(base)
        if base_luma >= 0.999:
            return src_hex
        t = max(0.0, min(1.0, (luma - base_luma) / (1.0 - base_luma)))
        return _hex(tuple(255 * t + c * (1 - t) for c in base))

    nearest = min(values, key=lambda h: _distance(src, _rgb(h)))
    saturation = _saturation(src)
    foreign = _distance(src, _rgb(nearest)) > FOREIGN_HUE_DISTANCE or (
        saturation >= FOREIGN_MIN_SATURATION
        and _hue_distance(src, _rgb(nearest)) > FOREIGN_HUE_GAP
    )
    if foreign and saturation >= NEUTRAL_SATURATION:
        return src_hex  # a distinct hue (status green/red): the palette has no stand-in
    return nearest


def _in_skipped_part(el) -> bool:
    node = el.getparent()
    while node is not None:
        if node.tag.split("}")[-1] in _SKIP_ANCESTORS:
            return True
        node = node.getparent()
    return False


def map_slide_colors(slide, palette: dict[str, str] | None = None) -> int:
    """Remap explicit RGB colors of the non-placeholder shapes of ``slide``; returns the count.

    Pictures, tables and chart frames are left alone (a picture's pixels are not recolored, and
    tables are styled by their own pass).
    """
    pal = palette or DEFAULT_PALETTE
    changed = 0
    for shape in slide.shapes:
        if shape.is_placeholder or shape.shape_type in (
            MSO_SHAPE_TYPE.PICTURE,
            MSO_SHAPE_TYPE.TABLE,
            MSO_SHAPE_TYPE.CHART,
        ):
            continue
        for el in shape._element.iter(f"{{{_A_NS}}}srgbClr"):
            if _in_skipped_part(el):
                continue
            old = el.get("val")
            if not old or len(old) != 6:
                continue
            new = map_color(old, pal)
            if new != old.upper():
                el.set("val", new)
                changed += 1
    return changed
