"""Academy brand pack: logos, theme colors, icon catalog (from template + elements)."""
from __future__ import annotations

import json
import os
import re
from functools import lru_cache
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

# McKinsey / partner decks often use cyan accents; map to academy accent1.
_NON_BRAND_ACCENTS = frozenset(
    {
        "00A3E0",
        "00B0F0",
        "00AEEF",
        "05A3E0",
    }
)


def _hex_to_rgb(value: str) -> RGBColor:
    h = value.strip().lstrip("#").upper()
    if len(h) != 6:
        raise ValueError(f"invalid hex color: {value!r}")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def resolve_brand_dir(template_path: Path | None = None) -> Path | None:
    """Locate brand/ directory next to TEMPLATE_PPTX or via BRAND_DIR."""
    env = os.environ.get("BRAND_DIR", "").strip()
    if env:
        p = Path(env).expanduser()
        if p.is_dir() and (p / "colors.json").is_file():
            return p
    if template_path is not None:
        cand = Path(template_path).expanduser().resolve().parent / "brand"
        if cand.is_dir() and (cand / "colors.json").is_file():
            return cand
    # Local Render-config mirror used by ppt-academizer developers.
    local = Path.home() / ".config" / "ppt-academizer-render" / "brand"
    if local.is_dir() and (local / "colors.json").is_file():
        return local
    return None


@lru_cache(maxsize=4)
def load_brand_pack(brand_dir: str) -> dict:
    root = Path(brand_dir)
    colors = json.loads((root / "colors.json").read_text(encoding="utf-8"))
    placements = {}
    if (root / "placements.json").is_file():
        placements = json.loads((root / "placements.json").read_text(encoding="utf-8"))
    icons = {"icons": []}
    if (root / "icon-catalog.json").is_file():
        icons = json.loads((root / "icon-catalog.json").read_text(encoding="utf-8"))
    theme = colors.get("theme") or {}
    return {
        "dir": root,
        "colors": colors,
        "theme": theme,
        "placements": placements,
        "icons": icons.get("icons") or [],
    }


def brand_rgb(pack: dict, role_or_hex: str, default: str = "#000000") -> RGBColor:
    theme = pack.get("theme") or {}
    roles = (pack.get("colors") or {}).get("roles") or {}
    key = roles.get(role_or_hex, role_or_hex)
    hex_v = theme.get(key) or theme.get(role_or_hex) or default
    if isinstance(hex_v, str) and hex_v.startswith("#"):
        return _hex_to_rgb(hex_v)
    return _hex_to_rgb(default)


def lookup_icon(pack: dict, label: str) -> Path | None:
    """Exact / normalized label match into brand/icons."""
    if not label:
        return None
    needle = re.sub(r"[^a-z0-9]+", "", label.lower())
    root = pack["dir"]
    for item in pack.get("icons") or []:
        lab = re.sub(r"[^a-z0-9]+", "", str(item.get("label") or "").lower())
        if lab == needle or needle in lab or lab in needle:
            path = root / item["file"]
            if path.is_file():
                return path
    return None


def _slide_has_picture_near(slide, left: int, top: int, *, tol: int = 200_000) -> bool:
    for sh in slide.shapes:
        if sh.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        if abs(int(sh.left or 0) - left) <= tol and abs(int(sh.top or 0) - top) <= tol:
            return True
    return False


def ensure_brand_pictures(slide, pack: dict, kind: str) -> int:
    """Stamp brand logos onto the slide when missing (layout-independent)."""
    placements = (pack.get("placements") or {}).get(kind) or {}
    if not placements:
        return 0
    root = pack["dir"]
    added = 0
    for _name, spec in placements.items():
        rel = spec.get("file")
        if not rel:
            continue
        path = root / rel
        if not path.is_file():
            continue
        left = int(spec["left"])
        top = int(spec["top"])
        width = int(spec["width"])
        height = int(spec["height"])
        if _slide_has_picture_near(slide, left, top):
            continue
        slide.shapes.add_picture(str(path), Emu(left), Emu(top), Emu(width), Emu(height))
        added += 1
    return added


def _rgb_hex_of_fill(shape) -> str | None:
    try:
        if shape.fill.type != MSO_FILL.SOLID:
            return None
        return str(shape.fill.fore_color.rgb).upper()
    except Exception:
        return None


def normalize_brand_accents(slide, pack: dict) -> int:
    """Remap common non-academy cyan accents to theme accent1."""
    accent = brand_rgb(pack, "accent", "#006DFF")
    target_hex = f"{accent[0]:02X}{accent[1]:02X}{accent[2]:02X}"
    changed = 0

    def walk(shapes) -> None:
        nonlocal changed
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                walk(shape.shapes)
                continue
            fill_hex = _rgb_hex_of_fill(shape)
            if fill_hex in _NON_BRAND_ACCENTS:
                try:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = accent
                    changed += 1
                except Exception:
                    pass
            if not getattr(shape, "has_text_frame", False):
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    try:
                        rh = str(run.font.color.rgb).upper()
                    except Exception:
                        continue
                    if rh in _NON_BRAND_ACCENTS:
                        run.font.color.rgb = accent
                        changed += 1

    walk(slide.shapes)
    # Avoid no-op warning when already academy blue
    _ = target_hex
    return changed
