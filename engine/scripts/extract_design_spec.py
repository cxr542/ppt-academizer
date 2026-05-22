#!/usr/bin/env python3
"""Extract canvas, layout names, and theme colors from the academy 2026 template."""

from __future__ import annotations

import json
import sys
from pathlib import Path

# Allow running as script from ppt-test root
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR

from scripts.academy_template import resolve_academy_template_path

_ACCENT_KEYS = (
    "ACCENT_1",
    "ACCENT_2",
    "ACCENT_3",
    "ACCENT_4",
    "ACCENT_5",
    "ACCENT_6",
    "TEXT_1",
    "TEXT_2",
    "BACKGROUND_1",
    "BACKGROUND_2",
    "HYPERLINK",
    "FOLLOWED_HYPERLINK",
)


def _rgb_hex(color) -> str | None:
    if color is None or color.type is None:
        return None
    try:
        if hasattr(color, "rgb") and color.rgb is not None:
            r, g, b = color.rgb
            return f"#{r:02X}{g:02X}{b:02X}"
    except (AttributeError, TypeError):
        pass
    return None


def extract_spec(template_path: Path) -> dict:
    prs = Presentation(str(template_path))
    layouts = []
    for i, layout in enumerate(prs.slide_master.slide_layouts):
        layouts.append({"index": i, "name": layout.name})

    colors: dict[str, str | None] = {}
    try:
        theme = prs.slide_master.theme
        for key in _ACCENT_KEYS:
            enum_val = getattr(MSO_THEME_COLOR, key, None)
            if enum_val is None:
                continue
            try:
                colors[key.lower()] = _rgb_hex(theme.theme_color_scheme.color(enum_val))
            except Exception:
                colors[key.lower()] = None
    except Exception as exc:
        colors["_error"] = str(exc)

    return {
        "template": str(template_path),
        "slide_width_emu": prs.slide_width,
        "slide_height_emu": prs.slide_height,
        "slide_width_in": round(prs.slide_width / 914400, 4),
        "slide_height_in": round(prs.slide_height / 914400, 4),
        "layout_count": len(layouts),
        "layouts": layouts,
        "theme_colors": colors,
    }


def main() -> None:
    path = resolve_academy_template_path()
    spec = extract_spec(path)
    if "--json" in sys.argv:
        print(json.dumps(spec, ensure_ascii=False, indent=2))
    else:
        print(f"Template: {spec['template']}")
        print(
            f"Canvas: {spec['slide_width_in']}\" x {spec['slide_height_in']}\" "
            f"({spec['slide_width_emu']} x {spec['slide_height_emu']} EMU)"
        )
        print(f"Layouts ({spec['layout_count']}):")
        for row in spec["layouts"]:
            print(f"  [{row['index']:2d}] {row['name']}")
        print("Theme colors (sample):")
        for k, v in spec.get("theme_colors", {}).items():
            if not k.startswith("_"):
                print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
