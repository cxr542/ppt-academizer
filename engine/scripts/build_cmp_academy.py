#!/usr/bin/env python3
"""Migrate CMP source deck into OKESTRO academy layout (academy-design.md §7)."""
from __future__ import annotations

import argparse
import importlib.util
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

PPT_TEST = Path(__file__).resolve().parent.parent
if str(PPT_TEST) not in sys.path:
    sys.path.insert(0, str(PPT_TEST))

from scripts.academy_template import resolve_academy_template_path  # noqa: E402
from scripts.deck_migrate_config import (  # noqa: E402
    DeckMigrateConfig,
    migrate_config_for_source,
)
from scripts.pptx_ingest import slide_notes_text  # noqa: E402

import scripts.academy_deck_build_lib as adl  # noqa: E402

from scripts.academy_deck_build_lib import (  # noqa: E402
    FONT_BODY,
    FONT_TITLE,
    _assign_plain,
    _autofit_horizontal_textbox,
    _capture_layout_placeholder_geometry,
    _capture_seed_placeholder_geometry,
    _fill_chapter_guide,
    _fill_contents_guide,
    _fill_cover,
    _ph_by_idx,
    _prepare_slide_for_editing,
    _restore_slide_placeholder_geometry,
    configure_text_frame_for_wrap,
    delete_slide,
    duplicate_slide_from_seed,
    layout_seed_slides,
    save_academy_deck,
    set_tf_font,
)

# design.md — OKESTRO theme text colors
CLR_DK1 = RGBColor(0x00, 0x00, 0x00)
CLR_DK2 = RGBColor(0x44, 0x54, 0x6A)
CLR_ACCENT1 = RGBColor(0x00, 0x6D, 0xFF)

_SHAPE_LIB = PPT_TEST / "scripts" / "shape_migrate_lib.py"
_SPEC = importlib.util.spec_from_file_location("bac", _SHAPE_LIB)
bac = importlib.util.module_from_spec(_SPEC)
assert _SPEC.loader is not None
_SPEC.loader.exec_module(bac)

# CMP diagrams use large exported PNG panels — not PaaS gray card backgrounds.
DROP_CARD_PANEL_PICTURES = False

DEFAULT_SOURCE = Path.home() / "Desktop" / "클라우드 구현기술(CMP)_v1.0_수정요청.pptx"

# Legacy aliases (CMP defaults — prefer DeckMigrateConfig).
SKIP_SRC_INDICES = frozenset({8})
CMP_TOC_SLIDE_RANGE = range(2, 8)

_HANGUL = re.compile(r"[\uAC00-\uD7A3]")
_STEP_NUMBER_LINE = re.compile(r"^\d{1,3}$")

# Academy template layouts — used to detect double-migration of an output deck.
ACADEMY_LAYOUT_MARKERS = frozenset(
    {"2_표지", "내지_거버닝 O", "1_내지_거버닝 X", "간지", "목차"}
)

# Official layout coords (내지_거버닝 O) — do not use PaaS-wide title width.
LAYOUT_PH10_LEFT = 781_263
CONTENT_BODY_LEFT = 962_025
CONTENT_COLUMN_GAP = 360_000
CONTENT_RIGHT_COL_RATIO = 0.47  # right column starts ~47% of slide width
TITLE_PT = 18
TITLE_MAX_CHARS = 72
TITLE_SIDEBAR_LEFT_MAX_IN = 0.32
BODY_SHAPE_PT = 12
CAPTION_SHAPE_PT = 12

LAYOUT_COVER = "2_표지"
LAYOUT_TOC = "목차"  # 간지1
LAYOUT_SECTION = "간지"  # 간지2
LAYOUT_CONTENT = "내지_거버닝 O"

HEADER_TITLE_TOP_MAX = 900_000
SECTION_PH10_MIN_WIDTH = int(Inches(3.6))
# Radial hub diagrams (CMP 6-tech slide): not 2-column body — skip column relayout.
RADIAL_MIN_BODY_LABELS = 5
FULL_BLEED_W_RATIO = 0.88
FULL_BLEED_H_RATIO = 0.55
TOP_BAND_PICTURE_RATIO = 0.45


def uses_academy_layouts(prs: Presentation) -> bool:
    """Source deck is authored on OKESTRO academy masters (allowed to migrate)."""
    names = {s.slide_layout.name for s in prs.slides}
    return len(names & ACADEMY_LAYOUT_MARKERS) >= 2


def is_likely_reacademize_output(prs: Presentation, *, filename_hint: str = "") -> bool:
    """True only for ppt-academizer *output* re-uploaded — not academy-template originals."""
    low = (filename_hint or "").lower().replace(" ", "-")
    for marker in (
        "academy-output",
        "academy-cmp-",
        "academy-cmp-academy",
        "_ppfixed",
        "_repaired",
    ):
        if marker in low:
            return True

    if not uses_academy_layouts(prs):
        return False

    dup_num_headers = 0
    content_slides = 0
    for slide in prs.slides:
        if slide.slide_layout.name != LAYOUT_CONTENT:
            continue
        content_slides += 1
        t10 = t12 = ""
        for sh in slide.placeholders:
            idx = sh.placeholder_format.idx
            if idx == 10:
                t10 = (sh.text or "").strip()
            elif idx == 12:
                t12 = (sh.text or "").strip()
        if t10 and t10 == t12 and _is_step_number_line(t10):
            dup_num_headers += 1
    if content_slides and dup_num_headers / content_slides >= 0.45:
        return True
    return False


# Back-compat alias (narrower than v1.3.4 guard).
is_academy_template_deck = is_likely_reacademize_output


def _is_step_number_line(line: str) -> bool:
    """Google/partner step column or re-migrated ph12 — not a slide title."""
    return bool(_STEP_NUMBER_LINE.match((line or "").strip()))


def _hangul_ratio(text: str) -> float:
    compact = re.sub(r"\s+", "", text or "")
    if not compact:
        return 0.0
    return len(_HANGUL.findall(compact)) / len(compact)


def classify_slide(
    slide,
    index: int,
    part_cover_indices: frozenset[int] | None = None,
) -> str:
    covers = part_cover_indices if part_cover_indices is not None else frozenset({0, 13})
    texts = [
        (int(sh.top or 0), sh.text.strip())
        for sh in slide.shapes
        if sh.has_text_frame and sh.text.strip()
    ]
    texts.sort()
    if len(slide.shapes) <= 1 and not texts:
        return "empty"
    if not texts:
        return "empty"
    first_top, first_text = texts[0]
    if index in covers:
        return "cover"
    if (
        len(texts) <= 3
        and first_top > 2_000_000
        and len(first_text) < 80
        and first_text == first_text.upper()
    ):
        return "section"
    return "content"


def extract_header(src_slide, kind: str) -> tuple[str | None, str | None]:
    texts = [
        (int(sh.top or 0), sh.text.strip())
        for sh in src_slide.shapes
        if sh.has_text_frame and sh.text.strip()
    ]
    texts.sort()
    if not texts:
        return None, None
    if kind == "cover":
        title = texts[0][1].replace("\n", " ").strip()
        governing = texts[1][1].split("\n")[0].strip() if len(texts) > 1 else None
        return title, governing
    if kind == "section":
        return texts[0][1].split("\n")[0].strip(), None
    header_candidates: list[tuple[int, int, str]] = []
    for top, text in texts:
        if top > HEADER_TITLE_TOP_MAX:
            continue
        line = text.replace("\x0b", " ").split("\n")[0].strip()
        if not line or len(line) > TITLE_MAX_CHARS or _is_step_number_line(line):
            continue
        header_candidates.append((top, len(line), line))
    if header_candidates:
        header_candidates.sort(key=lambda c: (c[0], c[1]))
        return header_candidates[0][2], None
    for top, text in texts:
        line = text.replace("\x0b", " ").split("\n")[0].strip()
        if len(line) <= TITLE_MAX_CHARS and not _is_step_number_line(line):
            return line, None
    return texts[0][1].replace("\x0b", " ").split("\n")[0].strip(), None


def build_slide_plan(
    src: Presentation,
    cfg: DeckMigrateConfig | None = None,
) -> list[tuple[int, str]]:
    if cfg is None:
        cfg = migrate_config_for_source("cmp-partner.pptx")
    plan: list[tuple[int, str]] = []
    for i, slide in enumerate(src.slides):
        if i in cfg.skip_src_indices:
            continue
        if i == 0:
            plan.append((i, "cover"))
            continue
        kind = classify_slide(slide, i, cfg.part_cover_indices)
        if kind == "empty":
            plan.append((i, "content"))
        else:
            plan.append((i, kind))
    return plan


def expected_output_slide_count(plan: list[tuple[int, str]], cfg: DeckMigrateConfig) -> int:
    extra_toc = sum(1 for _, k in plan if k == "cover" and cfg.auto_toc_after_cover)
    return len(plan) + extra_toc


def collect_section_titles(
    src: Presentation,
    after_src_idx: int,
    *,
    part_cover_indices: frozenset[int] | None = None,
) -> list[str]:
    """Section titles until the next part cover (간지1 목차용)."""
    covers = part_cover_indices or frozenset({0, 13})
    titles: list[str] = []
    for i in range(after_src_idx, len(src.slides)):
        if i != after_src_idx and classify_slide(src.slides[i], i, covers) == "cover":
            break
        if classify_slide(src.slides[i], i, covers) == "section":
            t, _ = extract_header(src.slides[i], "section")
            if t:
                titles.append(t)
    return titles


def is_full_bleed_picture(shape, slide_width: int, slide_height: int) -> bool:
    """Drop slide-wide export backgrounds only — keep right-column hero diagrams."""
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    w, h = int(shape.width or 0), int(shape.height or 0)
    top, left = int(shape.top or 0), int(shape.left or 0)
    if w >= slide_width * FULL_BLEED_W_RATIO and h >= slide_height * FULL_BLEED_H_RATIO:
        return True
    # Wide top banner (nearly full slide width), not ~50% column art.
    if top < 200_000 and w >= slide_width * 0.75:
        return True
    # Top-left logo strip from Google export.
    if top < 200_000 and left < slide_width * 0.08 and w >= slide_width * 0.35:
        return True
    return False


def is_card_panel_picture(shape, slide_height: int) -> bool:
    """Rounded gray card panels exported as large pictures — drop on white academy slides."""
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    w, h = int(shape.width or 0), int(shape.height or 0)
    top = int(shape.top or 0)
    if top < bac.HEADER_BOTTOM_EMU:
        return False
    if w >= 2_800_000 and h >= 1_800_000:
        return True
    return False


def _walk_shapes(container):
    """Depth-first over slide or group shapes."""
    for shape in container.shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk_shapes(shape)


def is_slide_background_picture(shape, slide_width: int, slide_height: int) -> bool:
    """Only drop true slide backgrounds (§6.7) — not chart/diagram images."""
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    w, h = int(shape.width or 0), int(shape.height or 0)
    top, left = int(shape.top or 0), int(shape.left or 0)
    if w >= slide_width * 0.97 and h >= slide_height * 0.93:
        return True
    if top < 200_000 and w >= slide_width * 0.80:
        return True
    if top < 200_000 and left < slide_width * 0.08 and w >= slide_width * 0.35:
        return True
    return False


def _is_diagram_raster_shape(shape, slide_width: int, slide_height: int) -> bool:
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return True
    if shape.shape_type == MSO_SHAPE_TYPE.CHART:
        return True
    if bac._is_graphic_frame(shape):
        return True
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return any(
            _is_diagram_raster_shape(child, slide_width, slide_height)
            for child in shape.shapes
        )
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return not is_slide_background_picture(shape, slide_width, slide_height)
    return False


def count_diagram_rasters(slide, slide_width: int, slide_height: int) -> int:
    return sum(
        1
        for shp in _walk_shapes(slide)
        if _is_diagram_raster_shape(shp, slide_width, slide_height)
    )


def is_image_primary_slide(src_slide, slide_width: int, slide_height: int) -> bool:
    """§6.7 — slide is mostly diagram/chart pictures with little body text."""
    rasters = count_diagram_rasters(src_slide, slide_width, slide_height)
    if rasters < 1:
        return False
    body_texts = 0
    for sh in _walk_shapes(src_slide):
        if not sh.has_text_frame or not (sh.text or "").strip():
            continue
        if int(sh.top or 0) <= bac.HEADER_BOTTOM_EMU:
            continue
        body_texts += 1
    return body_texts <= 2


def is_raster_heavy_slide(src_slide, slide_width: int, slide_height: int) -> bool:
    """Permissive §6.7 copy when diagrams dominate (incl. slides 11, 18, 23, 24)."""
    rasters = count_diagram_rasters(src_slide, slide_width, slide_height)
    if rasters == 0:
        return False
    if rasters >= 2:
        return True
    return is_image_primary_slide(src_slide, slide_width, slide_height)


def should_drop_body_picture(
    shape,
    slide_width: int,
    slide_height: int,
    *,
    permissive_raster: bool,
) -> bool:
    """§6.7 — on diagram-heavy slides only drop slide backgrounds."""
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    if permissive_raster:
        return is_slide_background_picture(shape, slide_width, slide_height)
    if is_slide_background_picture(shape, slide_width, slide_height):
        return True
    if is_full_bleed_picture(shape, slide_width, slide_height):
        return True
    if DROP_CARD_PANEL_PICTURES and is_card_panel_picture(shape, slide_height):
        return True
    return False


def _is_top_level_raster_candidate(shape) -> bool:
    if shape.shape_type in (
        MSO_SHAPE_TYPE.PICTURE,
        MSO_SHAPE_TYPE.GROUP,
        MSO_SHAPE_TYPE.CHART,
    ):
        return True
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return True
    return bac._is_graphic_frame(shape)


def _dst_has_similar_bbox(
    dst_slide,
    left: int,
    top: int,
    width: int,
    height: int,
    *,
    tol: int = 120_000,
) -> bool:
    for sh in dst_slide.shapes:
        if sh.is_placeholder:
            continue
        if int(sh.left or 0) - tol <= left <= int(sh.left or 0) + tol:
            if int(sh.top or 0) - tol <= top <= int(sh.top or 0) + tol:
                if abs(int(sh.width or 0) - width) < tol * 2:
                    return True
    return False


def sync_raster_diagram_assets(
    src_slide,
    dst_slide,
    title,
    governing,
    slide_width: int,
    slide_height: int,
    *,
    skip_connectors: bool = False,
    force_permissive: bool = False,
    src_classify_empty: bool = False,
) -> int:
    """§6.7 — copy missing chart/diagram images when dst has fewer rasters than src."""
    src_n = count_diagram_rasters(src_slide, slide_width, slide_height)
    dst_n = count_diagram_rasters(dst_slide, slide_width, slide_height)
    if src_n == 0:
        return 0
    permissive = (
        force_permissive
        or src_classify_empty
        or is_raster_heavy_slide(src_slide, slide_width, slide_height)
        or is_image_primary_slide(src_slide, slide_width, slide_height)
    )
    if dst_n >= src_n and not (src_classify_empty and src_n > 0):
        return 0

    id_remap: dict[str, str] = {}
    copied = 0
    for shape in src_slide.shapes:
        if not _is_top_level_raster_candidate(shape):
            continue
        if not bac.is_body_shape(shape, title, governing):
            continue
        if should_drop_body_picture(
            shape,
            slide_width,
            slide_height,
            permissive_raster=permissive,
        ):
            continue
        if skip_connectors and (
            bac._is_cxn_sp(shape) or shape.shape_type == MSO_SHAPE_TYPE.LINE
        ):
            continue
        left, top = int(shape.left or 0), int(shape.top or 0)
        w, h = int(shape.width or 0), int(shape.height or 0)
        if _dst_has_similar_bbox(dst_slide, left, top, w, h):
            continue
        n_before = len(dst_slide.shapes)
        bac.copy_shape_hybrid(dst_slide, shape, id_remap=id_remap)
        if len(dst_slide.shapes) > n_before:
            copied += 1
    if copied:
        renumber_map = bac.renumber_shape_ids(dst_slide)
        bac.fix_orphan_connector_refs(dst_slide, id_remap, renumber_map)
    return copied


def is_radial_diagram_slide(slide, slide_width: int, slide_height: int) -> bool:
    """Hub / 6-around-center layouts (not 2-column text). CMP slide 2 is the canonical case."""
    body: list[tuple[int, int]] = []
    for sh in slide.shapes:
        if not sh.has_text_frame or not (sh.text or "").strip():
            continue
        top = int(sh.top or 0)
        if top <= bac.HEADER_BOTTOM_EMU:
            continue
        body.append((int(sh.left or 0), top))
    if len(body) < RADIAL_MIN_BODY_LABELS:
        return False
    lefts = [p[0] for p in body]
    tops = [p[1] for p in body]
    if max(lefts) - min(lefts) < slide_width * 0.35:
        return False
    if max(tops) - min(tops) < slide_height * 0.22:
        return False
    mid_x = slide_width / 2
    bands = {"left": 0, "right": 0, "center": 0}
    for left, _top in body:
        if left < mid_x - slide_width * 0.12:
            bands["left"] += 1
        elif left > mid_x + slide_width * 0.12:
            bands["right"] += 1
        else:
            bands["center"] += 1
    return bands["left"] >= 2 and bands["right"] >= 1 and bands["center"] >= 1


def copy_body_shapes_filtered(
    src_slide,
    dst_slide,
    title,
    governing,
    slide_width: int,
    slide_height: int,
    *,
    skip_connectors: bool = False,
) -> dict[str, str]:
    permissive = is_raster_heavy_slide(src_slide, slide_width, slide_height)
    id_remap: dict[str, str] = {}
    for shape in src_slide.shapes:
        if not bac.is_body_shape(shape, title, governing):
            continue
        if should_drop_body_picture(
            shape, slide_width, slide_height, permissive_raster=permissive
        ):
            continue
        if skip_connectors and (
            bac._is_cxn_sp(shape) or shape.shape_type == MSO_SHAPE_TYPE.LINE
        ):
            continue
        bac.copy_shape_hybrid(dst_slide, shape, id_remap=id_remap)
    return id_remap


def _luminance(rgb: RGBColor) -> float:
    return 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]


def academy_color_for_run(run, *, title: bool = False) -> None:
    """Map Google-export light-on-dark text to design.md colors on white slides."""
    if title:
        run.font.color.rgb = CLR_DK1
        return
    try:
        c = run.font.color
        if c.type == MSO_COLOR_TYPE.RGB and c.rgb is not None:
            rgb = c.rgb
            lum = _luminance(rgb)
            if lum >= 220:
                run.font.color.rgb = CLR_DK1
            elif lum >= 140:
                run.font.color.rgb = CLR_DK2
            return
    except Exception:
        pass
    run.font.color.rgb = CLR_DK1


def apply_academy_colors_to_text_frame(tf, *, title: bool = False) -> None:
    configure_text_frame_for_wrap(tf)
    set_tf_font(tf, title=title)
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.name = FONT_TITLE if title else FONT_BODY
            academy_color_for_run(run, title=title)


def apply_academy_fonts_and_colors(slide) -> None:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        is_title_ph = shape.is_placeholder and shape.placeholder_format.idx == 10
        is_num_ph = shape.is_placeholder and shape.placeholder_format.idx == 12
        if is_title_ph:
            continue
        apply_academy_colors_to_text_frame(
            shape.text_frame,
            title=False,
        )
        if is_num_ph:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = CLR_DK2


def _apply_body_text_frame_style(tf, *, caption: bool = False) -> None:
    pt = CAPTION_SHAPE_PT if caption else BODY_SHAPE_PT
    font_name = FONT_TITLE if caption else FONT_BODY
    configure_text_frame_for_wrap(tf)
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.name = font_name
            run.font.size = Pt(pt)
            run.font.color.rgb = CLR_DK1
            if caption:
                run.font.bold = True


def apply_academy_body_shape_typography(slide) -> None:
    """§6.6 — non-placeholder body shapes and shapes inside GROUPs."""

    def walk(shapes) -> None:
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                walk(shape.shapes)
                continue
            if shape.is_placeholder:
                idx = shape.placeholder_format.idx
                if idx in (10, 12, 13):
                    continue
            if not shape.has_text_frame:
                continue
            text = (shape.text or "").strip()
            if not text:
                continue
            first_line = text.replace("\x0b", " ").split("\n")[0]
            caption = len(first_line) <= 40 or (len(text) <= 80 and text.count("\n") < 2)
            _apply_body_text_frame_style(shape.text_frame, caption=caption)

    walk(slide.shapes)


def hide_empty_governing_placeholder(slide) -> None:
    """Remove template '거버닝 메시지…' guide when there is no governing text."""
    for shape in list(slide.placeholders):
        if shape.placeholder_format.idx != 13:
            continue
        if (shape.text or "").strip():
            return
        shape._element.getparent().remove(shape._element)
        return


def apply_content_header_geometry(slide, slide_width: int | None = None) -> None:
    """Official template placeholder positions (not PaaS-wide title width)."""
    _restore_slide_placeholder_geometry(slide)
    ph10 = _ph_by_idx(slide, 10)
    ph10.left = LAYOUT_PH10_LEFT
    # Title (ph10) stays on seed row above step number (ph12); do not align tops.
    if slide_width:
        max_w = slide_width - LAYOUT_PH10_LEFT - bac.SLIDE_MARGIN_X
        ph10.width = min(int(max_w), int(Inches(10.5)))
    elif int(ph10.width) > int(Inches(5.5)):
        ph10.width = int(Inches(3.0))


def normalize_slide_title(title: str) -> str:
    return " ".join((title or "").split())


def assign_content_slide_title(ph10, title: str) -> None:
    """Title: one line; Korean/English 18pt; left-aligned, vertically centered in box."""
    title = normalize_slide_title(title)
    tf = ph10.text_frame
    configure_text_frame_for_wrap(tf)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ph10.text = ""
    while len(tf.paragraphs) > 1:
        tf._element.remove(tf.paragraphs[-1]._p)

    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    para.clear()

    if ": " in title:
        head, tail = title.split(": ", 1)
        if _HANGUL.search(head) and not _HANGUL.search(tail):
            run = para.add_run()
            run.text = title
            run.font.name = FONT_TITLE
            run.font.size = Pt(TITLE_PT)
            run.font.color.rgb = CLR_DK1
        else:
            r_ko = para.add_run()
            r_ko.text = f"{head}: "
            r_ko.font.name = FONT_TITLE
            r_ko.font.size = Pt(TITLE_PT)
            r_ko.font.color.rgb = CLR_DK1
            r_en = para.add_run()
            r_en.text = tail
            r_en.font.name = FONT_TITLE
            r_en.font.size = Pt(TITLE_PT)
            r_en.font.color.rgb = CLR_DK1
    else:
        run = para.add_run()
        run.text = title
        run.font.name = FONT_TITLE
        run.font.size = Pt(TITLE_PT)
        run.font.color.rgb = CLR_DK1


def _content_ph(slide, idx: int):
    """Placeholder lookup without re-applying seed geometry (avoids undoing width tweaks)."""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    raise KeyError(f"Placeholder idx={idx} not found on {slide.slide_layout.name!r}")


def apply_slide_title_layout(slide, slide_width: int, title: str) -> None:
    apply_content_header_geometry(slide, slide_width)
    ph10 = _content_ph(slide, 10)
    assign_content_slide_title(ph10, title)
    if slide_width:
        max_w = slide_width - LAYOUT_PH10_LEFT - bac.SLIDE_MARGIN_X
        ph10.width = min(int(max_w), int(Inches(10.5)))


def relayout_content_columns(slide, slide_width: int) -> None:
    """Two-column body: move right column left and stretch both columns to margins."""
    if is_speaker_script_slide(slide):
        return
    max_right = slide_width - bac.SLIDE_MARGIN_X
    mid_threshold = int(slide_width * 0.44)
    right_left = int(slide_width * CONTENT_RIGHT_COL_RATIO)
    left_width = right_left - CONTENT_BODY_LEFT - CONTENT_COLUMN_GAP
    if left_width < bac.MIN_SHAPE_W:
        return

    bodies = [
        sh
        for sh in slide.shapes
        if not sh.is_placeholder
        and sh.has_text_frame
        and int(sh.top or 0) >= bac.HEADER_BOTTOM_EMU
        and (sh.text or "").strip()
    ]
    left_shapes = [sh for sh in bodies if int(sh.left or 0) < mid_threshold]
    right_shapes = [sh for sh in bodies if int(sh.left or 0) >= mid_threshold]
    if not right_shapes:
        return

    for sh in left_shapes:
        sh.left = CONTENT_BODY_LEFT
        sh.width = left_width
    right_width = max_right - right_left
    for sh in right_shapes:
        sh.left = right_left
        sh.width = right_width


def is_speaker_script_slide(slide) -> bool:
    texts = [
        (shape.text or "").strip()
        for shape in slide.shapes
        if shape.has_text_frame and (shape.text or "").strip()
    ]
    merged = "\n".join(texts)
    if "강사용" in merged or "강의 포인트" in merged:
        return True
    quoted_lines = sum(1 for text in texts if "“" in text or '"' in text)
    short_labels = sum(1 for text in texts if len(text) <= 16 and "\n" not in text)
    return quoted_lines >= 3 and short_labels >= 3


def copy_speaker_notes(src_slide, dst_slide) -> None:
    notes = slide_notes_text(src_slide)
    if notes:
        dst_slide.notes_slide.notes_text_frame.text = notes


def fill_cover_from_source(slide, src_slide, cfg: DeckMigrateConfig) -> None:
    title, governing = extract_header(src_slide, "cover")
    cover_text = title or "강의"
    if cfg.cover_subtitle:
        cover_text = f"{cover_text}\n{cfg.cover_subtitle}"
    if governing:
        cover_text = f"{cover_text}\n{governing}"
    _fill_cover(slide, [cover_text], title_font_pt=36, body_font_pt=16)


def apply_academy_table_style(table, *, header_rows: int = 1) -> None:
    """§6.5 — editable cells, academy fonts/colors."""
    for r_idx, row in enumerate(table.rows):
        is_header = r_idx < header_rows
        for cell in row.cells:
            tf = cell.text_frame
            if not tf:
                continue
            apply_academy_colors_to_text_frame(tf, title=is_header)
            if is_header:
                for para in tf.paragraphs:
                    for run in para.runs:
                        run.font.bold = True
                try:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(0xE7, 0xE6, 0xE6)
                except Exception:
                    pass


def apply_academy_tables_on_slide(slide) -> None:
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            apply_academy_table_style(shape.table)
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                if child.shape_type == MSO_SHAPE_TYPE.TABLE:
                    apply_academy_table_style(child.table)


def _count_charts(slide) -> int:
    n = 0
    for sh in slide.shapes:
        if getattr(sh, "has_chart", False) and sh.has_chart:
            n += 1
        elif sh.shape_type == MSO_SHAPE_TYPE.CHART:
            n += 1
    return n


def fill_toc(slide, section_titles: list[str]) -> None:
    lines = [f"• {t}" for t in section_titles] if section_titles else ["• (목차)"]
    _fill_contents_guide(slide, ["\n".join(lines)], title_font_pt=28, body_font_pt=18)


def fill_section_from_source(slide, src_slide, chapter_num: int) -> None:
    title, _ = extract_header(src_slide, "section")
    _fill_chapter_guide(
        slide,
        [title or "", str(chapter_num)],
        title_font_pt=36,
        body_font_pt=16,
    )
    ph10 = _ph_by_idx(slide, 10)
    ph10.width = max(int(ph10.width), SECTION_PH10_MIN_WIDTH)
    _autofit_horizontal_textbox(
        ph10,
        font_pt=36,
        min_width_in=3.6,
        max_width_in=11.0,
        char_factor=0.58,
        width_padding=1.12,
    )


def migrate_content_body(
    slide,
    src_slide,
    step_num: str,
    slide_width: int,
    slide_height: int,
    *,
    src_classify_empty: bool = False,
) -> tuple[bool, int]:
    """Returns (is_radial_hub, raster_diagrams_restored_count)."""
    title, governing = extract_header(src_slide, "content")
    gov = governing or ""
    radial = is_radial_diagram_slide(src_slide, slide_width, slide_height)
    force_raster = src_classify_empty or count_diagram_rasters(
        src_slide, slide_width, slide_height
    ) >= 1

    apply_content_header_geometry(slide, slide_width)

    _assign_plain(
        _ph_by_idx(slide, 12),
        step_num,
        title=False,
        title_font_pt=14,
        body_font_pt=14,
    )
    if gov and len(gov) <= bac.GOV_MAX_LEN:
        _assign_plain(
            _ph_by_idx(slide, 13),
            gov,
            title=False,
            title_font_pt=14,
            body_font_pt=14,
        )
    else:
        hide_empty_governing_placeholder(slide)

    id_remap = copy_body_shapes_filtered(
        src_slide,
        slide,
        title,
        governing,
        slide_width,
        slide_height,
        skip_connectors=radial,
    )
    restored = sync_raster_diagram_assets(
        src_slide,
        slide,
        title,
        governing,
        slide_width,
        slide_height,
        skip_connectors=radial,
        force_permissive=force_raster,
        src_classify_empty=src_classify_empty,
    )
    renumber_map = bac.renumber_shape_ids(slide)
    if not radial:
        bac.fix_orphan_connector_refs(slide, id_remap, renumber_map)
    apply_slide_title_layout(slide, slide_width, title or "")
    if not radial:
        relayout_content_columns(slide, slide_width)
    apply_academy_fonts_and_colors(slide)
    apply_academy_body_shape_typography(slide)
    apply_academy_tables_on_slide(slide)
    return radial, restored


def add_toc_after_cover(
    prs: Presentation,
    seeds: dict,
    src: Presentation,
    cover_src_idx: int,
    cfg: DeckMigrateConfig,
) -> None:
    if not cfg.auto_toc_after_cover:
        return
    titles: list[str] = []
    if cfg.toc_slide_range:
        for i in cfg.toc_slide_range:
            if i >= len(src.slides):
                break
            t, _ = extract_header(src.slides[i], "content")
            if t:
                titles.append(t.replace("\x0b", " ").strip())
    if not titles:
        titles = collect_section_titles(
            src, cover_src_idx + 1, part_cover_indices=cfg.part_cover_indices
        )
    if not titles:
        return
    slide = duplicate_slide_from_seed(prs, seeds[LAYOUT_TOC])
    _prepare_slide_for_editing(slide)
    fill_toc(slide, titles)


def migrate_cmp_deck(
    source: Path,
    output: Path,
    *,
    template_path: Path | None = None,
    skip_powerpoint_repair: bool = True,
    skip_ooxml_repair: bool = False,
) -> tuple[Path, list[dict], int]:
    """§7 migrate: seed clone + shape transplant. Returns (output_path, warnings)."""
    source = Path(source).resolve()
    output = Path(output).resolve()
    template = Path(template_path).resolve() if template_path else resolve_academy_template_path()

    if not source.is_file():
        raise FileNotFoundError(source)
    if not template.is_file():
        raise FileNotFoundError(template)

    warnings: list[dict] = []
    cfg = migrate_config_for_source(source)

    src = Presentation(str(source))
    if is_likely_reacademize_output(src, filename_hint=source.name):
        raise ValueError(
            "ppt-academizer 결과물(academy-output 등)은 다시 변환할 수 없습니다. "
            "원본 파트너·교육 .pptx를 업로드해 주세요."
        )
    if uses_academy_layouts(src):
        warnings.append(
            {
                "code": "ACADEMY_SOURCE_LAYOUT",
                "message": "아카데미 마스터 기반 원본 — §7로 레이아웃·도형을 다시 맞춥니다.",
            }
        )
    plan = build_slide_plan(src, cfg)
    if not plan:
        raise ValueError("No slides to migrate.")
    expected_slides = expected_output_slide_count(plan, cfg)

    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        shutil.copy2(output, output.with_name(output.stem + "_backup.pptx"))

    shutil.copy2(template, output)
    prs = Presentation(str(output))
    adl._PLACEHOLDER_GEOM = (
        _capture_seed_placeholder_geometry(prs),
        _capture_layout_placeholder_geometry(prs),
    )
    try:
        return _migrate_cmp_deck_body(
            prs,
            src,
            output,
            template,
            cfg,
            plan,
            expected_slides,
            warnings,
            skip_ooxml_repair=skip_ooxml_repair,
            skip_powerpoint_repair=skip_powerpoint_repair,
        )
    finally:
        adl._PLACEHOLDER_GEOM = None


def _migrate_cmp_deck_body(
    prs,
    src,
    output: Path,
    template: Path,
    cfg,
    plan,
    expected_slides: int,
    warnings: list,
    *,
    skip_ooxml_repair: bool,
    skip_powerpoint_repair: bool,
) -> tuple[Path, list[dict], int]:
    sw, sh = int(prs.slide_width), int(prs.slide_height)
    seeds = layout_seed_slides(prs)
    for name in (LAYOUT_COVER, LAYOUT_TOC, LAYOUT_SECTION, LAYOUT_CONTENT):
        if name not in seeds:
            raise ValueError(f"Template missing starter slide for layout {name!r}")

    initial_part_ids = {id(s.part) for s in prs.slides}

    content_step = 0
    section_num = 0
    radial_slides: list = []
    for src_idx, kind in plan:
        src_slide = src.slides[src_idx]
        if kind == "cover":
            content_step = 0
            slide = duplicate_slide_from_seed(prs, seeds[LAYOUT_COVER])
            _prepare_slide_for_editing(slide)
            fill_cover_from_source(slide, src_slide, cfg)
            copy_speaker_notes(src_slide, slide)
            add_toc_after_cover(prs, seeds, src, src_idx, cfg)
        elif kind == "section":
            content_step = 0
            section_num += 1
            slide = duplicate_slide_from_seed(prs, seeds[LAYOUT_SECTION])
            _prepare_slide_for_editing(slide)
            fill_section_from_source(slide, src_slide, section_num)
            copy_speaker_notes(src_slide, slide)
        else:
            content_step += 1
            slide = duplicate_slide_from_seed(prs, seeds[LAYOUT_CONTENT])
            _prepare_slide_for_editing(slide)
            copy_speaker_notes(src_slide, slide)
            charts_in = _count_charts(src_slide)
            src_empty = (
                classify_slide(src_slide, src_idx, cfg.part_cover_indices) == "empty"
            )
            radial, restored = migrate_content_body(
                slide,
                src_slide,
                str(content_step),
                sw,
                sh,
                src_classify_empty=src_empty,
            )
            if radial:
                radial_slides.append(slide)
            if restored:
                warnings.append(
                    {
                        "code": "RASTER_DIAGRAM_RESTORED",
                        "src_index": src_idx,
                        "count": restored,
                    }
                )
            charts_out = _count_charts(slide)
            if charts_in > charts_out:
                warnings.append(
                    {
                        "code": "CHART_NOT_COPIED",
                        "message": f"Source slide {src_idx}: charts {charts_in} → {charts_out}",
                        "src_index": src_idx,
                    }
                )
            if classify_slide(src_slide, src_idx, cfg.part_cover_indices) == "empty":
                warnings.append(
                    {
                        "code": "SLIDE_KEPT_EMPTY",
                        "message": f"Source slide {src_idx} migrated as blank content",
                        "src_index": src_idx,
                    }
                )

    to_remove = [i for i, s in enumerate(prs.slides) if id(s.part) in initial_part_ids]
    for i in reversed(to_remove):
        delete_slide(prs, i)

    for slide in prs.slides:
        radial = slide in radial_slides
        bac.polish_slide(slide, skip_body_dedup=radial)
        if slide.slide_layout.name == LAYOUT_CONTENT:
            title = ""
            for sh in slide.placeholders:
                if sh.placeholder_format.idx == 10:
                    title = normalize_slide_title(sh.text or "")
            hide_empty_governing_placeholder(slide)
            if not radial:
                apply_academy_fonts_and_colors(slide)
                apply_academy_body_shape_typography(slide)
        if not radial:
            bac.fit_body_shapes(slide, sw, sh)
        if slide.slide_layout.name == LAYOUT_CONTENT and not radial:
            relayout_content_columns(slide, sw)
            t = ""
            for sh in slide.placeholders:
                if sh.placeholder_format.idx == 10:
                    t = normalize_slide_title(sh.text or "")
            apply_slide_title_layout(slide, sw, t)
        elif slide.slide_layout.name == LAYOUT_CONTENT and radial:
            t = ""
            for sh in slide.placeholders:
                if sh.placeholder_format.idx == 10:
                    t = normalize_slide_title(sh.text or "")
            apply_slide_title_layout(slide, sw, t)
            apply_academy_fonts_and_colors(slide)
            apply_academy_body_shape_typography(slide)

    from scripts.academy_deck_build_lib import fix_open_in_slide_view, polish_academy_presentation

    polish_academy_presentation(prs)
    for slide in prs.slides:
        if slide.slide_layout.name == LAYOUT_CONTENT:
            apply_academy_body_shape_typography(slide)
    slide_count = len(prs.slides)
    if slide_count != expected_slides:
        warnings.append(
            {
                "code": "SLIDE_COUNT_MISMATCH",
                "message": f"Expected {expected_slides} slides, got {slide_count}",
                "expected": expected_slides,
                "actual": slide_count,
                "deck_kind": cfg.deck_kind,
            }
        )

    save_academy_deck(prs, output)

    validate_script = PPT_TEST / "scripts" / "validate_pptx.py"
    if validate_script.is_file():
        import subprocess

        proc = subprocess.run(
            [sys.executable, str(validate_script), str(output)],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            raw = (proc.stderr or proc.stdout or "validate_pptx failed").strip()
            if "Traceback" in raw or "ImportError" in raw or "ModuleNotFoundError" in raw:
                msg = "validate_pptx_unavailable"
            else:
                msg = raw[:200]
            warnings.append(
                {
                    "code": "OOXML_VALIDATE_FAILED",
                    "message": msg,
                }
            )

    if not skip_ooxml_repair:
        from scripts.ooxml_repair_pptx import repair_pptx_in_place

        repair_pptx_in_place(output, template=template)
        fix_open_in_slide_view(output)
        from scripts.academy_deck_build_lib import finalize_academy_package

        finalize_academy_package(output)

    raw_output = output
    if not skip_powerpoint_repair:
        repaired_output = output.with_name(
            output.stem.replace("_아카데미적용_", "_아카데미적용_PP정리_") + ".pptx"
        )
        try:
            import platform

            if platform.system() == "Darwin":
                from scripts.powerpoint_repair_mac import powerpoint_repair_and_save

                powerpoint_repair_and_save(raw_output, repaired_output)
                output = repaired_output
        except Exception as exc:
            warnings.append(
                {"code": "PP_REPAIR_SKIPPED", "message": f"PowerPoint repair skipped: {exc}"}
            )

    warnings.insert(
        0,
        {
            "code": "MIGRATE_META",
            "deck_kind": cfg.deck_kind,
            "source_slides": len(src.slides),
            "plan_entries": len(plan),
            "expected_slides": expected_slides,
        },
    )
    return output, warnings, slide_count


def main() -> int:
    parser = argparse.ArgumentParser(description="Migrate CMP deck to academy template (§7)")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="Source .pptx")
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Output .pptx (default: source dir, stamped name)",
    )
    parser.add_argument("--template", type=Path, default=None, help="Academy template (else TEMPLATE_PPTX)")
    parser.add_argument(
        "--powerpoint-repair",
        action="store_true",
        help="Run macOS PowerPoint repair pass (slow)",
    )
    args = parser.parse_args()

    if not args.source.is_file():
        raise SystemExit(f"Missing source: {args.source}")

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out = args.out or args.source.with_name(
        f"{args.source.stem}_아카데미적용_{stamp}.pptx"
    )

    output, warnings, slide_count = migrate_cmp_deck(
        args.source,
        out,
        template_path=args.template,
        skip_powerpoint_repair=not args.powerpoint_repair,
    )
    print(f"Slides: {slide_count}")
    print(f"Saved: {output}")
    if warnings:
        for w in warnings:
            print(f"  warn: {w}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
