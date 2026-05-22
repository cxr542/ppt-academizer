#!/usr/bin/env python3
"""Extract text and assets from source .pptx (including Google Slides image slides)."""

from __future__ import annotations

from pathlib import Path

from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE

IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def iter_shapes(shapes, depth: int = 0):
    """Yield all shapes, recursing into groups."""
    for sh in shapes:
        yield sh, depth
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes, depth + 1)


def slide_text_blocks(slide) -> list[dict]:
    """All text-bearing shapes (including inside groups)."""
    blocks: list[dict] = []
    for sh, _depth in iter_shapes(slide.shapes):
        if not sh.has_text_frame:
            continue
        t = (sh.text or "").strip()
        if not t:
            continue
        blocks.append(
            {
                "top": int(sh.top) / 914400,
                "left": int(sh.left) / 914400,
                "text": t,
                "len": len(t),
            }
        )
    blocks.sort(key=lambda b: (b["top"], b["left"]))
    return blocks


def slide_notes_text(slide) -> str:
    """Speaker notes plain text (if any)."""
    try:
        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        return (tf.text or "").strip()
    except Exception:
        return ""


def is_image_only_slide(slide) -> bool:
    """Google Slides .pptx export: full-slide background picture, no shapes."""
    if slide_text_blocks(slide):
        return False
    if len(slide.shapes) > 0:
        # Allow only empty group shell
        has_real = False
        for sh, _ in iter_shapes(slide.shapes):
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                continue
            has_real = True
            break
        if has_real:
            return False
    try:
        return slide.background.fill.type == MSO_FILL.PICTURE
    except Exception:
        return False


def slide_background_image_part(slide):
    """Return ImagePart for slide background, or None."""
    for rel in slide.part.rels.values():
        if rel.reltype == IMAGE_REL and rel.target_ref.startswith("../media/"):
            # Prefer background blip (first image rel is background for Google export)
            return rel.target_part
    return None


def export_slide_background_image(slide, dest: Path) -> Path | None:
    """Write slide background image to ``dest``; return path or None."""
    part = slide_background_image_part(slide)
    if part is None or not getattr(part, "blob", None):
        return None
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    ext = ".png"
    ref = ""
    for rel in slide.part.rels.values():
        if rel.reltype == IMAGE_REL:
            ref = rel.target_ref
            break
    if ref.lower().endswith(".jpg") or ref.lower().endswith(".jpeg"):
        ext = ".jpg"
    if dest.suffix.lower() not in (".png", ".jpg", ".jpeg"):
        dest = dest.with_suffix(ext)
    dest.write_bytes(part.blob)
    return dest
