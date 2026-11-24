from __future__ import annotations

from typing import Any

from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt

from scripts.pptx_ingest import slide_text_blocks


# Covers and section dividers are migrated as title text only, so everything else on the slide
# is dropped. Only treat a slide as one when its non-title text is this short; otherwise fall back
# to "content", which keeps every text box and picture.
COVER_MAX_BODY_CHARS = 80

# A cover title is >= 32pt. ``slide_text_blocks`` reports font sizes in EMU (Pt(32) == 406400),
# but this was compared with 3200 (centipoints), so ANY explicit font size passed and every
# slide with <= 4 text blocks scored as a cover.
COVER_TITLE_MIN_EMU = int(Pt(32))

# A title that starts in the top header band (inches from the top) is a content-slide title:
# real cover titles sit lower on the slide. Without this, a content slide holding a chart
# *picture* plus a caption (<= 4 text blocks, 32pt+ title, nothing for ``has_data_object``)
# scored as a cover and lost its pictures. Misreading a cover as content keeps everything,
# so this errs on the safe side; slide 0 is still forced to be a cover by the plan builder.
COVER_TITLE_MIN_TOP_IN = 1.0


def has_data_object(slide) -> bool:
    """True when the slide carries a table, chart, SmartArt, OLE object or any other graphic
    frame (also inside groups).

    ``slide_text_blocks`` only sees text frames, so a slide whose body is a table looks like
    "a big title and almost no text" and was classified as a cover. A cover is migrated as
    title text only, which silently dropped the table and its content.
    """

    def _walk(shapes) -> bool:
        for sh in shapes:
            if getattr(sh, "has_table", False) and sh.has_table:
                return True
            if getattr(sh, "has_chart", False) and sh.has_chart:
                return True
            # SmartArt (dgm), embedded OLE, ... are <p:graphicFrame>s python-pptx does not expose
            # through has_table/has_chart.
            if str(getattr(sh, "_element", None).tag).endswith("}graphicFrame"):
                return True
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP and _walk(sh.shapes):
                return True
        return False

    return _walk(slide.shapes)


_P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
DARK_BACKGROUND_MAX_LUMA = 0.4


def has_dark_background(slide) -> bool:
    """True when the slide's own background (else its layout's) is dark: a solid or gradient fill
    whose explicit RGB colors average darker than 0.4 luma. Theme-colored fills are not judged."""
    for part in (slide, slide.slide_layout):
        bg = part._element.find(f".//{_P_NS}bg")
        if bg is None:
            continue
        lumas = []
        for clr in bg.iter(f"{_A_NS}srgbClr"):
            v = clr.get("val") or ""
            if len(v) == 6:
                r, g, b = (int(v[i : i + 2], 16) for i in (0, 2, 4))
                lumas.append((0.299 * r + 0.587 * g + 0.114 * b) / 255)
        return bool(lumas) and sum(lumas) / len(lumas) < DARK_BACKGROUND_MAX_LUMA
    return False


COVER_DIAGRAM_MIN_BOX_EMU = int(0.4 * 914400)
COVER_DIAGRAM_MAX_CHARS = 160


def _has_solid_fill(shape) -> bool:
    try:
        return shape.fill.type == MSO_FILL.SOLID
    except Exception:
        return False


def cover_diagram_shapes(slide) -> list:
    """Shapes of a "before -> after" style diagram on a cover: two or more filled boxes (at
    least 0.4in each way, not a full-slide background) and everything whose center lies in the
    area they span (their labels, the arrow between them). Empty when there is no such diagram."""
    prs = slide.part.package.presentation_part.presentation
    sw, sh_h = int(prs.slide_width), int(prs.slide_height)
    fills = [
        sh
        for sh in slide.shapes
        if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
        and not (sh.has_text_frame and sh.text_frame.text.strip())
        and _has_solid_fill(sh)
        and int(sh.width or 0) >= COVER_DIAGRAM_MIN_BOX_EMU
        and int(sh.height or 0) >= COVER_DIAGRAM_MIN_BOX_EMU
        and not (int(sh.width) >= 0.9 * sw and int(sh.height) >= 0.9 * sh_h)
    ]
    if len(fills) < 2:
        return []
    left = min(int(f.left) for f in fills)
    top = min(int(f.top) for f in fills)
    right = max(int(f.left) + int(f.width) for f in fills)
    bottom = max(int(f.top) + int(f.height) for f in fills)
    found = []
    for sh in slide.shapes:
        cx = int(sh.left or 0) + int(sh.width or 0) // 2
        cy = int(sh.top or 0) + int(sh.height or 0) // 2
        if left <= cx <= right and top <= cy <= bottom:
            found.append(sh)
    return found


def _norm_text(text: str) -> str:
    return "".join(text.split())


def _is_cover_with_diagram(slide, blocks: list[dict], data_object: bool) -> bool:
    """A cover whose body is one small diagram (boxes with short labels) around cover-like
    text: the diagram can be kept on the cover instead of dropping it."""
    if data_object or not blocks:
        return False
    diagram = cover_diagram_shapes(slide)
    if not diagram:
        return False
    labels = {
        _norm_text(sh.text_frame.text)
        for sh in diagram
        if sh.has_text_frame and sh.text_frame.text.strip()
    }
    if sum(len(t) for t in labels) > COVER_DIAGRAM_MAX_CHARS:
        return False
    rest = [b for b in blocks if _norm_text(b["text"]) not in labels]
    return _is_cover_meta_only(rest, False)


class TextBand:
    def __init__(self, blocks: list[dict]):
        self.blocks = blocks
        self.top = min(b["top"] for b in blocks) if blocks else 0
        self.left = min(b["left"] for b in blocks) if blocks else 0
        self.bottom = max(b["top"] + b["height"] for b in blocks) if blocks else 0
        self.right = max(b["left"] + b["width"] for b in blocks) if blocks else 0
        self.text = "\n".join(b["text"] for b in blocks)
        self.font_sizes = [b["font_size"] for b in blocks if b.get("font_size")]
        self.avg_font_size = sum(self.font_sizes) / len(self.font_sizes) if self.font_sizes else 0
        self.is_bold = any(b.get("is_bold") for b in blocks)
        self.is_placeholder = any(b.get("is_placeholder") for b in blocks)

def _is_stacked(band: TextBand) -> bool:
    """True when the blocks of a band read as one message (a single box, or boxes stacked on
    top of each other). Boxes side by side are separate columns/labels, not a governing line."""
    if len(band.blocks) < 2:
        return True
    a, b = band.blocks[0], band.blocks[1]
    overlap = min(a["left"] + a["width"], b["left"] + b["width"]) - max(a["left"], b["left"])
    narrower = min(a["width"], b["width"])
    return narrower > 0 and overlap >= 0.5 * narrower


# A governing message is one short sentence under the title. Anything longer or made of several
# paragraphs is body text (a card heading plus its bullets), and must stay in the body.
GOVERNING_MAX_CHARS = 120


def _looks_like_body(band: TextBand) -> bool:
    """True when a band is a list of points or a long text, not a one-line governing message."""
    for b in band.blocks:
        paragraphs = [p for p in (b["text"] or "").split("\n") if p.strip()]
        if len(paragraphs) >= 2:
            return True
    return len(band.text.replace("\n", " ").strip()) > GOVERNING_MAX_CHARS


def cluster_bands(blocks: list[dict], y_threshold: float = 0.15) -> list[TextBand]:
    """Group blocks by vertical proximity (Y-threshold in inches)."""
    if not blocks:
        return []
    
    # Sort top-down
    sorted_blocks = sorted(blocks, key=lambda b: b["top"])
    
    bands: list[list[dict]] = []
    current_band: list[dict] = [sorted_blocks[0]]
    
    for b in sorted_blocks[1:]:
        prev = current_band[-1]
        gap = b["top"] - (prev["top"] + prev["height"])
        
        prev_font = prev.get("font_size") or 100000
        curr_font = b.get("font_size") or 100000
        font_diff = abs(prev_font - curr_font) / max(prev_font, 1)
        
        if gap < y_threshold and font_diff < 0.3:
            current_band.append(b)
        else:
            bands.append(current_band)
            current_band = [b]
    
    if current_band:
        bands.append(current_band)
    
    return [TextBand(band) for band in bands]

COVER_META_MAX_BLOCKS = 5
COVER_META_MAX_CHARS = 80


def _is_cover_meta_only(blocks: list[dict], data_object: bool) -> bool:
    """Cover-like text: a few single-paragraph blocks, none of them a long passage."""
    if data_object or not blocks or len(blocks) > COVER_META_MAX_BLOCKS:
        return False
    return all(
        "\n" not in b["text"].strip() and len(b["text"].strip()) <= COVER_META_MAX_CHARS
        for b in blocks
    )


COVER_FOOTER_MIN_TOP_RATIO = 0.8


def split_cover_roles(blocks: list[dict], slide_height_in: float) -> dict:
    """Roles of the text on a cover: the largest text is the title; lines above it are the
    eyebrow, lines at the very bottom of the slide are the source/date footer, the rest below
    the title is the subtitle. Each group is in reading order (top to bottom)."""
    live = [b for b in blocks if b["text"].strip()]
    if not live:
        return {"title": None, "eyebrow": [], "subtitle": [], "footer": []}
    title = max(live, key=lambda b: (b.get("font_size") or 0, -b["top"]))
    rest = [b for b in live if b is not title]
    footer = [
        b for b in rest
        if b["top"] > title["top"] and b["top"] >= COVER_FOOTER_MIN_TOP_RATIO * slide_height_in
    ]
    eyebrow = [b for b in rest if b["top"] < title["top"] and b not in footer]
    subtitle = [b for b in rest if b not in footer and b not in eyebrow]

    def texts(group):
        return [b["text"].strip() for b in sorted(group, key=lambda b: (b["top"], b["left"]))]

    return {
        "title": title["text"].strip(),
        "eyebrow": texts(eyebrow),
        "subtitle": texts(subtitle),
        "footer": texts(footer),
    }


def analyze_slide_layout(slide) -> dict[str, Any]:
    """Extract semantic bands and compute scores for layouts."""
    blocks = slide_text_blocks(slide)
    bands = cluster_bands(blocks)
    
    # Identify roles
    title_band: TextBand | None = None
    governing_band: TextBand | None = None
    content_bands: list[TextBand] = []
    
    if bands:
        top_bands = [b for b in bands if b.top < 2.0]
        if top_bands:
            title_band = max(top_bands, key=lambda b: (b.avg_font_size, -b.top))
            
            remaining = [b for b in bands if b != title_band and b.top >= title_band.top]
            if remaining:
                candidate = remaining[0]
                if (
                    candidate.top < 3.0
                    and len(candidate.blocks) <= 2
                    and _is_stacked(candidate)
                    and not _looks_like_body(candidate)
                ):
                    governing_band = candidate
                    remaining = remaining[1:]
            
            # Content bands are everything else (including things above title like kickers)
            content_bands = [b for b in bands if b != title_band and b != governing_band]
        else:
            content_bands = bands

    title_text = title_band.text if title_band else None
    gov_text = governing_band.text if governing_band else None

    # If title and governing are clumped in a single multi-line block (e.g. title\ngoverning)
    if title_text and "\n" in title_text and not gov_text:
        lines = title_text.split("\n", 1)
        if len(lines[0]) < 80: # likely a title
            title_text = lines[0].strip()
            gov_text = lines[1].strip()

    scores = {
        "cover": 0.0,
        "toc": 0.0,
        "section": 0.0,
        "content": 0.0,
        "empty": 100.0 if not blocks else 0.0,
    }

    divider_band: TextBand | None = None
    if blocks:
        if (
            len(blocks) <= 4
            and title_band
            and title_band.avg_font_size >= COVER_TITLE_MIN_EMU
            and title_band.top >= COVER_TITLE_MIN_TOP_IN
        ):
            scores["cover"] += 50
            if not content_bands:
                scores["cover"] += 30
                
        if title_text and ("목차" in title_text or "CONTENTS" in title_text.upper()):
            scores["toc"] += 80
        else:
            num_lines = sum(1 for b in blocks if b["text"].strip().startswith(("1.", "2.", "3.", "1)", "2)", "3)")))
            if num_lines >= 2:
                scores["toc"] += 40

        # A divider's title sits mid-slide, so it is never in ``title_band`` (top < 2.0in only);
        # judge it by the largest band instead. A short subtitle under it is fine: real body
        # text is caught by the ``body_heavy`` guard below.
        divider_band = title_band
        if divider_band is None and bands:
            divider_band = max(bands, key=lambda b: (b.avg_font_size, -b.top))
            divider_band = divider_band if divider_band.avg_font_size >= COVER_TITLE_MIN_EMU else None
        if len(blocks) <= 3 and divider_band and divider_band.top >= 2.0:
            scores["section"] += 60
        elif (
            len(blocks) == 1
            and len(slide.shapes) <= 2
            and not has_data_object(slide)
            and not any(sh.shape_type == MSO_SHAPE_TYPE.PICTURE for sh in slide.shapes)
            and has_dark_background(slide)
        ):
            # A slide that holds nothing but one title line on a dark background is a divider
            # wherever the title sits (a top-left title on a gradient is a common divider
            # design). On a white slide a lone header title stays "content".
            scores["section"] += 60

        scores["content"] = 40
        if content_bands and len(blocks) > 3:
            scores["content"] += len(blocks) * 2

    body_chars = 0
    lead_band = title_band or divider_band
    if blocks and lead_band is not None:
        title_ids = {id(b) for b in lead_band.blocks}
        body_chars = sum(len(b["text"]) for b in blocks if id(b) not in title_ids)
    data_object = has_data_object(slide)
    body_heavy = data_object or (bool(blocks) and body_chars > COVER_MAX_BODY_CHARS)
    if blocks and body_heavy:
        # Tables/charts or a body of real text: never a cover or a section divider, because
        # those render the title only and would drop the rest of the slide.
        scores["cover"] = 0.0
        scores["section"] = 0.0
        if scores["toc"] <= 40:
            scores["content"] = max(scores["content"], 41)

    best_layout = max(scores.items(), key=lambda x: x[1])[0]
    if best_layout == "section" and not title_text and divider_band is not None:
        title_text = divider_band.text

    return {
        "layout": best_layout,
        "scores": scores,
        "title": title_text,
        "governing": gov_text,
        "blocks": blocks,
        # True when a cover/section rendering (title only) would drop real content.
        "body_heavy": body_heavy,
        # True when the "body" is only short one-line meta text (eyebrow, subtitle, source
        # line) around a title, as on a real cover; ``cover_lines`` keeps all of it.
        "cover_meta_only": _is_cover_meta_only(blocks, data_object),
        # True when a cover also carries one small diagram (boxes with short labels).
        "cover_diagram": _is_cover_with_diagram(slide, blocks, data_object),
    }
