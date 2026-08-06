from __future__ import annotations

from typing import Any
from scripts.pptx_ingest import slide_text_blocks

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
                if candidate.top < 3.0 and len(candidate.blocks) <= 2:
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

    if blocks:
        if len(blocks) <= 4 and title_band and title_band.avg_font_size >= 3200:
            scores["cover"] += 50
            if not content_bands:
                scores["cover"] += 30
                
        if title_text and ("목차" in title_text or "CONTENTS" in title_text.upper()):
            scores["toc"] += 80
        else:
            num_lines = sum(1 for b in blocks if b["text"].strip().startswith(("1.", "2.", "3.", "1)", "2)", "3)")))
            if num_lines >= 2:
                scores["toc"] += 40

        if len(blocks) <= 3 and title_band and title_band.top >= 2.0 and not content_bands:
            scores["section"] += 60

        scores["content"] = 40
        if content_bands and len(blocks) > 3:
            scores["content"] += len(blocks) * 2

    best_layout = max(scores.items(), key=lambda x: x[1])[0]

    return {
        "layout": best_layout,
        "scores": scores,
        "title": title_text,
        "governing": gov_text,
        "blocks": blocks,
    }
