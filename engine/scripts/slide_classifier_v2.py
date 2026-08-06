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

def cluster_bands(blocks: list[dict], y_threshold: float = 0.5) -> list[TextBand]:
    """Group blocks by vertical proximity (Y-threshold in inches)."""
    if not blocks:
        return []
    
    # Sort top-down
    sorted_blocks = sorted(blocks, key=lambda b: b["top"])
    
    bands: list[list[dict]] = []
    current_band: list[dict] = [sorted_blocks[0]]
    
    for b in sorted_blocks[1:]:
        prev = current_band[-1]
        # If the vertical gap is small, cluster them
        gap = b["top"] - (prev["top"] + prev["height"])
        if gap < y_threshold:
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
        # Title is usually the first band, or the largest text near the top
        top_bands = [b for b in bands if b.top < 2.0]
        if top_bands:
            title_band = max(top_bands, key=lambda b: (b.avg_font_size, -b.top))
            
            # Governing is usually right below title
            remaining = [b for b in bands if b != title_band]
            if remaining:
                candidate = remaining[0]
                if candidate.top < 3.0 and len(candidate.blocks) <= 2:
                    governing_band = candidate
                    remaining = remaining[1:]
            
            content_bands = remaining
        else:
            content_bands = bands

    # Scoring
    scores = {
        "cover": 0.0,
        "toc": 0.0,
        "section": 0.0,
        "content": 0.0,
        "empty": 100.0 if not blocks else 0.0,
    }

    if blocks:
        # Cover heuristic
        if len(blocks) <= 4 and title_band and title_band.avg_font_size >= 3200: # approx 32pt
            scores["cover"] += 50
            if not content_bands:
                scores["cover"] += 30
                
        # TOC heuristic
        if title_band and ("목차" in title_band.text or "CONTENTS" in title_band.text.upper()):
            scores["toc"] += 80
        else:
            # check density of numbers in content
            num_lines = sum(1 for b in blocks if b["text"].strip().startswith(("1.", "2.", "3.", "1)", "2)", "3)")))
            if num_lines >= 2:
                scores["toc"] += 40

        # Section heuristic
        if len(blocks) <= 3 and title_band and title_band.top >= 2.0 and not content_bands:
            scores["section"] += 60

        # Content heuristic (default)
        scores["content"] = 40
        if content_bands and len(blocks) > 3:
            scores["content"] += len(blocks) * 2

    # Determine best
    best_layout = max(scores.items(), key=lambda x: x[1])[0]

    return {
        "layout": best_layout,
        "scores": scores,
        "title": title_band.text if title_band else None,
        "governing": governing_band.text if governing_band else None,
        "blocks": blocks,
    }
