from __future__ import annotations
import json
import os
from typing import Any
from pydantic import BaseModel, Field

from scripts.pptx_ingest import slide_text_blocks

class SlideLayoutAnalysis(BaseModel):
    layout: str = Field(description="레이아웃 종류")
    title: str = Field(default="", description="슬라이드의 주요 제목")
    governing: str = Field(default="", description="핵심 요약 메시지")
    kicker: str = Field(default="", description="작은 소제목")

def analyze_slide_with_llm(slide, slide_index: int = -1) -> dict[str, Any]:
    """임시로 LLM 대신 V2 휴리스틱을 사용합니다 (서버 타임아웃 방지)"""
    blocks = slide_text_blocks(slide)
    if not blocks:
        return {
            "layout": "empty",
            "title": "",
            "governing": "",
            "kicker": "",
            "blocks": [],
        }

    from scripts.slide_classifier_v2 import analyze_slide_layout
    fallback_info = analyze_slide_layout(slide)
    
    return {
        "layout": fallback_info.get("layout", "content"),
        "title": fallback_info.get("title", ""),
        "governing": fallback_info.get("governing", ""),
        "kicker": fallback_info.get("kicker", ""),
        "blocks": blocks,
    }
