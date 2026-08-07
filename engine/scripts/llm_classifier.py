from __future__ import annotations

import json
import os
from typing import Any

from pydantic import BaseModel, Field
from google import genai
from google.genai import types

from scripts.pptx_ingest import slide_text_blocks

class SlideLayoutAnalysis(BaseModel):
    layout: str = Field(
        description="레이아웃 종류. 다음 중 하나여야 합니다: 'cover' (표지), 'toc' (목차), 'section' (간지/섹션), 'content' (본문), 'empty' (빈 슬라이드)"
    )
    title: str = Field(
        default="",
        description="슬라이드의 주요 제목(Title) 텍스트. (예: '1. 기술 개요'). Kicker(작은 소제목)는 제외하고 본 제목만 기재하세요."
    )
    governing: str = Field(
        default="",
        description="본문 슬라이드의 상단에 위치한 두 줄 이하의 핵심 요약 메시지(Governing Message). 없으면 빈 문자열을 반환하세요."
    )
    kicker: str = Field(
        default="",
        description="제목 상단에 위치하는 작은 분류명이나 부제목 (예: 'TECHNICAL', 'OVERVIEW'). 없으면 빈 문자열."
    )

def _get_gemini_client() -> genai.Client | None:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        return None
    return genai.Client(api_key=api_key)

def analyze_slide_with_llm(slide, slide_index: int = -1) -> dict[str, Any]:
    """LLM을 사용하여 슬라이드의 구조를 완벽하게 파악합니다."""
    blocks = slide_text_blocks(slide)
    if not blocks:
        return {
            "layout": "empty",
            "title": "",
            "governing": "",
            "kicker": "",
            "blocks": [],
        }

    client = _get_gemini_client()
    if not client:
        # Fallback to a basic heuristic if no API key is set
        # This shouldn't normally happen in production if we enforce it.
        return {
            "layout": "content",
            "title": blocks[0]["text"] if blocks else "",
            "governing": "",
            "kicker": "",
            "blocks": blocks,
        }

    prompt = f"""다음은 파워포인트 슬라이드의 텍스트 블록 목록입니다. (인덱스 {slide_index})
각 블록은 텍스트 내용, 좌표(좌상단 x, y), 크기(width, height), 폰트 크기(font_size), 굵게(is_bold) 정보를 포함합니다.
이 정보들을 바탕으로 슬라이드의 전체적인 레이아웃(표지, 목차, 간지, 본문)을 파악하고,
만약 본문(content)이라면 Kicker(작은 소제목), Title(주요 제목), Governing Message(상단 핵심 요약 1~2줄)를 분리해 주세요.

규칙:
1. 'title'은 슬라이드의 가장 핵심이 되는 제목(예: "1. 서론", "아키텍처 구성")입니다.
2. 'kicker'는 보통 제목 위나 좌측 상단에 작게 적힌 분류명(예: "OVERVIEW", "TECHNICAL")입니다.
3. 'governing'은 제목 바로 아래(또는 본문 최상단)에 위치하여 슬라이드 전체를 요약하는 1~2줄의 메시지(예: "본 솔루션은 높은 가용성을 제공합니다.")입니다.
4. 만약 하나의 텍스트 블록에 Kicker와 Title이 합쳐져 있거나, Title과 Governing이 합쳐져 있다면 이를 문맥상 분리하여 반환하세요.
5. 목차 슬라이드는 'toc', 중간 간지는 'section', 첫 장은 'cover'로 분류하세요.

텍스트 블록 정보:
{json.dumps(blocks, ensure_ascii=False, indent=2)}
"""

    try:
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=SlideLayoutAnalysis,
                temperature=0.0,
            ),
        )
        
        # 응답이 마크다운 json 블록으로 감싸져 있을 수 있으므로 처리
        text = response.text.strip()
        if text.startswith("```json"):
            text = text[7:]
        if text.startswith("```"):
            text = text[3:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()

        data = json.loads(text)
        analysis = SlideLayoutAnalysis.model_validate(data)
        
        return {
            "layout": analysis.layout,
            "title": analysis.title,
            "governing": analysis.governing,
            "kicker": analysis.kicker,
            "blocks": blocks,
        }
    except Exception as e:
        import traceback
        print(f"LLM Classification Error on slide {slide_index}: {traceback.format_exc()}")
        
        # LLM 실패 시 기존 V2 휴리스틱(slide_classifier_v2)으로 안전하게 폴백
        from scripts.slide_classifier_v2 import analyze_slide_layout
        fallback_info = analyze_slide_layout(slide)
        
        return {
            "layout": fallback_info.get("layout", "content"),
            "title": fallback_info.get("title", ""),
            "governing": fallback_info.get("governing", ""),
            "kicker": "",  # V2 에서는 kicker 필드가 명시적으로 분리되지 않음
            "blocks": blocks,
        }
