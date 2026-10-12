"""Title vs body mapping for spec ingest."""

from __future__ import annotations

from scripts.convert_legacy_deck_to_academy import _pick_title_and_body_from_blocks


def test_long_left_block_splits_to_body() -> None:
    usable = [
        {
            "top": 0.9,
            "left": 0.5,
            "text": "짧은 제목\n• 본문 첫 줄\n• 본문 둘째 줄",
            "len": 30,
        },
        {
            "top": 2.5,
            "left": 0.5,
            "text": "아래 단락 전체",
            "len": 8,
        },
    ]
    title, body, _sub = _pick_title_and_body_from_blocks(usable)
    assert title == "짧은 제목"
    assert "본문 첫 줄" in body
    assert "아래 단락" in body


def test_shortest_header_line_is_title() -> None:
    usable = [
        {
            "top": 1.0,
            "left": 0.2,
            "text": "슬라이드 제목",
            "len": 6,
        },
        {
            "top": 1.1,
            "left": 4.0,
            "text": "이건 옆 칼럼 본문이 아니라 제목 밴드에 가까운 짧은 라벨",
            "len": 20,
        },
        {
            "top": 3.0,
            "left": 0.5,
            "text": "실제 본문 단락입니다.",
            "len": 12,
        },
    ]
    title, body, _ = _pick_title_and_body_from_blocks(usable)
    assert title == "슬라이드 제목"
    assert "실제 본문" in body


def test_far_left_long_text_not_title() -> None:
    long_body = "A" * 50 + " 본문 시작"
    usable = [
        {"top": 0.95, "left": 0.15, "text": long_body, "len": len(long_body)},
        {"top": 1.0, "left": 3.5, "text": "실제 슬라이드 제목", "len": 10},
        {"top": 2.8, "left": 0.5, "text": "하단 본문", "len": 5},
    ]
    title, body, _ = _pick_title_and_body_from_blocks(usable)
    assert title == "실제 슬라이드 제목"
    assert long_body in body or "본문 시작" in body
