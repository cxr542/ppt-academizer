"""A title + table slide is content, not a cover (covers drop everything but the title)."""

from __future__ import annotations

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

from scripts.slide_classifier_v2 import analyze_slide_layout, has_data_object


def _titled_slide(prs, text="구분별 비교"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    box.text_frame.text = text
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(40)  # >= 32pt: the cover heuristic
    return slide


def test_title_only_slide_is_still_a_cover() -> None:
    prs = Presentation()
    slide = _titled_slide(prs, "클라우드 구현기술")
    assert analyze_slide_layout(slide)["layout"] == "cover"
    assert not has_data_object(slide)


def test_title_plus_table_is_content_not_cover() -> None:
    prs = Presentation()
    slide = _titled_slide(prs)
    slide.shapes.add_table(4, 3, Inches(1), Inches(2), Inches(8), Inches(3))
    assert has_data_object(slide)
    assert analyze_slide_layout(slide)["layout"] == "content"


def test_title_plus_chart_is_content_not_cover() -> None:
    prs = Presentation()
    slide = _titled_slide(prs)
    data = CategoryChartData()
    data.categories = ["A", "B"]
    data.add_series("S", (1, 2))
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(6), Inches(4), data)
    assert analyze_slide_layout(slide)["layout"] == "content"


def test_table_inside_group_is_detected() -> None:
    prs = Presentation()
    slide = _titled_slide(prs)
    grp = slide.shapes.add_group_shape()
    grp.shapes.add_shape(1, Inches(1), Inches(2), Inches(1), Inches(1))  # rectangle child
    assert not has_data_object(slide)


def _slide_with_blocks(prs, title, others):
    slide = _titled_slide(prs, title)
    for i, text in enumerate(others):
        box = slide.shapes.add_textbox(Inches(1), Inches(2 + i * 1.2), Inches(8), Inches(1))
        box.text_frame.text = text
        box.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
    return slide


def test_title_with_short_subtitle_is_a_cover() -> None:
    prs = Presentation()
    slide = _slide_with_blocks(prs, "THANK YOU", ["경청해 주셔서 감사합니다", "contact@example.com"])
    assert analyze_slide_layout(slide)["layout"] == "cover"


def test_title_with_card_text_is_content_not_cover() -> None:
    """Content slides with a few text cards were migrated as covers and lost their body text."""
    prs = Presentation()
    cards = [
        "독자적 기술력: 외산 솔루션에 의존하지 않는 순수 국산 기술로 개발되었습니다",
        "안정적 운영: 대규모 데이터센터 환경에서 검증된 가용성을 제공합니다",
        "유연한 확장성: 온프레미스와 클라우드를 아우르는 구조입니다",
    ]
    slide = _slide_with_blocks(prs, "콘트라베이스의 강점", cards)
    assert analyze_slide_layout(slide)["layout"] == "content"


def test_body_heavy_flag_for_first_slide_decisions() -> None:
    prs = Presentation()
    plain = _titled_slide(prs, "클라우드 구현기술")
    assert analyze_slide_layout(plain)["body_heavy"] is False

    with_table = _titled_slide(prs)
    with_table.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(6), Inches(2))
    assert analyze_slide_layout(with_table)["body_heavy"] is True

    # A table-only slide has no text blocks at all, yet is body content.
    only_table = prs.slides.add_slide(prs.slide_layouts[6])
    only_table.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(6), Inches(2))
    from scripts.llm_classifier import analyze_slide_with_llm

    assert analyze_slide_with_llm(only_table)["body_heavy"] is True
