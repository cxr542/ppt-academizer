"""A title + table slide is content, not a cover (covers drop everything but the title)."""

from __future__ import annotations

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

from scripts.slide_classifier_v2 import analyze_slide_layout, has_data_object


# Real cover titles sit mid-slide; content titles sit in the top header band.
COVER_TOP_IN = 1.6
CONTENT_TOP_IN = 0.6


def _titled_slide(prs, text="구분별 비교", top=COVER_TOP_IN):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(top), Inches(8), Inches(1))
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
        box = slide.shapes.add_textbox(Inches(1), Inches(3.4 + i * 1.2), Inches(8), Inches(1))
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


def test_cover_needs_a_32pt_title_not_just_any_explicit_font_size() -> None:
    """The threshold compared EMU with 3200, so a 30pt (even 12pt) title qualified as a cover."""
    prs = Presentation()
    for pt, expected in ((12, "cover"), (30, "cover"), (32, "cover"), (44, "cover")):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(COVER_TOP_IN), Inches(8), Inches(1))
        box.text_frame.text = "배포 파이프라인"
        box.text_frame.paragraphs[0].runs[0].font.size = Pt(pt)
        layout = analyze_slide_layout(slide)["layout"]
        assert (layout == "cover") == (pt >= 32), (pt, layout)


def test_header_title_with_chart_pictures_is_content_not_cover(tmp_path) -> None:
    """Chart pasted as pictures + caption under a top-of-slide 33pt title was read as a cover
    and its pictures were dropped (real deck: VMware win-back report, slide 6)."""
    from base64 import b64decode

    png = tmp_path / "chart.png"
    png.write_bytes(
        b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAEklEQVR4nGP4z8AAAAMBAQDJ"
            "/pLvAAAAAElFTkSuQmCC"
        )
    )
    prs = Presentation()
    slide = _titled_slide(prs, "주요 성장 지표 추이 (2023-2025)", top=CONTENT_TOP_IN)
    slide.shapes[0].text_frame.paragraphs[0].runs[0].font.size = Pt(33)
    note = slide.shapes.add_textbox(Inches(1.5), Inches(3.6), Inches(7), Inches(0.3))
    note.text_frame.text = "※ 2023년 대비 2025년 결산 기준의 폭발적 성장세 증명"
    note.text_frame.paragraphs[0].runs[0].font.size = Pt(15)
    slide.shapes.add_picture(str(png), Inches(1), Inches(1.6), Inches(4), Inches(2))
    slide.shapes.add_picture(str(png), Inches(5.2), Inches(1.6), Inches(4), Inches(2))

    result = analyze_slide_layout(slide)
    assert result["layout"] == "content"
    assert result["scores"]["cover"] == 0.0


def test_mid_slide_title_still_scores_as_cover() -> None:
    prs = Presentation()
    slide = _titled_slide(prs, "경청해 주셔서 감사합니다", top=COVER_TOP_IN)
    assert analyze_slide_layout(slide)["layout"] == "cover"


def _divider(prs, title, kicker=None, subtitle=None, title_top=2.3, title_pt=54):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if kicker:
        k = slide.shapes.add_textbox(Inches(1), Inches(title_top - 0.8), Inches(8), Inches(0.4))
        k.text_frame.text = kicker
        k.text_frame.paragraphs[0].runs[0].font.size = Pt(15)
    box = slide.shapes.add_textbox(Inches(1), Inches(title_top), Inches(10), Inches(1.2))
    box.text_frame.text = title
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(title_pt)
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(1), Inches(title_top + 1.5), Inches(10), Inches(0.5))
        sub.text_frame.text = subtitle
        sub.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
    return slide


def test_mid_slide_title_only_divider_is_a_section() -> None:
    """Section detection was unreachable: the title band only came from the top 2in, but the
    section rule needed the title at >= 2in, so no slide ever scored as a divider."""
    prs = Presentation()
    result = analyze_slide_layout(_divider(prs, "I. 시장 환경 및 배경"))
    assert result["layout"] == "section"
    assert result["title"] == "I. 시장 환경 및 배경"


def test_divider_with_one_line_subtitle_is_a_section() -> None:
    prs = Presentation()
    slide = _divider(prs, "I. 시장 환경 및 배경", subtitle="글로벌 가상화 시장의 급격한 정책 변화")
    assert analyze_slide_layout(slide)["layout"] == "section"


def test_divider_with_kicker_uses_the_big_line_as_title() -> None:
    prs = Presentation()
    result = analyze_slide_layout(_divider(prs, "전략적 참여 배경", kicker="STRATEGIC ANALYSIS", title_top=4.0))
    assert result["layout"] == "section"
    assert result["title"] == "전략적 참여 배경"


def test_divider_like_slide_with_real_body_text_is_content() -> None:
    prs = Presentation()
    body = "이 슬라이드는 간지처럼 보이지만 실제 본문 설명이 길게 이어지는 내용 슬라이드입니다. " * 2
    assert analyze_slide_layout(_divider(prs, "I. 시장 환경", subtitle=body))["layout"] != "section"


def test_small_mid_slide_text_is_not_a_divider() -> None:
    prs = Presentation()
    assert analyze_slide_layout(_divider(prs, "참고 자료", title_pt=18))["layout"] == "content"


def test_header_title_slide_is_not_a_section() -> None:
    prs = Presentation()
    result = analyze_slide_layout(_divider(prs, "시장 환경", title_top=0.6))
    assert result["layout"] != "section"


def test_slide_without_text_blocks_can_be_analyzed_directly() -> None:
    """analyze_slide_layout must not depend on callers short-circuiting empty slides."""
    prs = Presentation()
    empty = prs.slides.add_slide(prs.slide_layouts[6])
    result = analyze_slide_layout(empty)
    assert result["layout"] == "empty"
    assert result["scores"]["section"] == 0.0


def _bar_slide(top_in):
    from pptx.enum.shapes import MSO_SHAPE

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.62), Inches(top_in), Inches(0.08), Inches(0.67))
    return bar


def test_title_accent_bar_in_the_header_band_is_not_migrated() -> None:
    """The source's gold bar beside its title landed on the governing-message placeholder."""
    from scripts.shape_migrate_lib import is_body_shape

    assert not is_body_shape(_bar_slide(0.62), "제목", None)


def test_accent_bar_below_the_header_band_is_still_migrated() -> None:
    from scripts.shape_migrate_lib import is_body_shape

    assert is_body_shape(_bar_slide(3.0), "제목", None)
