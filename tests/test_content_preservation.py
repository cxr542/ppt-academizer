"""Silent content loss guards: body placeholders, long governing text, CONTENT_LOSS warning."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches

from core.warnings_display import format_warning
from scripts.build_cmp_academy import content_loss_warning, slide_text_chars
from scripts.shape_migrate_lib import is_body_shape


def _title_and_content_slide(body: str):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
    slide.shapes.title.text = "쿠버네티스 핵심 개념"
    slide.placeholders[1].text_frame.text = body
    return slide


def test_body_placeholder_with_text_is_kept_title_is_not() -> None:
    slide = _title_and_content_slide("Pod는 배포의 최소 단위입니다")
    assert not is_body_shape(slide.shapes.title, "쿠버네티스 핵심 개념", None)
    assert is_body_shape(slide.placeholders[1], "쿠버네티스 핵심 개념", None)


def test_empty_body_placeholder_is_dropped() -> None:
    slide = _title_and_content_slide("")
    assert not is_body_shape(slide.placeholders[1], "쿠버네티스 핵심 개념", None)


def test_body_placeholder_matching_governing_is_still_deduplicated() -> None:
    text = "Pod는 배포의 최소 단위입니다"
    slide = _title_and_content_slide(text)
    # When the same text was used for the governing line it must not be copied twice.
    assert not is_body_shape(slide.placeholders[1], "쿠버네티스 핵심 개념", text)


def test_slide_text_chars_counts_text_frames_and_table_cells() -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text_frame.text = "가 나 다"
    table = slide.shapes.add_table(2, 2, Inches(1), Inches(3), Inches(4), Inches(2)).table
    table.cell(0, 0).text = "라 마"
    table.cell(1, 1).text = "바"
    assert slide_text_chars(slide) == 3 + 2 + 1


def test_content_loss_warning_thresholds() -> None:
    assert content_loss_warning(3, 30, 0) is None  # too short to judge
    assert content_loss_warning(3, 100, 70) is None  # exactly 70% is fine
    w = content_loss_warning(3, 338, 14)
    assert w and w["code"] == "CONTENT_LOSS" and w["src_index"] == 3
    shown = format_warning(w)
    assert shown["level"] == "warn" and "4" in shown["message"]  # 1-based slide label


def test_slide_skipped_warning_is_displayed() -> None:
    shown = format_warning({"code": "SLIDE_SKIPPED", "src_index": 8})
    assert shown["level"] == "warn" and "9" in shown["message"]


def test_slide_zero_with_table_is_planned_as_content_not_cover() -> None:
    from scripts.build_cmp_academy import build_slide_plan

    prs = Presentation()
    first = prs.slides.add_slide(prs.slide_layouts[6])
    box = first.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    box.text_frame.text = "가상화 방식 비교"
    first.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(8), Inches(2))
    second = prs.slides.add_slide(prs.slide_layouts[6])
    second.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1)).text_frame.text = "본문"
    plan = build_slide_plan(prs)
    assert plan[0] == (0, "content")


def test_title_only_slide_zero_is_still_the_cover() -> None:
    from scripts.build_cmp_academy import build_slide_plan

    prs = Presentation()
    first = prs.slides.add_slide(prs.slide_layouts[6])
    first.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1)).text_frame.text = "Kubernetes 입문"
    assert build_slide_plan(prs)[0] == (0, "cover")


def test_table_loss_warns_even_for_tiny_tables() -> None:
    from scripts.build_cmp_academy import table_loss_warning

    assert table_loss_warning(0, 1, 1) is None
    w = table_loss_warning(0, 1, 0)
    assert w and w["code"] == "CONTENT_LOSS"
    assert "1" in format_warning(w)["message"]


def _png(path, color):
    from PIL import Image

    # A gradient, not a flat colour: flat images are treated as Google stub placeholders.
    img = Image.new("RGB", (800, 450), color)
    img.putdata([(color[0], (color[1] + x // 4) % 256, (color[2] + y // 4) % 256) for y in range(450) for x in range(800)])
    img.save(path)
    return str(path)


def test_unrepeated_full_slide_picture_is_reported_repeated_background_is_not(tmp_path) -> None:
    from scripts.build_cmp_academy import dropped_unrepeated_pictures, picture_hash_counts

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    design_bg = _png(tmp_path / "bg.png", (240, 244, 250))
    screenshot = _png(tmp_path / "shot.png", (10, 120, 200))
    for i in range(4):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(design_bg, 0, 0, prs.slide_width, prs.slide_height)  # on every slide
        slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1)).text_frame.text = f"제목 {i}"
    # slide 3 also carries a unique, slide-sized picture
    prs.slides[3].shapes.add_picture(screenshot, 0, 0, prs.slide_width, prs.slide_height)

    counts = picture_hash_counts(prs)
    sw, sh = int(prs.slide_width), int(prs.slide_height)

    assert dropped_unrepeated_pictures(prs.slides[0], sw, sh, counts) == []  # repeated design bg
    flagged = dropped_unrepeated_pictures(prs.slides[3], sw, sh, counts)
    assert len(flagged) == 1
    assert flagged[0]["width_pct"] == 100 and flagged[0]["height_pct"] == 100


def test_small_content_picture_is_not_reported(tmp_path) -> None:
    from scripts.build_cmp_academy import dropped_unrepeated_pictures, picture_hash_counts

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1)).text_frame.text = "제목"
    slide.shapes.add_picture(_png(tmp_path / "chart.png", (200, 50, 50)), Inches(1), Inches(2), Inches(5), Inches(3))
    counts = picture_hash_counts(prs)
    assert dropped_unrepeated_pictures(prs.slides[0], int(prs.slide_width), int(prs.slide_height), counts) == []


def test_picture_dropped_message_shows_slide_number() -> None:
    shown = format_warning({"code": "PICTURE_DROPPED", "src_index": 12})
    assert shown["level"] == "warn" and "13" in shown["message"]


def test_chart_starting_in_the_header_band_is_body_content() -> None:
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    data = CategoryChartData()
    data.categories = ["A", "B"]
    data.add_series("S", (1, 2))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(0.5), Inches(8), Inches(5), data
    )  # top = 0.5in, inside the header band
    assert is_body_shape(chart, "제목", None)


def test_dropped_header_text_reports_tags_but_not_titles_or_footers() -> None:
    from scripts.build_cmp_academy import dropped_header_text

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1), Inches(0.42), Inches(1.5), Inches(0.3)).text_frame.text = "TECHNICAL"
    slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(9), Inches(0.8)).text_frame.text = "쿠버네티스 핵심 개념"
    slide.shapes.add_textbox(Inches(11.5), Inches(6.9), Inches(0.5), Inches(0.3)).text_frame.text = "7"
    slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(9), Inches(1)).text_frame.text = "본문은 보고 대상이 아닙니다"

    assert dropped_header_text(slide, "쿠버네티스 핵심 개념", None) == ["TECHNICAL"]


def test_header_text_dropped_message_includes_the_text_and_slide() -> None:
    shown = format_warning({"code": "HEADER_TEXT_DROPPED", "src_index": 1, "text": "TECHNICAL"})
    assert shown["level"] == "info"
    assert "TECHNICAL" in shown["message"] and shown["message"].startswith("원본 2장")
