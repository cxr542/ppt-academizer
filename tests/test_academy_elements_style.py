"""Table and text styles of the elements deck: styling changes, content does not."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Emu, Inches, Pt

from scripts.academy_elements_style import (
    apply_elements_style,
    snap_size,
    style_table,
)

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _slide():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs, prs.slides.add_slide(prs.slide_layouts[6])


def _table(slide, rows=3, cols=3):
    frame = slide.shapes.add_table(rows, cols, Inches(1), Inches(1), Inches(8), Inches(3))
    t = frame.table
    for r in range(rows):
        for c in range(cols):
            t.cell(r, c).text_frame.text = f"r{r}c{c}"
    return frame, t


def _scheme(el):
    s = el.find(f"{A}solidFill")
    return None if s is None else s[0].get("val")


def test_snap_size_picks_the_nearest_step() -> None:
    assert snap_size(15) == 14
    assert snap_size(11) == 10
    assert snap_size(13) == 12
    assert snap_size(19) == 18
    assert snap_size(36) == 36
    assert snap_size(8) == 10


def test_table_header_body_and_first_column_follow_the_elements_style() -> None:
    _, slide = _slide()
    frame, t = _table(slide)

    styled = style_table(t)

    assert styled == 9
    head = t.cell(0, 1)._tc.find(f"{A}tcPr")
    assert _scheme(head) == "accent6"
    rpr = t.cell(0, 1)._tc.find(f".//{A}rPr")
    assert rpr.get("sz") == "1200" and rpr.get("b") == "0"
    assert rpr.find(f"{A}solidFill")[0].get("val") == "bg1"

    rowhead = t.cell(1, 0)._tc.find(f"{A}tcPr")
    assert _scheme(rowhead) == "bg1"
    assert rowhead.find(f"{A}solidFill")[0][0].tag == f"{A}lumMod"
    assert t.cell(1, 0)._tc.find(f".//{A}pPr").get("algn") == "ctr"

    body = t.cell(1, 1)._tc.find(f"{A}tcPr")
    assert body.find(f"{A}noFill") is not None
    assert t.cell(1, 1)._tc.find(f".//{A}rPr").get("sz") == "1000"
    assert t.cell(1, 1)._tc.find(f".//{A}pPr").get("algn") == "l"


def test_table_borders_are_hairlines_without_the_outer_sides() -> None:
    _, slide = _slide()
    _, t = _table(slide)
    style_table(t)

    def visible(cell, side):
        return cell._tc.find(f"{A}tcPr").find(f"{A}{side}").find(f"{A}noFill") is None

    assert not visible(t.cell(1, 0), "lnL")  # no left border on the first column
    assert not visible(t.cell(1, 2), "lnR")  # none on the last column
    assert not visible(t.cell(0, 1), "lnT")  # none above the header
    assert visible(t.cell(1, 1), "lnL") and visible(t.cell(1, 1), "lnT")
    assert visible(t.cell(2, 1), "lnB")
    assert t.cell(1, 1)._tc.find(f"{A}tcPr").find(f"{A}lnL").get("w") == "3175"


def test_table_content_and_geometry_are_untouched() -> None:
    _, slide = _slide()
    frame, t = _table(slide)
    geometry = (frame.left, frame.top, frame.width, frame.height)
    widths = [c.width for c in t.columns]

    style_table(t)

    assert [t.cell(r, c).text_frame.text for r in range(3) for c in range(3)] == [
        f"r{r}c{c}" for r in range(3) for c in range(3)
    ]
    assert (frame.left, frame.top, frame.width, frame.height) == geometry
    assert [c.width for c in t.columns] == widths


def test_tcpr_children_stay_in_schema_order() -> None:
    _, slide = _slide()
    _, t = _table(slide)
    style_table(t)
    names = [c.tag.replace(A, "") for c in t.cell(1, 1)._tc.find(f"{A}tcPr")]
    assert names == ["lnL", "lnR", "lnT", "lnB", "noFill"]


def test_text_sizes_snap_and_fonts_follow_the_theme() -> None:
    _, slide = _slide()
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    box.text_frame.text = "본문 글"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(15)
    big = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    big.text_frame.text = "큰 제목"
    big.text_frame.paragraphs[0].runs[0].font.size = Pt(22)

    apply_elements_style(slide)

    r1 = box.text_frame.paragraphs[0].runs[0]._r.find(f"{A}rPr")
    assert r1.get("sz") == "1400"
    assert r1.find(f"{A}latin").get("typeface") == "+mn-lt"
    assert r1.find(f"{A}ea").get("typeface") == "+mn-ea"
    r2 = big.text_frame.paragraphs[0].runs[0]._r.find(f"{A}rPr")
    assert r2.get("sz") == "2000"
    assert r2.find(f"{A}latin").get("typeface") == "+mj-lt"
    assert box.text_frame.text == "본문 글"


def test_a_box_of_sentences_becomes_a_bullet_list_but_a_label_does_not() -> None:
    _, slide = _slide()
    lst = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
    lst.text_frame.text = "첫 번째 항목은 이렇게 길게 적습니다"
    lst.text_frame.add_paragraph().text = "두 번째 항목도 충분히 깁니다"
    label = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(1), Inches(0.5))
    label.text_frame.text = "Q1"
    narrow = slide.shapes.add_textbox(Inches(4), Inches(4), Inches(1), Inches(1))
    narrow.text_frame.text = "아니오 라벨입니다 길이"
    narrow.text_frame.add_paragraph().text = "예 라벨입니다 길이가 깁니다"

    apply_elements_style(slide)

    for p in lst.text_frame.paragraphs:
        assert p._p.find(f"{A}pPr").find(f"{A}buChar") is not None
    assert label._element.find(f".//{A}buChar") is None
    assert narrow._element.find(f".//{A}buChar") is None  # too narrow to be a list


def test_placeholders_are_left_to_the_layout() -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text_frame.text = "제목"
    slide.shapes.title.text_frame.paragraphs[0].runs[0].font.size = Pt(15)

    apply_elements_style(slide)

    rpr = slide.shapes.title.text_frame.paragraphs[0].runs[0]._r.find(f"{A}rPr")
    assert rpr.get("sz") == "1500"


def test_applying_twice_changes_nothing_more() -> None:
    _, slide = _slide()
    _, t = _table(slide)
    box = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(5), Inches(1))
    box.text_frame.text = "본문 글"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(15)
    apply_elements_style(slide)
    from lxml import etree

    first = etree.tostring(slide._element)
    apply_elements_style(slide)
    assert etree.tostring(slide._element) == first


def test_a_card_heading_line_is_not_bulleted() -> None:
    _, slide = _slide()
    card = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
    card.text_frame.text = "VIOLA 제안"
    card.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
    for line in ("표준 파이프라인으로 빠르게 시작", "Java Node.js Python Go 표준 환경"):
        para = card.text_frame.add_paragraph()
        para.text = line
        para.runs[0].font.size = Pt(14)

    apply_elements_style(slide)

    paras = card.text_frame.paragraphs
    assert paras[0]._p.find(f"{A}pPr") is None or paras[0]._p.find(f"{A}pPr").find(f"{A}buChar") is None
    assert all(p._p.find(f"{A}pPr").find(f"{A}buChar") is not None for p in paras[1:])


def test_a_heading_with_a_single_line_gets_no_bullets() -> None:
    _, slide = _slide()
    card = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
    card.text_frame.text = "TROMBONE 제안"
    card.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
    para = card.text_frame.add_paragraph()
    para.text = "티켓과 워크플로우로 변경 이력을 통제"
    para.runs[0].font.size = Pt(14)

    apply_elements_style(slide)

    assert card._element.find(f".//{A}buChar") is None
