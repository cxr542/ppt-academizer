"""Table and text styles of the academy template elements deck, applied to copied body shapes.

Source: ``2.아카데미_템플릿 엘리먼트_v1_250411.pptx`` ("표_기본" and the text-size pages).

Tables ("표_기본"):
* header row: fill accent6, white 12pt text, centered, major (heading) font;
* body first column: fill bg1 at 95% (light gray), 10pt, centered;
* other body cells: no fill, 10pt, left;
* hairline (0.25pt) accent6 borders between cells; no outer left/right border and no border
  above the header; cell margins 72000/36000 EMU, text centered vertically.

Text outside tables: sizes snap to the template's size steps (10/12/14/16/18/20/...), body text
uses the theme body font and large text (20pt and up) the theme heading font, and a box of
several sentences becomes a bullet list with an accent6 bullet.

Only styling changes: text, cell content, positions and sizes are untouched.
"""
from __future__ import annotations

from lxml import etree
from pptx.enum.shapes import MSO_SHAPE_TYPE

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

SIZE_STEPS = (10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 54)
HEADING_MIN_PT = 20
TABLE_HEADER_PT = 12
TABLE_BODY_PT = 10
BORDER_W = "3175"
BULLET_MIN_PARAGRAPHS = 2
BULLET_MIN_CHARS = 10
BULLET_MIN_WIDTH_EMU = 1_800_000
BULLET_MARGIN = "171450"
BULLET_INDENT = "-82550"


def _q(tag: str) -> str:
    return f"{{{A_NS}}}{tag}"


def snap_size(points: float) -> int:
    """The template size step nearest to ``points`` (ties go to the smaller step)."""
    return min(SIZE_STEPS, key=lambda s: (abs(s - points), s))


# ---------------------------------------------------------------- run / paragraph helpers

def _set_font_refs(rpr, *, heading: bool) -> None:
    latin_ref, ea_ref = ("+mj-lt", "+mj-ea") if heading else ("+mn-lt", "+mn-ea")
    for tag, ref in (("latin", latin_ref), ("ea", ea_ref)):
        el = rpr.find(_q(tag))
        if el is None:
            el = etree.SubElement(rpr, _q(tag))
        el.set("typeface", ref)
        for attr in ("panose", "pitchFamily", "charset"):
            el.attrib.pop(attr, None)
    _order_rpr_children(rpr)


_RPR_ORDER = (
    "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst",
    "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
    "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
)


def _order_rpr_children(rpr) -> None:
    """Keep rPr children in schema order after adding latin/ea."""
    rank = {name: i for i, name in enumerate(_RPR_ORDER)}
    children = list(rpr)
    children.sort(key=lambda e: rank.get(etree.QName(e).localname, len(rank)))
    for child in children:
        rpr.remove(child)
    for child in children:
        rpr.append(child)


def _run_properties(run_el):
    rpr = run_el.find(_q("rPr"))
    if rpr is None:
        rpr = etree.Element(_q("rPr"))
        run_el.insert(0, rpr)
    return rpr


def _runs_of(paragraph_el):
    return [r for r in paragraph_el if etree.QName(r).localname in ("r", "fld")]


def _para_size_pt(paragraph_el) -> float | None:
    for r in _runs_of(paragraph_el):
        rpr = r.find(_q("rPr"))
        if rpr is not None and rpr.get("sz"):
            return int(rpr.get("sz")) / 100.0
    return None


def _set_solid_scheme(parent, scheme: str, *, lum_mod: str | None = None) -> etree._Element:
    fill = etree.SubElement(parent, _q("solidFill"))
    clr = etree.SubElement(fill, _q("schemeClr"))
    clr.set("val", scheme)
    if lum_mod:
        etree.SubElement(clr, _q("lumMod")).set("val", lum_mod)
    return fill


# ---------------------------------------------------------------------------- tables

def _border(tag: str, *, visible: bool) -> etree._Element:
    ln = etree.Element(_q(tag), {"w": BORDER_W, "cap": "flat", "cmpd": "sng", "algn": "ctr"})
    if visible:
        _set_solid_scheme(ln, "accent6")
    else:
        etree.SubElement(ln, _q("noFill"))
    etree.SubElement(ln, _q("prstDash")).set("val", "solid")
    return ln


def _style_cell_properties(tc, *, fill: str, visible: dict[str, bool]) -> None:
    """Margins, anchor, four borders and the fill of one cell, in schema order."""
    tcpr = tc.find(_q("tcPr"))
    if tcpr is None:
        tcpr = etree.SubElement(tc, _q("tcPr"))
    keep = [
        child
        for child in tcpr
        if etree.QName(child).localname
        not in ("lnL", "lnR", "lnT", "lnB", "noFill", "solidFill", "gradFill", "blipFill",
                "pattFill", "grpFill")
    ]
    for child in list(tcpr):
        tcpr.remove(child)
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        tcpr.append(_border(tag, visible=visible[tag]))
    if fill == "header":
        _set_solid_scheme(tcpr, "accent6")
    elif fill == "rowhead":
        _set_solid_scheme(tcpr, "bg1", lum_mod="95000")
    else:
        etree.SubElement(tcpr, _q("noFill"))
    for child in keep:
        tcpr.append(child)
    tcpr.set("marL", "72000")
    tcpr.set("marR", "72000")
    tcpr.set("marT", "36000")
    tcpr.set("marB", "36000")
    tcpr.set("anchor", "ctr")


def _style_cell_text(tc, *, role: str) -> None:
    body = tc.find(_q("txBody"))
    if body is None:
        return
    size = TABLE_HEADER_PT if role == "header" else TABLE_BODY_PT
    align = "l" if role == "body" else "ctr"
    for para in body.findall(_q("p")):
        ppr = para.find(_q("pPr"))
        if ppr is None:
            ppr = etree.Element(_q("pPr"))
            para.insert(0, ppr)
        ppr.set("algn", align)
        for r in _runs_of(para):
            rpr = _run_properties(r)
            rpr.set("sz", str(size * 100))
            if role == "header":
                rpr.set("b", "0")
                for old in rpr.findall(_q("solidFill")) + rpr.findall(_q("noFill")):
                    rpr.remove(old)
                fill = etree.SubElement(rpr, _q("solidFill"))
                etree.SubElement(fill, _q("schemeClr")).set("val", "bg1")
            _set_font_refs(rpr, heading=(role == "header"))
        end = para.find(_q("endParaRPr"))
        if end is not None:
            end.set("sz", str(size * 100))


def style_table(table, *, header_rows: int = 1) -> int:
    """Apply the elements table style to ``table``; returns the number of cells styled."""
    n_rows, n_cols = len(table.rows), len(table.columns)
    tbl_pr = table._tbl.find(_q("tblPr"))
    if tbl_pr is not None:
        for attr in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol", "bandCol"):
            tbl_pr.attrib.pop(attr, None)
        for child in list(tbl_pr):
            if etree.QName(child).localname == "tableStyleId":
                tbl_pr.remove(child)
    styled = 0
    for r in range(n_rows):
        for c in range(n_cols):
            cell = table.cell(r, c)
            if cell.is_spanned:
                continue
            tc = cell._tc
            is_header = r < header_rows
            is_rowhead = (not is_header) and c == 0 and n_cols > 1
            role = "header" if is_header else ("rowhead" if is_rowhead else "body")
            visible = {
                "lnL": c != 0,
                "lnR": c != n_cols - 1,
                "lnT": r != 0,
                "lnB": True,
            }
            _style_cell_properties(tc, fill=role if role != "body" else "none", visible=visible)
            _style_cell_text(tc, role=role)
            styled += 1
    return styled


# -------------------------------------------------------------------------------- text

def _snap_paragraph_sizes(paragraph_el) -> int:
    changed = 0
    for r in _runs_of(paragraph_el):
        rpr = r.find(_q("rPr"))
        if rpr is None or not rpr.get("sz"):
            continue
        points = int(rpr.get("sz")) / 100.0
        target = snap_size(points)
        if abs(target - points) > 1e-6:
            rpr.set("sz", str(target * 100))
            changed += 1
    return changed


def _style_runs(paragraph_el) -> None:
    size = _para_size_pt(paragraph_el)
    heading = size is not None and size >= HEADING_MIN_PT
    for r in _runs_of(paragraph_el):
        rpr = _run_properties(r)
        _set_font_refs(rpr, heading=heading)


def _has_bullet(paragraph_el) -> bool:
    ppr = paragraph_el.find(_q("pPr"))
    if ppr is None:
        return False
    return any(
        etree.QName(c).localname in ("buChar", "buAutoNum", "buBlip") for c in ppr
    )


def _has_explicit_no_bullet(paragraph_el) -> bool:
    ppr = paragraph_el.find(_q("pPr"))
    return ppr is not None and ppr.find(_q("buNone")) is not None


def _paragraph_text(paragraph_el) -> str:
    return "".join(t.text or "" for t in paragraph_el.iter(_q("t"))).strip()


def _set_no_bullet(paragraph_el) -> None:
    """Say explicitly that a heading line has no bullet (renderers differ on inheriting one)."""
    ppr = paragraph_el.find(_q("pPr"))
    if ppr is None:
        ppr = etree.Element(_q("pPr"))
        paragraph_el.insert(0, ppr)
    for old in [c for c in ppr if etree.QName(c).localname.startswith("bu")]:
        ppr.remove(old)
    etree.SubElement(ppr, _q("buNone"))
    _order_ppr_children(ppr)


def _bulletize(shape) -> bool:
    """Turn a box of several sentences into an accent6 bullet list."""
    if not shape.has_text_frame or int(shape.width or 0) < BULLET_MIN_WIDTH_EMU:
        return False
    paras = [p for p in shape.text_frame._txBody.findall(_q("p")) if _paragraph_text(p)]
    # A first line set larger than the rest is the card heading: it stays a plain line and only
    # the lines under it become bullets.
    if len(paras) >= 2:
        first, rest = _para_size_pt(paras[0]), [_para_size_pt(p) for p in paras[1:]]
        rest = [r for r in rest if r is not None]
        if first is not None and rest and first > max(rest):
            heading, paras = paras[0], paras[1:]
            if len(paras) >= BULLET_MIN_PARAGRAPHS:
                _set_no_bullet(heading)
    if len(paras) < BULLET_MIN_PARAGRAPHS:
        return False
    if any(len(_paragraph_text(p)) < BULLET_MIN_CHARS or _has_bullet(p) for p in paras):
        return False
    # The source said "no bullet" outright (an explicit buNone): do not guess a list there.
    if any(_has_explicit_no_bullet(p) for p in paras):
        return False
    for p in paras:
        ppr = p.find(_q("pPr"))
        if ppr is None:
            ppr = etree.Element(_q("pPr"))
            p.insert(0, ppr)
        ppr.set("marL", BULLET_MARGIN)
        ppr.set("indent", BULLET_INDENT)
        for old in [c for c in ppr if etree.QName(c).localname.startswith("bu")]:
            ppr.remove(old)
        bu_clr = etree.SubElement(ppr, _q("buClr"))
        etree.SubElement(bu_clr, _q("schemeClr")).set("val", "accent6")
        font = etree.SubElement(ppr, _q("buFont"))
        font.set("typeface", "Arial")
        etree.SubElement(ppr, _q("buChar")).set("char", "•")
        _order_ppr_children(ppr)
    return True


_PPR_ORDER = (
    "lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts",
    "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip", "tabLst", "defRPr",
    "extLst",
)


def _order_ppr_children(ppr) -> None:
    rank = {name: i for i, name in enumerate(_PPR_ORDER)}
    children = sorted(list(ppr), key=lambda e: rank.get(etree.QName(e).localname, len(rank)))
    for child in children:
        ppr.remove(child)
    for child in children:
        ppr.append(child)


def style_text_shape(shape) -> int:
    """Size steps, theme fonts and bullets for one non-table text shape; returns changes."""
    if not shape.has_text_frame or not shape.text_frame.text.strip():
        return 0
    changed = 0
    for para in shape.text_frame._txBody.findall(_q("p")):
        changed += _snap_paragraph_sizes(para)
        _style_runs(para)
    if _bulletize(shape):
        changed += 1
    return changed


# --------------------------------------------------------------------------- the slide

def apply_elements_style(slide) -> dict[str, int]:
    """Style the tables and the text shapes of ``slide`` (placeholders are left to the layout)."""
    counts = {"tables": 0, "text_shapes": 0}

    def walk(shapes) -> None:
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                walk(shape.shapes)
                continue
            if shape.is_placeholder:
                continue
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                style_table(shape.table)
                counts["tables"] += 1
            elif shape.has_text_frame and shape.text_frame.text.strip():
                style_text_shape(shape)
                counts["text_shapes"] += 1

    walk(slide.shapes)
    return counts
