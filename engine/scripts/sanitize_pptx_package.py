#!/usr/bin/env python3
"""Remove parts that trigger PowerPoint 'Repair' on Mac + renumber slides slide1..N."""

from __future__ import annotations

import platform
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

# GUID fragments inside a:ext/@uri that Mac PowerPoint often rejects (Repair dialog).
OFFICE_EXT_URIS = (
    "FF2B5EF4",
    "DCECCB84",
    "BB962C8B",
    "C183EC19",
    "B7FDDCBA",
    "96DAC541",  # asvg:svgBlip — PNG media with SVG extension triggers Repair
    "28A0092B",  # a14:useLocalDpi
    "9D8B030D",  # table gridCol colId (office/drawing/2014)
    "0D108BD9",  # table grid rowId (office/drawing/2014)
    "D42A27DB",  # table-related ext seen on §7 migrate decks
)
# Local-name tags (any namespace) that should be dropped when nested under a:extLst.
OFFICE_EXT_TAGS = (
    "creationId",
    "custDataLst",
    "svgBlip",
    "colId",
    "rowId",
)
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _normalize_slide_connectors(data: bytes) -> bytes:
    """Mac PowerPoint repair often strips cxnSp with empty ln fill or p:style."""
    from lxml import etree

    root = etree.fromstring(data)
    tag_cxn = f"{{{P_NS}}}cxnSp"
    tag_style = f"{{{P_NS}}}style"
    tag_ln = f"{{{A_NS}}}ln"
    tag_solid = f"{{{A_NS}}}solidFill"
    tag_scheme = f"{{{A_NS}}}schemeClr"

    for cxn in list(root.iter(tag_cxn)):
        for style in cxn.findall(f".//{tag_style}"):
            parent = style.getparent()
            if parent is not None:
                parent.remove(style)
        ln = cxn.find(f".//{tag_ln}")
        if ln is None:
            parent = cxn.getparent()
            if parent is not None:
                parent.remove(cxn)
            continue
        for sf in list(ln.findall(tag_solid)):
            if len(sf) == 0:
                ln.remove(sf)
        if ln.find(tag_solid) is None:
            sf = etree.SubElement(ln, tag_solid)
            sc = etree.SubElement(sf, tag_scheme)
            sc.set("val", "accent1")
        if not ln.get("w"):
            ln.set("w", "12700")

    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def _strip_office_extensions_xml(data: bytes) -> bytes:
    from lxml import etree

    root = etree.fromstring(data)
    remove: list = []
    for el in root.iter():
        tag = el.tag.split("}")[-1] if "}" in str(el.tag) else str(el.tag)
        if tag in OFFICE_EXT_TAGS:
            remove.append(el)
            continue
        if tag == "ext":
            uri = el.get("uri") or ""
            # Keep geometric a:ext under a:xfrm (no uri) — those are width/height.
            if not uri:
                continue
            if any(marker in uri for marker in OFFICE_EXT_URIS):
                remove.append(el)
                continue
            # Defense: office drawing extensions nested under a:ext (SVG / 2014 table ids).
            if any(
                "SVG/main" in str(c.tag)
                or "drawing/2014" in str(c.tag)
                or (c.tag.split("}")[-1] if "}" in str(c.tag) else "") in OFFICE_EXT_TAGS
                for c in el
            ):
                remove.append(el)
                continue
            if "schemas.microsoft.com/office/drawing" in uri:
                remove.append(el)
                continue
    for el in remove:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    # Drop empty extLst left behind after stripping SVG / office extensions.
    for el in list(root.iter()):
        tag = el.tag.split("}")[-1] if "}" in str(el.tag) else str(el.tag)
        if tag == "extLst" and len(el) == 0:
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def _strip_text_outline_xml(data: bytes) -> bytes:
    from lxml import etree

    root = etree.fromstring(data)
    text_props = {
        f"{{{A_NS}}}rPr",
        f"{{{A_NS}}}defRPr",
        f"{{{A_NS}}}endParaRPr",
    }
    outline_tags = {
        f"{{{A_NS}}}ln",
        f"{{{A_NS}}}effectLst",
        f"{{{A_NS}}}effectDag",
    }
    remove: list = []
    for prop in root.iter():
        if prop.tag not in text_props:
            continue
        for child in list(prop.iter()):
            if child is prop or child.tag not in outline_tags:
                continue
            remove.append(child)
    for child in remove:
        parent = child.getparent()
        if parent is not None:
            parent.remove(child)
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


_OVERRIDE_PART_RE = re.compile(
    r'<Override\b[^>]*PartName="(?P<part>[^"]+)"[^>]*/>',
    re.IGNORECASE,
)
_DEFAULT_FNTDATA_RE = re.compile(
    r'<Default\b[^>]*Extension="fntdata"[^>]*/>',
    re.IGNORECASE,
)

# Part paths that trigger PowerPoint "Repair" on Mac when left in the package.
_DROP_PART_PREFIXES = (
    "/ppt/changesInfos/",
    "/ppt/fonts/",
    "/ppt/authors.xml",
    "/ppt/revisionInfo.xml",
)
_DROP_PART_FILES = ("ppt/authors.xml", "ppt/revisionInfo.xml")


def _drop_parts_from_content_types(ct: str) -> str:
    """Remove specific Override/Default entries (safe for single-line [Content_Types].xml)."""

    def _override_repl(match: re.Match[str]) -> str:
        part = match.group("part")
        if any(part.startswith(prefix) or prefix in part for prefix in _DROP_PART_PREFIXES):
            return ""
        return match.group(0)

    ct = _OVERRIDE_PART_RE.sub(_override_repl, ct)
    return _DEFAULT_FNTDATA_RE.sub("", ct)


_RELATIONSHIP_RE = re.compile(r"<Relationship\b[^>]*/>", re.IGNORECASE)


def _should_drop_relationship(tag: str) -> bool:
    low = tag.lower()
    if "changesinfo" in low:
        return True
    if "ppt/fonts/" in low or 'target="fonts/' in low:
        return True
    if "relationships/authors" in low or 'target="authors.xml"' in low:
        return True
    if "revisioninfo" in low:
        return True
    return False


def _drop_rels(rels: str) -> tuple[str, set[str]]:
    """Remove changesInfo / embedded-font relationships; return dropped rIds."""
    dropped: set[str] = set()

    def _keep_rel(match: re.Match[str]) -> str:
        tag = match.group(0)
        if not _should_drop_relationship(tag):
            return tag
        rid = re.search(r'\bId="(rId\d+)"', tag)
        if rid:
            dropped.add(rid.group(1))
        return ""

    return _RELATIONSHIP_RE.sub(_keep_rel, rels), dropped


_EMBEDDED_FONT_LST_RE = re.compile(
    r"<p:embeddedFontLst>.*?</p:embeddedFontLst>",
    re.DOTALL | re.IGNORECASE,
)


def _scrub_presentation_xml(pres: str, dropped_rids: set[str]) -> str:
    """Remove embeddedFontLst and orphan r:id refs after font rels are dropped."""
    if not dropped_rids:
        return pres
    pres = _EMBEDDED_FONT_LST_RE.sub("", pres)
    pres = pres.replace('embedTrueTypeFonts="1"', 'embedTrueTypeFonts="0"')
    for rid in dropped_rids:
        pres = re.sub(rf'\s*r:id="{re.escape(rid)}"', "", pres)
    return pres


def repair_orphan_presentation_refs(path: Path) -> None:
    """Fix presentation.xml r:id pointers removed from .rels (e.g. after font strip)."""
    path = path.resolve()
    tmp = path.with_suffix(".repair.zip")
    rid_in_xml = re.compile(r'\br:id="(rId\d+)"')

    with tempfile.TemporaryDirectory() as tmpdir:
        root = Path(tmpdir)
        with zipfile.ZipFile(path, "r") as zin:
            zin.extractall(root)
        pres_path = root / "ppt/presentation.xml"
        rels_path = root / "ppt/_rels/presentation.xml.rels"
        pres = pres_path.read_text(encoding="utf-8")
        rels = rels_path.read_text(encoding="utf-8")
        valid = set(re.findall(r'\bId="(rId\d+)"', rels))
        orphan = {m.group(1) for m in rid_in_xml.finditer(pres) if m.group(1) not in valid}
        if orphan:
            pres_path.write_text(_scrub_presentation_xml(pres, orphan), encoding="utf-8")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for file_path in sorted(root.rglob("*")):
                if file_path.is_file():
                    zout.write(file_path, file_path.relative_to(root).as_posix())
        tmp.replace(path)


def _convert_svg_to_png(svg_path: Path, png_path: Path) -> bool:
    """Rasterize one SVG for Mac PowerPoint (embedded SVG in layouts triggers Repair)."""
    svg_path = svg_path.resolve()
    png_path = png_path.resolve()
    if platform.system() == "Darwin" and shutil.which("qlmanage"):
        out_dir = png_path.parent
        proc = subprocess.run(
            ["qlmanage", "-t", "-s", "1600", "-o", str(out_dir), str(svg_path)],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            return False
        produced = svg_path.with_name(svg_path.name + ".png")
        if produced.is_file():
            produced.replace(png_path)
            return True
    try:
        import cairosvg  # type: ignore[import-untyped]

        cairosvg.svg2png(url=str(svg_path), write_to=str(png_path))
        return png_path.is_file() and png_path.stat().st_size > 0
    except Exception:
        return False


def rasterize_svg_media(root: Path) -> int:
    """Convert ppt/media/*.svg → .png and rewrite package references."""
    media = root / "ppt" / "media"
    if not media.is_dir():
        return 0

    mapping: dict[str, str] = {}
    for svg_path in sorted(media.glob("*.svg")):
        png_name = svg_path.stem + ".png"
        png_path = media / png_name
        if not _convert_svg_to_png(svg_path, png_path):
            continue
        mapping[svg_path.name] = png_name
        svg_path.unlink(missing_ok=True)

    if not mapping:
        return 0

    ct_path = root / "[Content_Types].xml"
    if ct_path.is_file():
        ct = ct_path.read_text(encoding="utf-8")
        for old, new_name in mapping.items():
            ct = re.sub(
                rf'(PartName="/ppt/media/{re.escape(old)}"\s+ContentType=")image/svg\+xml',
                r'\1image/png',
                ct,
            )
            ct = ct.replace(f"/ppt/media/{old}", f"/ppt/media/{new_name}")
        ct = _fix_png_parts_declared_as_svg(ct)
        ct_path.write_text(ct, encoding="utf-8")

    for xml_path in root.rglob("*"):
        if not xml_path.is_file():
            continue
        if xml_path.name == "[Content_Types].xml":
            continue
        if xml_path.suffix.lower() not in (".xml", ".rels"):
            continue
        text = xml_path.read_text(encoding="utf-8")
        new = text
        for old, new_name in mapping.items():
            new = new.replace(old, new_name)
        if new != text:
            xml_path.write_text(new, encoding="utf-8")

    return len(mapping)


def _fix_png_parts_declared_as_svg(content_types_xml: str) -> str:
    """PNG bytes with image/svg+xml in [Content_Types] breaks Mac PowerPoint display."""
    return re.sub(
        r'(PartName="/ppt/media/[^"]+\.png"\s+ContentType=")image/svg\+xml',
        r"\1image/png",
        content_types_xml,
    )


def finalize_pptx_package(path: Path) -> int:
    """In-place: drop embedded fonts/changesInfo; rename slides to slide1..slideN in order.

    Returns the number of SVG media parts rasterized to PNG.
    """
    path = path.resolve()
    tmp = path.with_suffix(".finalize.zip")

    with tempfile.TemporaryDirectory() as tmpdir:
        root = Path(tmpdir)
        with zipfile.ZipFile(path, "r") as zin:
            zin.extractall(root)

        pres_rels_path = root / "ppt/_rels/presentation.xml.rels"
        pres_path = root / "ppt/presentation.xml"
        pres_rels = pres_rels_path.read_text(encoding="utf-8")
        pres = pres_path.read_text(encoding="utf-8")

        slide_rid_order: list[tuple[str, str]] = []
        for m in re.finditer(r'<p:sldId[^>]*r:id="(rId\d+)"', pres):
            rid = m.group(1)
            tm = re.search(
                rf'<Relationship[^>]*Id="{rid}"[^>]*Target="([^"]+)"',
                pres_rels,
            )
            if tm:
                slide_rid_order.append((rid, tm.group(1)))

        slides_dir = root / "ppt/slides"
        rels_dir = slides_dir / "_rels"
        mapping: dict[str, str] = {}
        for i, (_, target) in enumerate(slide_rid_order, start=1):
            old_name = target.split("/")[-1]
            new_name = f"slide{i}.xml"
            mapping[old_name] = new_name

        for old_name, new_name in mapping.items():
            old_path = slides_dir / old_name
            new_path = slides_dir / new_name
            if old_path.is_file() and old_name != new_name:
                new_path.write_bytes(old_path.read_bytes())
                old_path.unlink(missing_ok=True)
            old_rel = rels_dir / f"{old_name}.rels"
            new_rel = rels_dir / f"{new_name}.rels"
            if old_rel.is_file() and old_name != new_name:
                new_rel.write_bytes(old_rel.read_bytes())
                old_rel.unlink(missing_ok=True)

        notes_rels_dir = root / "ppt" / "notesSlides" / "_rels"
        if notes_rels_dir.is_dir():
            for notes_rels_path in notes_rels_dir.glob("*.rels"):
                notes_rels = notes_rels_path.read_text(encoding="utf-8")
                for old_name, new_name in mapping.items():
                    notes_rels = notes_rels.replace(
                        f"../slides/{old_name}", f"../slides/{new_name}"
                    )
                notes_rels_path.write_text(notes_rels, encoding="utf-8")

        new_pres_rels = pres_rels
        for old_name, new_name in mapping.items():
            new_pres_rels = new_pres_rels.replace(
                f"slides/{old_name}", f"slides/{new_name}"
            )
        pres_rels_path.write_text(new_pres_rels, encoding="utf-8")

        ct_path = root / "[Content_Types].xml"
        ct = ct_path.read_text(encoding="utf-8")
        for old_name, new_name in mapping.items():
            ct = ct.replace(f"slides/{old_name}", f"slides/{new_name}")
        ct = _drop_parts_from_content_types(ct)
        ct = _fix_png_parts_declared_as_svg(ct)
        ct_path.write_text(ct, encoding="utf-8")

        cleaned_rels, dropped_rids = _drop_rels(pres_rels_path.read_text(encoding="utf-8"))
        pres_rels_path.write_text(cleaned_rels, encoding="utf-8")
        pres_path.write_text(
            _scrub_presentation_xml(pres_path.read_text(encoding="utf-8"), dropped_rids),
            encoding="utf-8",
        )

        for folder in ("ppt/changesInfos", "ppt/fonts"):
            p = root / folder
            if p.is_dir():
                for f in p.iterdir():
                    f.unlink()
                p.rmdir()

        for rel_path in _DROP_PART_FILES:
            part = root / rel_path
            if part.is_file():
                part.unlink()

        for xml_file in root.rglob("*.xml"):
            if xml_file.name == "[Content_Types].xml":
                continue
            try:
                raw = xml_file.read_bytes()
                rel = str(xml_file.relative_to(root).as_posix())
                # Slides + layouts/masters: drop p:style on cxnSp (Mac Repair trigger).
                if (
                    (re.match(r"slide\d+\.xml$", xml_file.name) and "ppt/slides/" in rel)
                    or "ppt/slideLayouts/" in rel
                    or "ppt/slideMasters/" in rel
                ):
                    raw = _normalize_slide_connectors(raw)
                raw = _strip_office_extensions_xml(raw)
                raw = _strip_text_outline_xml(raw)
                xml_file.write_bytes(raw)
            except Exception:
                pass

        svg_rasterized = rasterize_svg_media(root)

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for file_path in sorted(root.rglob("*")):
                if file_path.is_file():
                    zout.write(file_path, file_path.relative_to(root).as_posix())
        tmp.replace(path)

    repair_orphan_presentation_refs(path)
    return svg_rasterized


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", type=Path)
    args = parser.parse_args()
    finalize_pptx_package(args.pptx)
    print(f"Finalized: {args.pptx}")


if __name__ == "__main__":
    main()
