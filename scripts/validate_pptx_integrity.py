#!/usr/bin/env python3
# /// script
# requires-python = ">=3.9"
# ///
from __future__ import annotations

import posixpath
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile

RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
OFFICE_REL_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
REQUIRED_PARTS = ("[Content_Types].xml", "ppt/presentation.xml")


@dataclass(frozen=True)
class Relationship:
    source: str
    target: str
    rel_type: str


def _part_dir_for_rels(rels_name: str) -> str:
    if rels_name == "_rels/.rels":
        return ""
    directory, filename = posixpath.split(rels_name)
    if not directory.endswith("_rels"):
        return posixpath.dirname(rels_name)
    owner_dir = posixpath.dirname(directory)
    owner_name = filename[:-5] if filename.endswith(".rels") else filename
    return owner_dir if owner_name else posixpath.dirname(rels_name)


def _resolve_target(rels_name: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = _part_dir_for_rels(rels_name)
    return posixpath.normpath(posixpath.join(base, target))


def _parse_xml_parts(zf: ZipFile) -> list[str]:
    errors: list[str] = []
    for name in sorted(zf.namelist()):
        if not (name.endswith(".xml") or name.endswith(".rels")):
            continue
        try:
            ET.fromstring(zf.read(name))
        except ET.ParseError as err:
            errors.append(f"{name}: XML parse error {err}")
        except KeyError as err:
            errors.append(f"{name}: zip read error {err}")
    return errors


def _relationships(zf: ZipFile) -> tuple[list[Relationship], list[str]]:
    rels: list[Relationship] = []
    errors: list[str] = []
    for name in sorted(n for n in zf.namelist() if n.endswith(".rels")):
        try:
            root = ET.fromstring(zf.read(name))
        except ET.ParseError as err:
            errors.append(f"{name}: relationship XML parse error {err}")
            continue
        for rel in root.findall(f"{RELS_NS}Relationship"):
            target = rel.attrib.get("Target", "")
            mode = rel.attrib.get("TargetMode", "")
            if not target or mode == "External":
                continue
            rels.append(
                Relationship(
                    source=name,
                    target=_resolve_target(name, target),
                    rel_type=rel.attrib.get("Type", ""),
                )
            )
    return rels, errors


def _validate_relationship_targets(names: set[str], rels: list[Relationship]) -> list[str]:
    errors: list[str] = []
    for rel in rels:
        if rel.target not in names:
            errors.append(f"{rel.source}: missing target {rel.target} ({rel.rel_type})")
    return errors


def _validate_notes(names: set[str], rels: list[Relationship]) -> list[str]:
    errors: list[str] = []
    notes_targets: dict[str, list[str]] = {}
    for rel in rels:
        if rel.rel_type == f"{OFFICE_REL_PREFIX}notesSlide":
            notes_targets.setdefault(rel.target, []).append(rel.source)
    for target, sources in sorted(notes_targets.items()):
        if target not in names:
            errors.append(f"notesSlide target missing: {target} from {sources}")
        if len(sources) > 1:
            errors.append(f"notesSlide target shared by multiple slides: {target} from {sources}")
        notes_rels = posixpath.join(
            posixpath.dirname(target),
            "_rels",
            f"{posixpath.basename(target)}.rels",
        )
        if notes_rels not in names:
            errors.append(f"{target}: missing notesSlide rels {notes_rels}")
    if notes_targets and "ppt/notesMasters/notesMaster1.xml" not in names:
        errors.append("notes slides exist but ppt/notesMasters/notesMaster1.xml is missing")
    return errors


def validate_pptx(path: Path) -> list[str]:
    errors: list[str] = []
    try:
        with ZipFile(path) as zf:
            names = set(zf.namelist())
            for required in REQUIRED_PARTS:
                if required not in names:
                    errors.append(f"missing required part: {required}")
            slide_parts = sorted(
                name
                for name in names
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            )
            if not slide_parts:
                errors.append("missing slide parts: ppt/slides/slide*.xml")
            errors.extend(_parse_xml_parts(zf))
            rels, rel_errors = _relationships(zf)
            errors.extend(rel_errors)
            errors.extend(_validate_relationship_targets(names, rels))
            errors.extend(_validate_notes(names, rels))
    except BadZipFile as err:
        errors.append(f"not a valid zip package: {err}")
    return errors


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: validate_pptx_integrity.py <file.pptx> [file.pptx ...]", file=sys.stderr)
        return 2
    failed = False
    for raw_path in sys.argv[1:]:
        path = Path(raw_path)
        errors = validate_pptx(path)
        if errors:
            failed = True
            print(f"FAILED {path}")
            for error in errors:
                print(f"- {error}")
            continue
        print(f"OK {path}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
