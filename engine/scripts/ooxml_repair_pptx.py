#!/usr/bin/env python3
"""Unpack a .pptx, apply OOXML auto-repair (whitespace), repack in place."""

from __future__ import annotations

import argparse
import sys
import tempfile
import zipfile
from pathlib import Path

def _resolve_office_dir() -> Path:
    """OOXML validators live under ``office/`` (bundled engine or ppt-test skill path)."""
    engine_root = Path(__file__).resolve().parents[1]
    bundled = engine_root / "office"
    if (bundled / "validators" / "__init__.py").is_file():
        return bundled
    legacy = engine_root / ".cursor" / "skills" / "pptx" / "scripts" / "office"
    if (legacy / "validators" / "__init__.py").is_file():
        return legacy
    raise ModuleNotFoundError(
        "validators (OOXML office bundle missing). "
        "Run: cd apps/ppt-academizer && python scripts/sync_engine_from_ppt_test.py"
    )

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def _repair_whitespace_lxml(unpacked: Path) -> int:
    """Add xml:space='preserve' on a:t with edge whitespace (lxml, not minidom)."""
    from lxml import etree

    repairs = 0
    for xml_file in unpacked.rglob("*.xml"):
        try:
            tree = etree.parse(str(xml_file))
        except etree.XMLSyntaxError:
            continue
        root = tree.getroot()
        modified = False
        for el in root.iter(f"{{{A_NS}}}t"):
            text = el.text or ""
            if text and (text[0].isspace() or text[-1].isspace()):
                if el.get(f"{{{XML_NS}}}space") != "preserve":
                    el.set(f"{{{XML_NS}}}space", "preserve")
                    repairs += 1
                    modified = True
        if modified:
            tree.write(
                str(xml_file),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
    return repairs


def repair_pptx_in_place(pptx: Path, *, template: Path | None = None, verbose: bool = False) -> int:
    pptx = pptx.resolve()
    if not pptx.is_file():
        raise FileNotFoundError(pptx)

    office_dir = _resolve_office_dir()
    if str(office_dir) not in sys.path:
        sys.path.insert(0, str(office_dir))
    from validators import PPTXSchemaValidator  # noqa: E402

    original = template.resolve() if template else None

    with tempfile.TemporaryDirectory() as tmp:
        unpacked = Path(tmp)
        with zipfile.ZipFile(pptx, "r") as zin:
            zin.extractall(unpacked)

        repairs = _repair_whitespace_lxml(unpacked)

        validator = PPTXSchemaValidator(unpacked, original, verbose=verbose)
        if not validator.validate():
            raise RuntimeError(f"OOXML validation failed after repair: {pptx}")

        repacked = pptx.with_suffix(".repack.pptx")
        with zipfile.ZipFile(repacked, "w", zipfile.ZIP_DEFLATED) as zout:
            for file_path in sorted(unpacked.rglob("*")):
                if file_path.is_file():
                    arc = file_path.relative_to(unpacked).as_posix()
                    zout.write(file_path, arc)
        repacked.replace(pptx)

    return repairs


def main() -> None:
    parser = argparse.ArgumentParser(description="OOXML auto-repair for .pptx")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--template", type=Path, default=None, help="Baseline .pptx for diff validation")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()
    n = repair_pptx_in_place(args.pptx, template=args.template, verbose=args.verbose)
    print(f"Repaired {n} issue(s); saved {args.pptx}")


if __name__ == "__main__":
    main()
