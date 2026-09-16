"""Hardened lxml parsing helpers for untrusted OOXML (.pptx) content.

Uploaded .pptx files are user-controlled ZIP archives of XML — the XML
inside them must be treated the same as any other untrusted input.
lxml's *default* parser resolves external/internal entities and expands
DTDs, which can enable XXE (reading local files via a crafted
``<!ENTITY xxe SYSTEM "file:///...">``) or entity-expansion ("billion
laughs") denial of service.

Use ``safe_fromstring``/``safe_parse`` instead of calling
``lxml.etree.fromstring``/``etree.parse`` directly whenever the bytes
being parsed originate from an uploaded file.
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

__all__ = ["safe_fromstring", "safe_parse"]


def _hardened_parser() -> "etree.XMLParser":
    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        dtd_validation=False,
        huge_tree=False,
    )


def safe_fromstring(data: bytes) -> "etree._Element":
    """Drop-in replacement for ``lxml.etree.fromstring`` that disables
    entity resolution and external DTD/network access."""
    return etree.fromstring(data, parser=_hardened_parser())


def safe_parse(source) -> "etree._ElementTree":
    """Drop-in replacement for ``lxml.etree.parse`` that disables entity
    resolution and external DTD/network access.

    Accepts anything ``lxml.etree.parse`` accepts (a path, a ``Path``, an
    open file object, ...); ``Path``/``str`` values are stringified as
    ``lxml.etree.parse`` expects.
    """
    if isinstance(source, (str, Path)):
        source = str(source)
    return etree.parse(source, parser=_hardened_parser())
