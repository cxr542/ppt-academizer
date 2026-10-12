"""Strip asvg:svgBlip extensions that trigger Mac PowerPoint Repair."""

from __future__ import annotations

from scripts.sanitize_pptx_package import _strip_office_extensions_xml


def test_strip_asvg_svgblip_ext() -> None:
    raw = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:blipFill>
    <a:blip r:embed="rId2">
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip r:embed="rId2"/>
        </a:ext>
      </a:extLst>
    </a:blip>
  </p:blipFill>
</p:pic>
"""
    out = _strip_office_extensions_xml(raw).decode("utf-8")
    assert "svgBlip" not in out
    assert "96DAC541" not in out
    assert 'r:embed="rId2"' in out
    assert "extLst" not in out


def test_keeps_unrelated_markup() -> None:
    raw = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:off xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" x="1" y="2"/>
"""
    out = _strip_office_extensions_xml(raw).decode("utf-8")
    assert 'x="1"' in out
    assert 'y="2"' in out


def test_strip_table_colid_ext() -> None:
    raw = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:gridCol xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" w="100">
  <a:extLst>
    <a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}">
      <a16:colId val="123"/>
    </a:ext>
  </a:extLst>
</a:gridCol>
"""
    out = _strip_office_extensions_xml(raw).decode("utf-8")
    assert "colId" not in out
    assert "9D8B030D" not in out
    assert 'w="100"' in out
    assert "extLst" not in out


def test_keeps_xfrm_geometric_ext() -> None:
    raw = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:off x="1" y="2"/>
  <a:ext cx="10" cy="20"/>
</a:xfrm>
"""
    out = _strip_office_extensions_xml(raw).decode("utf-8")
    assert 'cx="10"' in out
    assert 'cy="20"' in out


_SVG_ONLY_PIC = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:blipFill>
    <a:blip>
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip r:embed="rId3"/>
        </a:ext>
      </a:extLst>
    </a:blip>
  </p:blipFill>
</p:pic>
"""


def test_svg_only_blip_is_kept() -> None:
    """Template logos have no PNG fallback; stripping svgBlip leaves an empty <a:blip/>."""
    out = _strip_office_extensions_xml(_SVG_ONLY_PIC).decode("utf-8")
    assert "svgBlip" in out
    assert 'r:embed="rId3"' in out


def test_svg_media_in_use_and_rasterize_skips_it(tmp_path, monkeypatch) -> None:
    from scripts import sanitize_pptx_package as sp

    layouts = tmp_path / "ppt" / "slideLayouts"
    (layouts / "_rels").mkdir(parents=True)
    media = tmp_path / "ppt" / "media"
    media.mkdir(parents=True)
    (layouts / "slideLayout2.xml").write_bytes(_SVG_ONLY_PIC)
    (layouts / "_rels" / "slideLayout2.xml.rels").write_text(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="../media/image2.svg"/></Relationships>',
        encoding="utf-8",
    )
    (media / "image2.svg").write_text("<svg/>", encoding="utf-8")
    (media / "image9.svg").write_text("<svg/>", encoding="utf-8")  # not referenced by any svgBlip

    assert sp.svg_media_in_use(tmp_path) == {"image2.svg"}

    converted: list[str] = []

    def fake_convert(svg_path, png_path) -> bool:
        converted.append(svg_path.name)
        png_path.write_bytes(b"\x89PNG")
        return True

    monkeypatch.setattr(sp, "_convert_svg_to_png", fake_convert)
    sp.rasterize_svg_media(tmp_path)

    assert converted == ["image9.svg"]
    assert (media / "image2.svg").is_file()
    assert not (media / "image2.png").exists()
