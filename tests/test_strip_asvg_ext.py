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
