"""The layout copyright text stays readable: pale grey on a light layout is darkened, and the
white footer of a dark layout keeps a minimum opacity."""

from pptx import Presentation
from pptx.util import Inches

import scripts.build_cmp_academy as bca


class _Layout:
    def __init__(self, slide):
        self.name, self.shapes, self.part = "L", slide.shapes, slide.part


class _Deck:
    def __init__(self, slide):
        self.slide_layouts = [_Layout(slide)]


def _footer_xml(color: str, alpha: int) -> str:
    return bca._footer_textbox_xml(99, _Box(), "© 2026. OKESTRO Co., Ltd. All Rights Reserved.", color, alpha, 9.5)


class _Box:
    left, top, width, height = Inches(8), Inches(7), Inches(4), Inches(0.2)


def _deck_with_footer(color: str, alpha: int):
    from lxml import etree

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes._spTree.append(etree.fromstring(_footer_xml(color, alpha)))
    return _Deck(slide), slide


def _clr(slide):
    return slide.shapes[-1]._element.xpath(".//a:rPr/a:solidFill/a:srgbClr")[0]


def test_pale_grey_footer_on_a_light_layout_becomes_a_readable_grey():
    deck, slide = _deck_with_footer("C8C8D0", 37)
    assert bca.strengthen_footer_legibility(deck) == 1
    clr = _clr(slide)
    assert clr.get("val") == bca.FOOTER_LIGHT_BG_COLOR
    assert not clr.getchildren()  # fully opaque


def test_white_footer_on_a_dark_layout_keeps_white_but_gets_an_opacity_floor():
    deck, slide = _deck_with_footer("FFFFFF", 37)
    assert bca.strengthen_footer_legibility(deck) == 1
    clr = _clr(slide)
    assert clr.get("val") == "FFFFFF"
    assert int(clr.getchildren()[0].get("val")) == bca.FOOTER_DARK_BG_MIN_ALPHA_PCT * 1000


def test_an_already_visible_footer_is_left_alone():
    deck, slide = _deck_with_footer("FFFFFF", 80)
    assert bca.strengthen_footer_legibility(deck) == 0
