"""Opt-in: turn text baked into pictures into editable text boxes (Tesseract OCR).

The picture is kept. Each recognised line becomes a text box placed over the same spot,
filled with the surrounding background colour so the original glyphs are hidden. Delete the
box to get the original picture back; fix a misread word by editing the box.
Needs the ``tesseract`` binary + kor/eng data; without it nothing happens and the caller
gets a warning instead of an error.
"""
from __future__ import annotations

import io
import re
import shutil
import statistics
from pathlib import Path
from typing import Any

from PIL import Image, ImageOps

MIN_PICTURE_EMU = 914400 * 1.5          # skip icons/logos smaller than 1.5in on a side
MIN_PICTURE_PX = 200
MAX_PIXELS = 12_000_000                 # skip huge rasters (time/memory guard)
MAX_PICTURES = 40
SCALE = 2.0
SHAPE_NAME = "ocr_text"
_HANGUL = re.compile(r"[가-힣]")


def tesseract_available() -> bool:
    if shutil.which("tesseract") is None:
        return False
    try:
        import pytesseract  # noqa: F401
    except ImportError:
        return False
    return True


def _otsu(gray: Image.Image) -> int:
    hist = gray.histogram()
    total = sum(hist)
    sum_all = sum(i * h for i, h in enumerate(hist))
    wb = sb = 0
    best = (0.0, 127)
    for t in range(256):
        wb += hist[t]
        if wb == 0:
            continue
        wf = total - wb
        if wf == 0:
            break
        sb += t * hist[t]
        var = wb * wf * (sb / wb - (sum_all - sb) / wf) ** 2
        if var > best[0]:
            best = (var, t)
    return best[1]


def _words_to_lines(d: dict, scale: float) -> list[dict]:
    groups: dict[tuple, dict] = {}
    for i, raw in enumerate(d["text"]):
        text = raw.strip()
        if not text:
            continue
        try:
            conf = float(d["conf"][i])
        except (TypeError, ValueError):
            continue
        if conf < 0:
            continue
        key = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
        x, y, w, h = d["left"][i], d["top"][i], d["width"][i], d["height"][i]
        g = groups.setdefault(key, {"w": [], "x0": x, "y0": y, "x1": x + w, "y1": y + h, "c": []})
        g["w"].append(text)
        g["x0"], g["y0"] = min(g["x0"], x), min(g["y0"], y)
        g["x1"], g["y1"] = max(g["x1"], x + w), max(g["y1"], y + h)
        g["c"].append(conf)
    return [
        {
            "text": " ".join(g["w"]),
            "box": [g["x0"] / scale, g["y0"] / scale, g["x1"] / scale, g["y1"] / scale],
            "conf": sum(g["c"]) / len(g["c"]),
        }
        for g in groups.values()
    ]


def _run(img: Image.Image, inverted: bool, psm: int) -> list[dict]:
    import pytesseract

    gray = img.convert("L")
    if inverted:
        gray = ImageOps.invert(gray)
    gray = gray.resize((int(gray.width * SCALE), int(gray.height * SCALE)), Image.LANCZOS)
    t = _otsu(gray)
    gray = gray.point(lambda p: 255 if p > t else 0)
    data = pytesseract.image_to_data(
        gray, lang="kor+eng", config=f"--psm {psm}", output_type=pytesseract.Output.DICT
    )
    return _words_to_lines(data, SCALE)


def _good_line(text: str, conf: float) -> bool:
    letters = len(re.findall(r"[A-Za-z0-9가-힣]", text))
    if letters < 2:
        return False
    if _HANGUL.search(text):
        return conf >= 60
    return letters >= 3 and conf >= 82 and bool(re.search(r"[A-Za-z]{3,}|\d{2,}", text))


def _iou(a, b) -> float:
    ix, iy = max(0, min(a[2], b[2]) - max(a[0], b[0])), max(0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    union = (a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - inter
    return inter / union if union else 0.0


def _inside(a, b) -> float:
    ix, iy = max(0, min(a[2], b[2]) - max(a[0], b[0])), max(0, min(a[3], b[3]) - max(a[1], b[1]))
    area = (a[2] - a[0]) * (a[3] - a[1])
    return ix * iy / area if area else 0.0


def ocr_image(img: Image.Image) -> list[dict]:
    """Recognised text lines ``{text, box=[x0,y0,x1,y1] in image px, conf}``, reading order."""
    cands = []
    for inverted in (False, True):          # inverted pass catches light text on dark/colour
        for psm in (6, 11):
            cands.extend(_run(img, inverted, psm))
    cands = [c for c in cands if _good_line(c["text"], c["conf"])]
    cands.sort(key=lambda c: (-(c["conf"] * min(1, len(c["text"]) / 6)), -len(c["text"])))
    kept: list[dict] = []
    for c in cands:
        if any(
            _iou(c["box"], k["box"]) >= 0.25
            or _inside(c["box"], k["box"]) >= 0.5
            or _inside(k["box"], c["box"]) >= 0.5
            for k in kept
        ):
            continue
        kept.append(c)
    return sorted(kept, key=lambda c: (c["box"][1], c["box"][0]))


def _lum(c) -> float:
    return 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]


def _ring_color(im: Image.Image, box, pad: int = 6):
    x0, y0, x1, y1 = (int(v) for v in box)
    px = im.load()
    s = []
    for x in range(max(0, x0 - pad), min(im.width, x1 + pad)):
        s += [px[x, max(0, y0 - pad)], px[x, min(im.height - 1, y1 + pad)]]
    for y in range(max(0, y0 - pad), min(im.height, y1 + pad)):
        s += [px[max(0, x0 - pad), y], px[min(im.width - 1, x1 + pad), y]]
    return tuple(int(statistics.median(c[i] for c in s)) for i in range(3))


def _text_color(im: Image.Image, box, bg):
    x0, y0, x1, y1 = (int(v) for v in box)
    return max(im.crop((x0, y0, x1, y1)).getdata(), key=lambda p: abs(_lum(p) - _lum(bg)))


def _eligible(shape, slide_w: int, slide_h: int) -> bool:
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    if shape.width < MIN_PICTURE_EMU or shape.height < MIN_PICTURE_EMU:
        return False
    # full-bleed background art, not content
    return not (shape.width >= slide_w * 0.98 and shape.height >= slide_h * 0.98)


def _overlay(slide, pic, im: Image.Image, lines: list[dict]) -> int:
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.util import Pt

    sx, sy = pic.width / im.width, pic.height / im.height
    n = 0
    for line in lines:
        x0, y0, x1, y1 = line["box"]
        bg = _ring_color(im, line["box"])
        fg = _text_color(im, line["box"], bg)
        padx, pady = 6, 4
        tb = slide.shapes.add_textbox(
            pic.left + int((x0 - padx) * sx),
            pic.top + int((y0 - pady) * sy),
            int((x1 - x0 + 2 * padx) * sx),
            int((y1 - y0 + 2 * pady) * sy),
        )
        tb.name = SHAPE_NAME
        tf = tb.text_frame
        tf.word_wrap = False
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tb.fill.solid()
        tb.fill.fore_color.rgb = RGBColor(*bg)
        tf.text = line["text"]
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.LEFT
        size = max(8.0, (y1 - y0) * sy / 12700 * 0.80)
        for r in para.runs:
            r.font.size = Pt(size)
            r.font.color.rgb = RGBColor(*fg)
        n += 1
    return n


def apply_image_ocr(pptx_path: Path) -> list[dict[str, Any]]:
    """Add editable text boxes over eligible pictures in *pptx_path* (in place).

    Returns raw warnings (``IMAGE_OCR_DONE`` / ``IMAGE_OCR_UNAVAILABLE``).
    """
    from pptx import Presentation

    if not tesseract_available():
        return [{"code": "IMAGE_OCR_UNAVAILABLE", "message": ""}]
    prs = Presentation(str(pptx_path))
    sw, sh = int(prs.slide_width), int(prs.slide_height)
    pictures = boxes = failed = 0
    for slide in prs.slides:
        for pic in list(slide.shapes):
            if pictures >= MAX_PICTURES:
                break
            if not _eligible(pic, sw, sh):
                continue
            try:
                im = Image.open(io.BytesIO(pic.image.blob))
                if (
                    im.width < MIN_PICTURE_PX
                    or im.height < MIN_PICTURE_PX
                    or im.width * im.height > MAX_PIXELS
                ):
                    continue
                im = im.convert("RGB")
                # honour picture cropping so boxes line up with what is shown
                if any((pic.crop_left, pic.crop_right, pic.crop_top, pic.crop_bottom)):
                    continue
                pictures += 1
                lines = ocr_image(im)
                boxes += _overlay(slide, pic, im, lines)
            except Exception:  # noqa: BLE001 - one bad picture must not fail the deck
                failed += 1
    if boxes:
        prs.save(str(pptx_path))
    return [
        {
            "code": "IMAGE_OCR_DONE",
            "message": "",
            "pictures": pictures,
            "boxes": boxes,
            "failed": failed,
        }
    ]
