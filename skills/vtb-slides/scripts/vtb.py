"""VTB slide kit: builds .pptx decks in the VTB corporate style.

Base: assets/vtb_template.pptx (VTB master, logo, page-number plate, 20 cover layouts).
Every slide type is a method of `Deck`; all geometry is in inches on a 13.333 x 7.5 canvas.

    from vtb import Deck
    d = Deck()
    d.cover("Стратегия 2027", "Совет директоров · октябрь 2026", variant=4)
    d.cards("Ключевые направления", [{"icon": "target", "title": "Рост", "text": "..."}])
    d.save("out.pptx")
"""
from __future__ import annotations

import copy
import io
import math
import os
import re
import zipfile

from lxml import etree
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")
TEMPLATE = os.path.join(ASSETS, "vtb_template.pptx")
ICONS = os.path.join(ASSETS, "icons")
PHOTOS = os.path.join(ASSETS, "photos")


def photo_path(p):
    """Path or name of a brand photo from assets/photos (moscow_city, towers, team, office, ...)."""
    if p and not os.path.exists(str(p)):
        cand = os.path.join(PHOTOS, f"{p}.jpg")
        if os.path.exists(cand):
            return cand
    return p

# ---------------------------------------------------------------- brand tokens
NAVY = "002882"      # accent1 — главный цвет бренда
NAVY_D = "001E62"    # accent1 lumMod 75% — тёмный конец градиента
CYAN = "00AAFF"      # accent4 — акцент
CYAN_L = "5FC8FF"    # светлый циан (плашка подзаголовка на обложке)
CYAN_BG = "E5F6FF"   # заливка выделенных строк таблиц
BLUE2 = "33539B"     # accent2
BLUE3 = "667EB4"     # accent3
TEXT = "1C1C1C"      # dk1 — основной текст
GRAY = "878787"      # lt2 — вторичный текст, подписи
GRAY_D = "58595B"    # dk2
LINE = "D9D9D9"      # тонкие разделители
CARD_LINE = "BFBFBF" # контур карточек
BG_SOFT = "F3F5F9"   # подложки
GREEN = "78B497"     # accent5 — позитив в графиках
RED = "E62632"       # accent6 — негатив в графиках
WHITE = "FFFFFF"
CHART_COLORS = [NAVY, CYAN, BLUE3, GRAY, BLUE2, "9FDBFF", GREEN, RED]

SW, SH = 13.333, 7.5
L, R = 0.56, 12.77          # поля
CW = R - L                  # ширина контентной зоны 12.21"
TOP, BOT = 1.35, 6.8        # контентная зона по вертикали (заголовок выше, сноска/номер ниже)
SLANT = 0.366               # фирменный косой срез: смещение = 0.366 * высота (≈20°)

EMU = 914400
HEAD = "+mj-lt"             # VTB Group Demi Bold (тема)
BODY = "+mn-lt"             # VTB Group Book (тема)

LAYOUT_CONTENT = "VTB Content"
LAYOUT_PHOTO = {"people": "VTB Content Photo People", "meeting": "VTB Content Photo Meeting"}
LAYOUT_DARK = "VTB Dark"
LAYOUT_BLANK = "VTB Blank"


def E(v):
    return Emu(int(round(v * EMU)))


def rgb(h):
    return RGBColor.from_string(h.replace("#", "").upper())


# ---------------------------------------------------------------- text metrics
FONTS_DIR = os.path.join(ASSETS, "fonts")


def _font_file(bold):
    """Метрики: VTB Group (assets/fonts или установленный в системе), иначе Inter/Liberation/DejaVu."""
    vtb = "VTB65.otf" if bold else "VTB55.otf"
    for d in (FONTS_DIR, os.path.expanduser("~/.local/share/fonts/vtb"), os.path.expanduser("~/.fonts"),
              "/usr/share/fonts/vtb", "/Library/Fonts", os.path.expanduser("~/Library/Fonts"),
              "C:/Windows/Fonts"):
        p = os.path.join(d, vtb)
        if os.path.exists(p):
            return p
    cands = (["/usr/share/fonts/opentype/inter/Inter-SemiBold.otf",
              "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"] if bold else
             ["/usr/share/fonts/opentype/inter/Inter-Regular.otf",
              "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"])
    for c in cands:
        if os.path.exists(c):
            return c
    return None


_FONTS = {}


def _measure(text, pt, bold=False):
    """Width of a single line in inches (approximate, conservative)."""
    key = (bold, 100)
    if key not in _FONTS:
        f = _font_file(bold)
        _FONTS[key] = ImageFont.truetype(f, 100) if f else None
    font = _FONTS[key]
    if font is None:
        return len(text) * pt * 0.55 / 72
    return font.getlength(text) / 100 * pt / 72 * _SLACK


def has_vtb_fonts():
    f = _font_file(False)
    return bool(f) and os.path.basename(f).startswith("VTB")


_SLACK = 1.02 if has_vtb_fonts() else 1.04


def _strip_md(s):
    return s.replace("**", "")


def _lines_needed(text, w, pt, bold=False):
    n = 0
    for para in _strip_md(str(text)).split("\n"):
        words = para.split(" ")
        line, cnt = "", 1
        for wd in words:
            t = (line + " " + wd).strip()
            if _measure(t, pt, bold) <= w or not line:
                line = t
                if _measure(line, pt, bold) > w:  # single word wider than box
                    cnt += int(_measure(line, pt, bold) // w)
            else:
                cnt += 1
                line = wd
        n += cnt
    return n


def fit_size(text, w, h, max_pt, min_pt, bold=False, spacing=1.2):
    """Largest size in [min_pt, max_pt] at which text fits in w x h inches."""
    pt = max_pt
    while pt >= min_pt:
        if _lines_needed(text, w, pt, bold) * pt * spacing / 72 <= h:
            return pt, True
        pt -= 0.5
    return min_pt, False


# ---------------------------------------------------------------- XML helpers
def _sppr(shape):
    return shape._element.spPr


def _clear_fill(sppr):
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        for el in sppr.findall(qn(tag)):
            sppr.remove(el)


def _insert_fill(sppr, el):
    geom = sppr.find(qn("a:prstGeom"))
    if geom is None:
        geom = sppr.find(qn("a:custGeom"))
    if geom is not None:
        geom.addnext(el)
    else:
        sppr.insert(1, el)


def set_gradient(shape, c1, c2, angle=45, p1=0, p2=100):
    """Linear gradient fill (angle in degrees, clockwise from left-to-right)."""
    sppr = _sppr(shape)
    _clear_fill(sppr)
    xml = (f'<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1">'
           f'<a:gsLst><a:gs pos="{p1 * 1000}"><a:srgbClr val="{c1}"/></a:gs>'
           f'<a:gs pos="{p2 * 1000}"><a:srgbClr val="{c2}"/></a:gs></a:gsLst>'
           f'<a:lin ang="{int(angle * 60000)}" scaled="0"/></a:gradFill>')
    _insert_fill(sppr, etree.fromstring(xml))


def set_alpha(shape, alpha):
    """alpha 0..100 (opacity) for a solid fill."""
    sf = _sppr(shape).find(qn("a:solidFill"))
    if sf is not None and len(sf):
        clr = sf[0]
        for a in clr.findall(qn("a:alpha")):
            clr.remove(a)
        a = etree.SubElement(clr, qn("a:alpha"))
        a.set("val", str(int(alpha * 1000)))


def set_shadow(shape, blur=9, dist=5, alpha=8, angle=45):
    sppr = _sppr(shape)
    for el in sppr.findall(qn("a:effectLst")):
        sppr.remove(el)
    xml = (f'<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           f'<a:outerShdw blurRad="{int(blur * 12700)}" dist="{int(dist * 12700)}" dir="{int(angle * 60000)}" '
           f'algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="{alpha * 1000}"/></a:srgbClr>'
           f'</a:outerShdw></a:effectLst>')
    el = etree.fromstring(xml)
    ln = sppr.find(qn("a:ln"))
    if ln is not None:
        ln.addnext(el)
    else:
        sppr.append(el)


def no_shadow(shape):
    sppr = _sppr(shape)
    for el in sppr.findall(qn("a:effectLst")):
        sppr.remove(el)
    el = etree.SubElement(sppr, qn("a:effectLst"))
    ln = sppr.find(qn("a:ln"))
    if ln is not None:
        ln.addnext(el)


def _bullet(paragraph, char="•", color=NAVY, indent=0.17):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(indent * EMU)))
    pPr.set("indent", str(int(-indent * EMU)))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(buClr, qn("a:srgbClr")).set("val", color)
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", char)


def _no_bullet(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    if pPr.find(qn("a:buNone")) is None:
        etree.SubElement(pPr, qn("a:buNone"))


# ---------------------------------------------------------------- Deck
class Deck:
    """A VTB deck. Methods add slides; `save()` writes the file."""

    def __init__(self, template=TEMPLATE, font_mode="vtb", title_case="upper"):
        """font_mode: 'vtb' keeps VTB Group fonts from the theme (install them on the presenting PC);
        'safe' rewrites theme fonts to Arial so the deck looks the same anywhere.
        title_case: 'upper' (как в шаблоне ВТБ) | 'as_is'."""
        self.prs = Presentation(template)
        self.font_mode = font_mode
        self.title_case = title_case
        self.warnings = []
        self._layouts = {l.name: l for l in self.prs.slide_layouts}
        self._n = 0

    # ------------------------------------------------------------ infrastructure
    def _layout(self, name):
        if name not in self._layouts:
            raise KeyError(f"layout {name!r} not in template; have {list(self._layouts)}")
        return self._layouts[name]

    def new_slide(self, title=None, layout=LAYOUT_CONTENT, note=None, notes=None):
        """Blank content slide with VTB frame (logo, page number). Returns slide."""
        s = self.prs.slides.add_slide(self._layout(layout))
        self._n += 1
        s._vtb_name = f"slide {self._n}"
        for ph in list(s.placeholders):
            t = ph.placeholder_format.type
            idx = ph.placeholder_format.idx
            if idx == 0 or "TITLE" in str(t):
                if title:
                    txt = title.upper() if self.title_case == "upper" else title
                    ph.text_frame.text = txt
                    pt, ok = fit_size(txt, 10.2, 0.95, 28, 20, bold=True)
                    for p in ph.text_frame.paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(pt)
                    if not ok:
                        self._warn(s, f"title too long: {title[:40]}…")
                else:
                    ph._element.getparent().remove(ph._element)
            elif idx == 10 and layout in (LAYOUT_CONTENT, LAYOUT_DARK, *LAYOUT_PHOTO.values()):
                if note:
                    ph.text_frame.text = note
                else:
                    ph._element.getparent().remove(ph._element)
        self._add_slide_number(s, layout)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s

    def _add_slide_number(self, s, layout):
        lay = self._layout(layout)
        for ph in lay.placeholders:
            if "SLIDE_NUMBER" in str(ph.placeholder_format.type):
                el = copy.deepcopy(ph._element)
                s.shapes._spTree.append(el)
                return

    def _warn(self, s, msg):
        self.warnings.append(f"[{getattr(s, '_vtb_name', '?')}] {msg}")

    # ------------------------------------------------------------ primitives
    def text(self, s, x, y, w, h, text, size=16, color=TEXT, bold=False, font=None, align="l",
             anchor="t", min_size=None, bullets=False, bullet_color=NAVY, spacing=1.0,
             para_space=4, hl_color=NAVY, name=None, fit=True):
        """Text box. `text` may be str (\\n = new paragraph) or list (one paragraph each).
        **двойные звёздочки** -> полужирный в цвете hl_color (фирменное выделение ВТБ)."""
        items = text if isinstance(text, (list, tuple)) else str(text).split("\n")
        items = [str(i) for i in items]
        joined = "\n".join(items)
        if fit and joined.strip():
            mn = min_size or max(10, size - 4)
            eff_w = w - (0.17 if bullets else 0)
            sz, ok = fit_size(joined, eff_w, h + 0.05, size, mn, bold,
                              spacing=1.2 * spacing + (para_space / 72 * 0 if True else 0))
            if not ok:
                self._warn(s, f"text overflow ({w:.1f}x{h:.1f}in): {joined[:50]}…")
            size = sz
        tb = s.shapes.add_textbox(E(x), E(y), E(w), E(h))
        if name:
            tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
            if spacing != 1.0:
                p.line_spacing = spacing
            if i > 0 and para_space:
                p.space_before = Pt(para_space)
            if bullets:
                _bullet(p, color=bullet_color)
            parts = re.split(r"(\*\*[^*]+\*\*)", item)
            for part in parts:
                if not part:
                    continue
                hl = part.startswith("**") and part.endswith("**")
                r = p.add_run()
                r.text = part[2:-2] if hl else part
                f = r.font
                f.size = Pt(size)
                strong = hl or bold
                f.color.rgb = rgb(hl_color if hl else color)
                f.name = font or (HEAD if strong else BODY)
                # В теме ВТБ «жирный» = отдельное начертание Demi Bold; флаг b=1 дал бы ложное утолщение.
                f.bold = strong if (self.font_mode == "safe" or (font and font not in (HEAD, BODY))) else False
        return tb

    def shape(self, s, kind, x, y, w, h, fill=None, line=None, line_w=0.75, grad=None, shadow=False,
              name=None, alpha=None, adj=None):
        shp = s.shapes.add_shape(kind, E(x), E(y), E(w), E(h))
        if adj:
            for i, v in enumerate(adj):
                shp.adjustments[i] = v
        self._style(shp, fill, line, line_w, grad, shadow, alpha)
        if name:
            shp.name = name
        return shp

    def _style(self, shp, fill=None, line=None, line_w=0.75, grad=None, shadow=False, alpha=None):
        if grad:
            set_gradient(shp, *grad)
        elif fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = rgb(fill)
            if alpha is not None:
                set_alpha(shp, alpha)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = rgb(line)
            shp.line.width = Pt(line_w)
        else:
            shp.line.fill.background()
        if shadow:
            set_shadow(shp)
        else:
            no_shadow(shp)
        if shp.has_text_frame:
            shp.text_frame.margin_left = shp.text_frame.margin_right = 0
        return shp

    def poly(self, s, pts, fill=None, line=None, grad=None, shadow=False, name=None, line_w=0.75, alpha=None):
        """Closed polygon from (x, y) inch points."""
        fb = s.shapes.build_freeform(E(pts[0][0]), E(pts[0][1]), scale=1.0)
        fb.add_line_segments([(E(px), E(py)) for px, py in pts[1:]], close=True)
        shp = fb.convert_to_shape()
        self._style(shp, fill, line, line_w, grad, shadow, alpha)
        if name:
            shp.name = name
        return shp

    def plate(self, s, x, y, w, h, text="", color="navy", slant="right", size=18, text_color=WHITE,
              pad=0.22, bold=False, align="l", name=None):
        """Фирменная плашка: прямоугольник с косым срезом 20° и градиентом.
        color: 'navy' | 'cyan' | 'gray' | 'light' | hex; slant: right | left | both | none."""
        k = SLANT * h
        if slant == "right":
            pts = [(x, y), (x + w, y), (x + w - k, y + h), (x, y + h)]
        elif slant == "left":
            pts = [(x + k, y), (x + w, y), (x + w, y + h), (x, y + h)]
        elif slant == "both":
            pts = [(x + k, y), (x + w, y), (x + w - k, y + h), (x, y + h)]
        else:
            pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
        grads = {"navy": (NAVY, NAVY_D, 45, 50, 100), "cyan": (CYAN, "0091DA", 45, 0, 100),
                 "gray": ("8C8C8C", "6E6E6E", 45, 0, 100), "light": (CYAN_L, CYAN_L, 0, 0, 100),
                 "mid": (BLUE3, BLUE2, 45, 0, 100)}
        g = grads.get(color)
        shp = self.poly(s, pts, grad=g, fill=None if g else color, name=name)
        if text:
            left_pad = pad + (k if slant in ("left", "both") else 0)
            right_pad = 0.1 + (k if slant in ("right", "both") else 0)
            self.text(s, x + left_pad, y + 0.04, w - left_pad - right_pad, h - 0.08, text, size=size,
                      color=text_color, bold=bold, anchor="m", align=align, min_size=max(10, size - 6),
                      hl_color=WHITE if color == "cyan" else CYAN_L)
        return shp

    def card(self, s, x, y, w, h, fill=WHITE, name=None):
        """Белая карточка с тонким контуром и мягкой тенью (как в шаблоне)."""
        shp = self.shape(s, MSO_SHAPE.RECTANGLE, x, y, w, h, fill=fill, line=CARD_LINE, line_w=0.5,
                         shadow=True, name=name)
        return shp

    def hline(self, s, x, y, w, color=LINE, width=0.75):
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x), E(y), E(x + w), E(y))
        c.line.color.rgb = rgb(color)
        c.line.width = Pt(width)
        return c

    def vline(self, s, x, y, h, color=LINE, width=0.75):
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x), E(y), E(x), E(y + h))
        c.line.color.rgb = rgb(color)
        c.line.width = Pt(width)
        return c

    def line(self, s, x1, y1, x2, y2, color=LINE, width=0.75, dash=False, arrow=False):
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2))
        c.line.color.rgb = rgb(color)
        c.line.width = Pt(width)
        ln = c.line._get_or_add_ln()
        if dash:
            etree.SubElement(ln, qn("a:prstDash")).set("val", "dash")
        if arrow:
            te = etree.SubElement(ln, qn("a:tailEnd"))
            te.set("type", "triangle")
            te.set("w", "med")
            te.set("len", "med")
        return c

    def icon(self, s, name, x, y, size=0.5, color="navy"):
        """Иконка ВТБ. name — одно из 34 имён базового набора (target, team, …) или ID векторной
        библиотеки (o-finance-012 — контурная, g-finance-012 — залитая; каталог: references/icons/).
        color: navy | cyan | white | gray | HEX. Векторные иконки вписываются в квадрат size×size."""
        if not name:
            return None
        if re.match(r"^[og]-[a-z-]+-\d{3}$", str(name)):
            return self._vector_icon(s, name, x, y, size, color)
        if color not in ("navy", "cyan", "white", "gray"):
            color = "navy"
        p = os.path.join(ICONS, f"{name}-{color}.png")
        if not os.path.exists(p):
            self._warn(s, f"icon {name!r} not found; see references/icons.md")
            return None
        return s.shapes.add_picture(p, E(x), E(y), E(size), E(size))

    def _vector_icon(self, s, iid, x, y, size, color):
        lib = icon_library()
        if iid not in lib:
            self._warn(s, f"icon {iid!r} not in library; see references/icons/")
            return None
        it = lib[iid]
        hexc = {"navy": NAVY, "cyan": CYAN, "white": WHITE, "gray": GRAY}.get(color, str(color).lstrip("#").upper())
        el = etree.fromstring(f'<r xmlns:a="{_NS_A}" xmlns:p="{_NS_P}" xmlns:r="{_NS_R}">{it["xml"]}</r>')[0]
        for clr in list(el.iter(qn("a:srgbClr"), qn("a:schemeClr"))):
            if clr.get("val", "").upper() in ("FFFFFF", "BG1", "LT1"):
                continue
            new = etree.Element(qn("a:srgbClr"))
            new.set("val", hexc)
            clr.getparent().replace(clr, new)
        k = size / max(it["w"], it["h"])
        w, h = it["w"] * k, it["h"] * k
        xf = el.find("./p:spPr/a:xfrm", _NSM)
        if xf is None:
            xf = el.find("./p:grpSpPr/a:xfrm", _NSM)
        xf.find("a:off", _NSM).set("x", str(int((x + (size - w) / 2) * EMU)))
        xf.find("a:off", _NSM).set("y", str(int((y + (size - h) / 2) * EMU)))
        xf.find("a:ext", _NSM).set("cx", str(int(w * EMU)))
        xf.find("a:ext", _NSM).set("cy", str(int(h * EMU)))
        tree = s.shapes._spTree
        nid = max([int(c.get("id")) for c in tree.iter(qn("p:cNvPr"))] + [1]) + 1
        for c in el.iter(qn("p:cNvPr")):
            c.set("id", str(nid))
            c.set("name", f"Icon {iid}")
            nid += 1
        tree.append(el)
        return el

    def circle(self, s, x, y, d, fill=None, line=None, line_w=0.75, grad=None, shadow=False, alpha=None):
        return self.shape(s, MSO_SHAPE.OVAL, x, y, d, d, fill=fill, line=line, line_w=line_w, grad=grad,
                          shadow=shadow, alpha=alpha)

    def number_badge(self, s, n, x, y, d=0.5, color="navy", size=None):
        g = {"navy": (NAVY, NAVY_D, 45), "cyan": (CYAN, "0091DA", 45)}.get(color)
        c = self.circle(s, x, y, d, grad=g, fill=None if g else color)
        self.text(s, x, y, d, d, str(n), size=size or d * 30, color=WHITE, align="c", anchor="m", fit=False,
                  font=HEAD)
        return c

    def picture(self, s, path, x, y, w, h, circle=False):
        """Picture cropped to fill w x h (center crop); circle=True -> round mask."""
        im = Image.open(photo_path(path)).convert("RGBA")
        tr = w / h
        ir = im.width / im.height
        if ir > tr:
            nw = int(im.height * tr)
            im = im.crop(((im.width - nw) // 2, 0, (im.width - nw) // 2 + nw, im.height))
        else:
            nh = int(im.width / tr)
            im = im.crop((0, (im.height - nh) // 2, im.width, (im.height - nh) // 2 + nh))
        if circle:
            mask = Image.new("L", im.size, 0)
            ImageDraw.Draw(mask).ellipse((0, 0, im.width - 1, im.height - 1), fill=255)
            im.putalpha(mask)
        buf = io.BytesIO()
        im.save(buf, "PNG")
        buf.seek(0)
        return s.shapes.add_picture(buf, E(x), E(y), E(w), E(h))

    def section_header(self, s, x, y, w, text, color="navy", h=0.48, size=16):
        """Плашка-заголовок блока («Заголовок блока»)."""
        return self.plate(s, x, y, w, h, text, color=color, size=size)

    # ================================================================ SLIDES
    # ---------------------------------------------------------------- титул / разделы
    def cover(self, title, subtitle="", variant=4, notes=None):
        """Титульный слайд. variant 1–20 (см. references/covers.md):
        плашка на фото: 1,2,5,6,9,10,11,12,13,17,18; тёмная панель с диагональю: 3,4,7,8,14,15,16,19,20."""
        lay = self._layout(f"VTB Cover {int(variant):02d}")
        s = self.prs.slides.add_slide(lay)
        self._n += 1
        s._vtb_name = f"slide {self._n} (cover)"
        body_done = False
        for ph in list(s.placeholders):
            t = str(ph.placeholder_format.type)
            if "TITLE" in t:
                ph.text_frame.text = title
                pt, ok = fit_size(title, 4.7, 1.75, 36, 24, bold=True)
                for p in ph.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(pt)
                if not ok:
                    self._warn(s, "cover title too long (≤ 3 строки)")
            elif "BODY" in t and not body_done and subtitle:
                ph.text_frame.text = subtitle
                body_done = True
            else:
                ph._element.getparent().remove(ph._element)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s

    def closing(self, title="Спасибо за внимание!", subtitle="", variant=3, contacts=None, notes=None):
        """Финальный слайд на основе обложки. contacts: list[str] — строки контактов."""
        s = self.cover(title, subtitle, variant=variant, notes=notes)
        if contacts:
            split = int(variant) in (3, 4, 7, 8, 14, 15, 16, 19, 20)
            self.text(s, 0.62, 5.55, 5.0 if split else 6.0, 1.3, contacts, size=14, color=WHITE,
                      para_space=2)
        return s

    def section(self, title, number=None, subtitle="", notes=None):
        """Разделитель раздела: тёмный фирменный фон, номер в циановой плашке."""
        s = self.new_slide(None, layout=LAYOUT_DARK, notes=notes)
        y = 2.55
        if number is not None:
            self.plate(s, 0, y, 3.2, 1.25, "", color="cyan", slant="right")
            self.text(s, L, y, 2.3, 1.25, f"{int(number):02d}" if str(number).isdigit() else str(number),
                      size=60, color=WHITE, anchor="m", bold=True, fit=False)
            y += 1.55
        self.text(s, L, y, 9.5, 1.5, title, size=36, color=WHITE, bold=True, min_size=26)
        if subtitle:
            self.text(s, L, y + 1.55, 8.5, 0.9, subtitle, size=18, color=CYAN_L, min_size=14)
        return s

    def agenda(self, items, title="Содержание", active=None, note=None, notes=None):
        """Повестка: список разделов с номерами. active — индекс (с 0) текущего раздела (выделяется плашкой)."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(items)
        cols = 1 if n <= 5 else 2
        per = math.ceil(n / cols)
        colw = CW if cols == 1 else (CW - 0.6) / 2
        rowh = min(0.95, (BOT - TOP - 0.1) / per)
        for i, it in enumerate(items):
            c, r = divmod(i, per)
            x = L + c * (colw + 0.6)
            y = TOP + 0.1 + r * rowh
            label = it if isinstance(it, str) else it.get("title", "")
            sub = "" if isinstance(it, str) else it.get("text", "")
            on = active is not None and i == active
            if on:
                self.plate(s, x - 0.56 if c == 0 else x, y + 0.04, (colw + 0.56 if c == 0 else colw) * 0.92,
                           rowh - 0.12, "", color="cyan")
            num_c = WHITE if on else CYAN
            txt_c = WHITE if on else NAVY
            self.text(s, x, y + 0.04, 0.85, rowh - 0.12, f"{i + 1:02d}", size=30, color=num_c, anchor="m",
                      fit=False)
            tw = colw * 0.92 - 1.25 if on else colw - 1.0
            if sub:
                self.text(s, x + 0.95, y + 0.1, tw, rowh * 0.42, label, size=20, color=txt_c, anchor="b",
                          min_size=14)
                self.text(s, x + 0.95, y + 0.1 + rowh * 0.45, tw, rowh * 0.4, sub, size=13,
                          color=WHITE if on else GRAY, min_size=11)
            else:
                self.text(s, x + 0.95, y + 0.04, tw, rowh - 0.12, label, size=20, color=txt_c, anchor="m",
                          min_size=14)
            if not on:
                self.hline(s, x, y + rowh - 0.04, colw * 0.92)
        return s

    # ---------------------------------------------------------------- текстовые раскладки
    def bullets(self, title, items, photo=None, note=None, notes=None):
        """Строки «иконка + заголовок + текст» с разделителями (слайд 1–2 шаблона).
        items: [{icon, title, text}], 2–5 шт. photo: 'people' | 'meeting' | путь к картинке | None.
        Первый элемент с highlight=True выделяется синей плашкой."""
        layout = LAYOUT_PHOTO.get(photo, LAYOUT_CONTENT) if isinstance(photo, str) else LAYOUT_CONTENT
        s = self.new_slide(title, layout=layout, note=note, notes=notes)
        custom_photo = photo and photo not in LAYOUT_PHOTO and os.path.exists(str(photo_path(photo)))
        w = 5.2 if photo else CW
        if custom_photo:
            self.image_panel(s, photo)
        n = len(items)
        rowh = min(1.45, (BOT - TOP) / n)
        for i, it in enumerate(items):
            y = TOP + i * rowh
            hl = it.get("highlight")
            if hl:
                self.plate(s, 0, y - 0.05, w + 0.95, rowh - 0.12, "", color="cyan")
            tc = WHITE if hl else NAVY
            if it.get("icon"):
                self.icon(s, it["icon"], L, y + 0.06, 0.42, "white" if hl else "navy")
            self.text(s, L + 0.62, y + 0.05, w - 0.7, 0.4, it.get("title", ""), size=20, color=tc, min_size=14)
            self.text(s, L + 0.62, y + 0.5, w - 0.7, rowh - 0.7, it.get("text", ""), size=15,
                      color=WHITE if hl else TEXT, min_size=11, hl_color=WHITE if hl else NAVY)
            if not hl and i < n - 1:
                self.hline(s, L, y + rowh - 0.08, w, color=CARD_LINE)
        return s

    def image_panel(self, s, path, x0=6.4):
        """Фото справа с фирменной диагональю и тонкой циановой линией (как на обложках)."""
        self.picture(s, path, x0, 0, SW - x0, SH)
        k = SH * SLANT * 0.55
        self.poly(s, [(x0 - 0.01, 0), (x0 + k, 0), (x0, SH), (x0 - 0.01, SH)], fill=WHITE)
        self.line(s, x0 + k, 0, x0, SH, color=CYAN, width=1.5)
        # белая косая подложка под логотип, чтобы он читался на любом фото
        self.poly(s, [(10.05, 0.2), (SW, 0.2), (SW, 1.02), (9.75, 1.02)], fill=WHITE)
        self.logo(s)
        self._page_plate_on_top(s)

    def _page_plate_on_top(self, s):
        """Поднять плашку номера страницы и сам номер над фото."""
        lay = self._layout(LAYOUT_CONTENT)
        for el in lay.shapes._spTree:
            if el.tag == qn("p:sp") and el.find(".//" + qn("p:ph")) is None and el.find(".//" + qn("a:custGeom")) is not None:
                s.shapes._spTree.append(copy.deepcopy(el))
                break
        for el in list(s.shapes._spTree):
            ph = el.find(".//" + qn("p:ph"))
            if ph is not None and ph.get("type") == "sldNum":
                s.shapes._spTree.remove(el)
                s.shapes._spTree.append(el)

    def logo(self, s, white=False):
        """Повторить логотип ВТБ поверх содержимого слайда (если фото/фон перекрыли логотип макета)."""
        lay = self._layout(LAYOUT_DARK if white else LAYOUT_CONTENT)
        for el in lay.shapes._spTree:
            if el.tag == qn("p:grpSp"):
                s.shapes._spTree.append(copy.deepcopy(el))
                return

    def text_image(self, title, text, image, note=None, notes=None, kicker=None):
        """Текст слева, фото справа с диагональю. text: str | list (абзацы); **выделение**."""
        s = self.new_slide(title, note=note, notes=notes)
        self.image_panel(s, image, x0=7.0)
        y = TOP + 0.1
        if kicker:
            self.plate(s, 0, y, 5.4, 0.6, kicker, color="navy", size=18, pad=L)
            y += 0.9
        self.text(s, L, y, 5.6, BOT - y, text, size=16, para_space=8, min_size=11)
        return s

    def statement(self, title, text, kicker=None, source=None, note=None, notes=None):
        """Ключевой тезис / цитата крупно: циановая плашка-маркер слева, текст 28pt."""
        s = self.new_slide(title, note=note, notes=notes)
        y = 2.2
        self.plate(s, 0, y, 1.45, 1.0, "", color="cyan")
        if kicker:
            self.text(s, 2.0, y - 0.6, 10, 0.45, kicker, size=16, color=CYAN, bold=True)
        self.text(s, 2.0, y, 10.2, 3.0, text, size=30, color=NAVY, min_size=18, spacing=1.05)
        if source:
            self.text(s, 2.0, 5.6, 10, 0.5, source, size=14, color=GRAY)
        return s

    def quote(self, text, author, role="", title=None, photo=None, note=None, notes=None):
        """Цитата: крупные кавычки в плашке, текст, автор (с фото по желанию)."""
        s = self.new_slide(title, note=note, notes=notes)
        self.plate(s, 0, 2.2, 1.6, 1.1, "«", color="navy", size=66, align="r", pad=0.3)
        self.text(s, 2.1, 2.2, 10.2, 2.7, text, size=28, color=NAVY, min_size=16, spacing=1.05)
        x = 2.1
        if photo and os.path.exists(str(photo_path(photo))):
            self.picture(s, photo, 2.1, 5.15, 0.9, 0.9, circle=True)
            x = 3.2
        self.text(s, x, 5.2, 8, 0.4, author, size=18, color=TEXT, bold=True)
        if role:
            self.text(s, x, 5.62, 8, 0.4, role, size=14, color=GRAY)
        return s

    # ---------------------------------------------------------------- карточки и колонки
    def cards(self, title, items, cols=None, highlight=None, note=None, notes=None):
        """Сетка карточек «иконка + заголовок + текст» (слайды 3–5). 2–8 элементов.
        highlight: индекс карточки, выделяемой циановым параллелограммом."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(items)
        cols = cols or (n if n <= 4 else (3 if n in (5, 6, 9) else 4))
        rows = math.ceil(n / cols)
        gap = 0.45
        cw = (CW - gap * (cols - 1)) / cols
        ch = (BOT - TOP - 0.15 - (rows - 1) * 0.35) / rows
        for i, it in enumerate(items):
            r, c = divmod(i, cols)
            x = L + c * (cw + gap)
            y = TOP + 0.15 + r * (ch + 0.35)
            on = highlight is not None and i == highlight
            if on:
                k = SLANT * 0.6
                self.poly(s, [(x - 0.12 + k, y - 0.1), (x + cw + 0.12, y - 0.1), (x + cw + 0.12 - k, y + ch),
                              (x - 0.12, y + ch)], grad=(CYAN, "0091DA", 45))
            self.icon(s, it.get("icon"), x + (0.1 if on else 0), y, 0.45, "white" if on else "navy")
            ty = y + (0.62 if it.get("icon") else 0)
            nl = _lines_needed(it.get("title", ""), cw - 0.45, 19)
            th = min(0.8, ch * 0.32, nl * 19 * 1.2 / 72 + 0.08)
            self.text(s, x + (0.1 if on else 0), ty, cw - (0.4 if on else 0.05), th, it.get("title", ""),
                      size=19, color=WHITE if on else NAVY, min_size=13)
            self.text(s, x + (0.1 if on else 0), ty + th + 0.08, cw - (0.45 if on else 0.05),
                      ch - (ty - y) - th - 0.25, it.get("text", ""), size=15, color=WHITE if on else TEXT,
                      min_size=10.5, hl_color=WHITE if on else NAVY)
            if not on:
                self.hline(s, x, y + ch - 0.04, cw, color=CARD_LINE)
        return s

    def columns(self, title, cols, color_scheme="navy", note=None, notes=None):
        """Колонки с плашками-заголовками и буллетами в карточках (слайды 21–23).
        cols: [{title, icon?, items: [str], highlight?}] 2–4 шт. color_scheme: navy | mixed (последняя циан)."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(cols)
        gap = 0.18
        cw = (CW - gap * (n - 1)) / n
        has_icon = any(c.get("icon") for c in cols)
        y0 = TOP + (0.6 if has_icon else 0)
        for i, col in enumerate(cols):
            x = L + i * (cw + gap)
            color = "cyan" if (col.get("highlight") or (color_scheme == "mixed" and i == n - 1)) else (
                "gray" if col.get("muted") else "navy")
            if col.get("icon"):
                self.icon(s, col["icon"], x + 0.05, TOP, 0.45, "navy" if color != "cyan" else "cyan")
            self.card(s, x, y0, cw, BOT - y0)
            self.plate(s, x, y0, cw, 0.62, col.get("title", ""), color=color, size=17)
            self.text(s, x + 0.2, y0 + 0.85, cw - 0.35, BOT - y0 - 1.0, col.get("items", []), size=15,
                      bullets=True, min_size=10.5, para_space=6)
        return s

    def numbered(self, title, items, panel=None, header=None, accent="navy", note=None, notes=None):
        """Нумерованный список 01–07 с разделителями (слайды 13–15).
        panel: {icon?, text} — циановая панель слева. header — плашка над списком."""
        s = self.new_slide(title, note=note, notes=notes)
        x, w = L, CW
        if panel:
            pw = 3.6
            self.shape(s, MSO_SHAPE.RECTANGLE, L, TOP, pw, BOT - TOP, grad=(CYAN, "0091DA", 45))
            if panel.get("icon"):
                self.icon(s, panel["icon"], L + 0.35, TOP + 0.4, 0.55, "white")
            self.text(s, L + 0.35, TOP + 1.15, pw - 0.7, BOT - TOP - 1.4, panel.get("text", ""), size=18,
                      color=WHITE, min_size=12, hl_color=WHITE)
            x, w = L + pw + 0.4, CW - pw - 0.4
        y = TOP
        if header:
            self.plate(s, x, y, w, 0.55, header, color=accent, size=17)
            y += 0.7
        n = len(items)
        rowh = (BOT - y) / n
        for i, it in enumerate(items):
            ry = y + i * rowh
            self.text(s, x + 0.05, ry, 0.85, rowh - 0.08, f"{i + 1:02d}", size=min(34, rowh * 40),
                      color=CYAN if accent == "cyan" else NAVY, anchor="m", fit=False)
            self.text(s, x + 1.0, ry + 0.06, w - 1.1, rowh - 0.18, it, size=16, anchor="m", min_size=11)
            if i < n - 1:
                self.hline(s, x, ry + rowh - 0.02, w)
        return s

    def two_lists(self, title, left, right, note=None, notes=None):
        """Два нумерованных списка в карточках с плашками (слайд 13). left/right: {title, items}."""
        s = self.new_slide(title, note=note, notes=notes)
        w = (CW - 0.3) / 2
        for j, blk in enumerate((left, right)):
            x = L + j * (w + 0.3)
            self.card(s, x, TOP, w, BOT - TOP)
            col = "navy" if j == 0 else "cyan"
            self.plate(s, x, TOP, w, 0.58, blk.get("title", ""), color=col, size=17)
            items = blk.get("items", [])
            rowh = (BOT - TOP - 0.75) / max(1, len(items))
            for i, it in enumerate(items):
                ry = TOP + 0.7 + i * rowh
                self.text(s, x + 0.2, ry, 0.75, rowh - 0.06, f"{i + 1:02d}", size=min(32, rowh * 38),
                          color=NAVY if j == 0 else CYAN, anchor="m", fit=False)
                self.text(s, x + 1.0, ry, w - 1.2, rowh - 0.06, it, size=15, anchor="m", min_size=10.5)
                if i < len(items) - 1:
                    self.hline(s, x, ry + rowh, w)
        return s

    # ---------------------------------------------------------------- цифры
    def kpi_cards(self, title, items, note=None, notes=None):
        """Карточки показателей с диагональной синей панелью (слайд 30). 2–4 шт.
        items: [{name, icon?, value?, unit?, text?, bullets?: [str], label?, highlight?}]"""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(items)
        cols = 1 if n == 1 else 2
        rows = math.ceil(n / cols)
        gx, gy = 0.3, 0.45
        cw = (CW - gx * (cols - 1)) / cols
        ch = (BOT - TOP - (rows - 1) * gy - 0.35) / rows
        for i, it in enumerate(items):
            r, c = divmod(i, cols)
            x = L + c * (cw + gx)
            y = TOP + 0.35 + r * (ch + gy)
            if it.get("label"):
                self.text(s, x + cw * 0.47, y - 0.36, cw * 0.5, 0.3, it["label"], size=13, color=GRAY,
                          anchor="b")
                self.hline(s, x + cw * 0.47, y - 0.04, cw * 0.53, color=CARD_LINE, width=1)
            self.card(s, x, y, cw, ch)
            pw = cw * 0.47
            k = ch * SLANT
            hl = it.get("highlight")
            self.poly(s, [(x, y), (x + pw, y), (x + pw - k, y + ch), (x, y + ch)],
                      grad=(CYAN, "0091DA", 45) if hl else (NAVY, NAVY_D, 45))
            self.icon(s, it.get("icon"), x + 0.25, y + ch - 1.2, 0.45, "white")
            self.text(s, x + 0.25, y + ch - 0.62, pw - k - 0.2, 0.5, it.get("name", ""), size=18,
                      color=WHITE, anchor="m", min_size=12)
            tx, tw = x + pw + 0.1, cw - pw - 0.3
            if it.get("value") is not None:
                val_c = CYAN if hl else NAVY
                tb = self.text(s, tx, y + 0.25, tw, 0.85, str(it["value"]), size=40, color=val_c, anchor="b",
                               min_size=24, fit=True)
                if it.get("unit"):
                    p = tb.text_frame.paragraphs[0]
                    r_ = p.add_run()
                    r_.text = " " + it["unit"]
                    r_.font.size = Pt(24)
                    r_.font.color.rgb = rgb(val_c)
                    r_.font.name = BODY
                self.text(s, tx, y + 1.18, tw, ch - 1.35, it.get("text", ""), size=15, min_size=10.5)
            else:
                self.text(s, tx, y + 0.25, tw, ch - 0.45, it.get("bullets", it.get("text", "")), size=14,
                          bullets=bool(it.get("bullets")), anchor="m", min_size=10, para_space=6)
        return s

    def stats(self, title, items, text=None, note=None, notes=None):
        """Крупные цифры в ряд (2–4): значение 54pt, подпись, иконка. text — вывод под рядом."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(items)
        gap = 0.4
        cw = (CW - gap * (n - 1)) / n
        y = TOP + 0.4
        h = 3.3 if text else 4.4
        for i, it in enumerate(items):
            x = L + i * (cw + gap)
            hl = it.get("highlight", i == 0)
            self.card(s, x, y, cw, h)
            self.icon(s, it.get("icon"), x + 0.3, y + 0.3, 0.5, "cyan" if hl else "navy")
            tb = self.text(s, x + 0.3, y + 0.95, cw - 0.5, 1.0, str(it.get("value", "")), size=54,
                           color=CYAN if hl else NAVY, anchor="b", min_size=28)
            if it.get("unit"):
                r_ = tb.text_frame.paragraphs[0].add_run()
                r_.text = " " + it["unit"]
                r_.font.size = Pt(26)
                r_.font.color.rgb = rgb(CYAN if hl else NAVY)
                r_.font.name = BODY
            self.hline(s, x + 0.3, y + 2.1, cw - 0.6, color=CARD_LINE)
            self.text(s, x + 0.3, y + 2.25, cw - 0.55, h - 2.4, it.get("label", ""), size=16, min_size=11)
        if text:
            self.plate(s, 0, y + h + 0.45, CW * 0.95, 0.95 if len(text) > 90 else 0.75, text, color="navy",
                       size=17, pad=L)
        return s

    def progress(self, title, items, note=None, notes=None):
        """Горизонтальные шкалы прогресса (ADKAR/навыки/выполнение плана).
        items: [{label, value: 0–100, text?, icon?}]"""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(items)
        rowh = min(1.05, (BOT - TOP) / n)
        for i, it in enumerate(items):
            y = TOP + 0.1 + i * rowh
            x = L
            if it.get("icon"):
                self.circle(s, x, y + 0.05, 0.6, grad=(NAVY, NAVY_D, 45))
                self.icon(s, it["icon"], x + 0.12, y + 0.17, 0.36, "white")
                x += 0.8
            self.text(s, x, y, 3.3, 0.35, it.get("label", ""), size=17, color=NAVY, bold=True, min_size=12)
            if it.get("text"):
                self.text(s, x, y + 0.36, 3.3, rowh - 0.45, it["text"], size=12, color=GRAY, min_size=10)
            bx, bw = L + 4.4, CW - 4.4 - 1.0
            by = y + 0.22
            self.shape(s, MSO_SHAPE.RECTANGLE, bx, by, bw, 0.16, fill="E6EAF2")
            v = max(0, min(100, float(it.get("value", 0))))
            if v > 0:
                self.shape(s, MSO_SHAPE.RECTANGLE, bx, by, bw * v / 100, 0.16, grad=(CYAN, NAVY, 0))
            self.text(s, bx + bw + 0.15, y + 0.02, 0.85, 0.5, f"{v:g}%", size=20, color=NAVY, fit=False,
                      anchor="m")
        return s

    # ---------------------------------------------------------------- процессы и время
    def process(self, title, steps, label="Шаг", note=None, notes=None):
        """Этапы с шевронными плашками «ШАГ 1 → 2 → 3» и карточками (слайд 63). 2–5 шагов.
        steps: [{title, text?, items?: [str]}]"""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(steps)
        gap = 0.08
        cw = (CW - gap * (n - 1)) / n
        for i, st in enumerate(steps):
            x = L + i * (cw + gap)
            h = 0.6
            tip = 0.28
            pts = [(x, TOP), (x + cw - tip, TOP), (x + cw + tip * 0.0, TOP + h / 2), (x + cw - tip, TOP + h),
                   (x, TOP + h)]
            if i > 0:
                pts.insert(5, (x + tip, TOP + h / 2))
            c1 = CYAN if i == 0 else (BLUE3 if i < n - 1 else NAVY)
            frac = i / max(1, n - 1)
            c = _mix(CYAN, NAVY, frac)
            self.poly(s, pts, grad=(c, _mix(c, NAVY_D, 0.35), 0))
            self.text(s, x + (0.35 if i else 0.2), TOP, cw - 0.6, h, f"{label.upper()} {i + 1}", size=17,
                      color=WHITE, anchor="m", bold=True, fit=False)
            cy = TOP + h + 0.18
            self.card(s, x, cy, cw - 0.1, BOT - cy)
            self.text(s, x + 0.2, cy + 0.2, cw - 0.45, 0.85, st.get("title", ""), size=17, color=CYAN,
                      min_size=12)
            body = st.get("items") or st.get("text", "")
            self.text(s, x + 0.2, cy + 1.1, cw - 0.45, BOT - cy - 1.25, body, size=14,
                      bullets=bool(st.get("items")), min_size=10, para_space=5)
        return s

    def chevrons(self, title, steps, note=None, notes=None):
        """Поток шевронов с иконками и описанием под ними (категории Arrow/Flow/Process). 3–6 шагов.
        steps: [{title, text, icon?}]"""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(steps)
        cw = CW / n
        y, h = TOP + 0.9, 1.25
        tip = 0.4
        for i, st in enumerate(steps):
            x = L + i * cw
            pts = [(x, y), (x + cw - 0.05, y), (x + cw + tip - 0.05, y + h / 2), (x + cw - 0.05, y + h),
                   (x, y + h)]
            if i > 0:
                pts.insert(5, (x + tip, y + h / 2))
            c = _mix(CYAN, NAVY, i / max(1, n - 1))
            self.poly(s, pts, grad=(c, _mix(c, NAVY_D, 0.3), 0))
            ix = x + (tip if i else 0) + (cw - tip) / 2 - 0.25
            if st.get("icon"):
                self.icon(s, st["icon"], ix, y + 0.33, 0.5, "white")
            else:
                self.text(s, x + (tip if i else 0.1), y, cw - tip, h, f"{i + 1:02d}", size=30, color=WHITE,
                          align="c", anchor="m", fit=False)
            self.text(s, x + 0.1, y + h + 0.4, cw - 0.3, 0.45, st.get("title", ""), size=18, color=NAVY,
                      min_size=12, bold=True)
            self.hline(s, x + 0.1, y + h + 0.95, 0.5, color=CYAN, width=1.5)
            self.text(s, x + 0.1, y + h + 1.1, cw - 0.3, BOT - (y + h + 1.1), st.get("text", ""), size=15,
                      min_size=10)
        return s

    def flow_rows(self, title, rows, note=None, notes=None):
        """Строки «иконка в квадрате → стрелка → пояснение» (слайды 33–34). 2–5 строк.
        rows: [{icon?, label, text, highlight?}]"""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(rows)
        gap = 0.18
        rh = (BOT - TOP - gap * (n - 1)) / n
        for i, rw in enumerate(rows):
            y = TOP + i * (rh + gap)
            hl = rw.get("highlight")
            self.card(s, L, y, CW, rh)
            sq = min(rh, 1.0)
            self.shape(s, MSO_SHAPE.RECTANGLE, L, y, sq, rh, grad=(CYAN, "0091DA", 45) if hl else (NAVY, NAVY_D, 45))
            self.icon(s, rw.get("icon"), L + sq / 2 - 0.22, y + rh / 2 - 0.22, 0.44, "white")
            self.text(s, L + sq + 0.3, y + 0.08, 3.4, rh - 0.16, rw.get("label", ""), size=16,
                      color=CYAN if hl else NAVY, anchor="m", min_size=11)
            ax = L + sq + 3.9
            self.poly(s, [(ax, y + 0.12), (ax + 0.12, y + 0.12), (ax + 0.32, y + rh / 2), (ax + 0.12, y + rh - 0.12),
                          (ax, y + rh - 0.12), (ax + 0.2, y + rh / 2)], grad=(CYAN, NAVY, 90))
            self.text(s, ax + 0.6, y + 0.08, R - ax - 0.8, rh - 0.16, rw.get("text", ""), size=14, anchor="m",
                      min_size=10)
        return s

    def timeline(self, title, events, note=None, notes=None):
        """Горизонтальная шкала событий, подписи чередуются сверху/снизу (слайды 64–66). 3–8 событий.
        events: [{date, title, text?, highlight?}]"""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(events)
        ay = 4.05
        self.line(s, L, ay, R, ay, color=CARD_LINE, width=1.25, dash=True)
        step = CW / n
        for i, ev in enumerate(events):
            cx = L + step * (i + 0.5)
            hl = ev.get("highlight")
            col = CYAN if hl else NAVY
            self.circle(s, cx - 0.11, ay - 0.11, 0.22, fill=col)
            self.circle(s, cx - 0.2, ay - 0.2, 0.4, line=col, line_w=0.75)
            up = i % 2 == 0
            tx = cx - 0.04
            tw = min(2.6, step * 1.75, R + 0.4 - tx - 0.15)
            if up:
                self.vline(s, cx, TOP + 0.15, ay - 0.25 - TOP - 0.15, color=col if hl else CARD_LINE, width=1)
                self.text(s, tx + 0.15, TOP + 0.1, tw, 0.4, ev.get("date", ""), size=20, color=col, bold=True,
                          fit=False)
                self.text(s, tx + 0.15, TOP + 0.55, tw, 0.6, ev.get("title", ""), size=16, color=NAVY,
                          min_size=11)
                self.text(s, tx + 0.15, TOP + 1.15, tw, ay - TOP - 1.45, ev.get("text", ""), size=13,
                          color=GRAY_D, min_size=10)
            else:
                self.vline(s, cx, ay + 0.25, BOT - ay - 0.3, color=col if hl else CARD_LINE, width=1)
                self.text(s, tx + 0.15, ay + 0.35, tw, 0.4, ev.get("date", ""), size=20, color=col, bold=True,
                          fit=False)
                self.text(s, tx + 0.15, ay + 0.8, tw, 0.6, ev.get("title", ""), size=16, color=NAVY,
                          min_size=11)
                self.text(s, tx + 0.15, ay + 1.4, tw, BOT - ay - 1.45, ev.get("text", ""), size=13,
                          color=GRAY_D, min_size=10)
        return s

    def roadmap(self, title, periods, rows, header="Направление", note=None, notes=None):
        """Диаграмма Ганта / дорожная карта (слайды 67–68).
        periods: ['Q1', 'Q2', ...]; rows: [{name, start, end, highlight?, label?}] (start/end — индексы периодов, end включительно)."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(rows)
        lw = 3.3
        gx, gw = L + lw, CW - lw
        pw = gw / len(periods)
        hh = 0.42
        rh = min(0.75, (BOT - TOP - hh) / n)
        self.shape(s, MSO_SHAPE.RECTANGLE, gx, TOP, gw, hh, grad=(NAVY, NAVY_D, 0))
        self.text(s, L + 0.1, TOP, lw - 0.2, hh, header, size=13, color=GRAY, anchor="m", fit=False)
        for j, p in enumerate(periods):
            self.text(s, gx + j * pw, TOP, pw, hh, str(p), size=12, color=WHITE, align="c", anchor="m",
                      min_size=9)
            if j:
                self.vline(s, gx + j * pw, TOP + hh, rh * n, color="E6EAF2")
        for i, rw in enumerate(rows):
            y = TOP + hh + i * rh
            if i % 2 == 1:
                self.shape(s, MSO_SHAPE.RECTANGLE, L, y, CW, rh, fill="F5F8FC")
            self.text(s, L + 0.1, y, 0.45, rh, f"{i + 1:02d}", size=12, color=CYAN, anchor="m", fit=False)
            self.text(s, L + 0.55, y, lw - 0.65, rh, rw.get("name", ""), size=13, anchor="m", min_size=9)
            st, en = rw.get("start", 0), rw.get("end", 0)
            bx = gx + st * pw + 0.05
            bw = (en - st + 1) * pw - 0.1
            g = (NAVY, NAVY_D, 0) if rw.get("highlight") else (CYAN, "0091DA", 0)
            self.shape(s, MSO_SHAPE.RECTANGLE, bx, y + rh * 0.25, bw, rh * 0.5, grad=g)
            if rw.get("label"):
                self.text(s, bx + 0.08, y + rh * 0.25, bw - 0.1, rh * 0.5, rw["label"], size=10, color=WHITE,
                          anchor="m", min_size=8)
            self.hline(s, L, y + rh, CW, color="E6EAF2")
        return s

    # ---------------------------------------------------------------- диаграммы-схемы
    def funnel(self, title, stages, note=None, notes=None):
        """Воронка (слайд 62): сужающиеся сегменты от синего к циану, подписи чередуются слева/справа.
        stages: [{label, text?, value?}] 3–6 шт."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(stages)
        cx = SW / 2
        top_w, bot_w = 4.6, 1.1
        y0, total = TOP + 0.1, BOT - TOP - 0.2
        sh = total / n
        for i, st in enumerate(stages):
            y = y0 + i * sh
            w1 = top_w - (top_w - bot_w) * i / n
            w2 = top_w - (top_w - bot_w) * (i + 1) / n
            c = _mix(NAVY_D, CYAN, i / max(1, n - 1))
            self.poly(s, [(cx - w1 / 2, y), (cx + w1 / 2, y), (cx + w2 / 2, y + sh - 0.06),
                          (cx - w2 / 2, y + sh - 0.06)], grad=(c, _mix(c, NAVY_D, 0.25), 90))
            self.text(s, cx - w2 / 2, y, w2, sh - 0.06, st.get("value", f"{i + 1:02d}"), size=18, color=WHITE,
                      align="c", anchor="m", fit=False)
            right = i % 2 == 0
            ly = y + (sh - 0.06) / 2
            edge = (w1 + w2) / 4
            if right:
                self.line(s, cx + edge + 0.1, ly, cx + 2.9, ly, color=CARD_LINE, width=0.75)
                self.icon_dot(s, cx + 2.9, ly)
                self.text(s, cx + 3.15, ly - sh / 2 + 0.05, 3.4, 0.4, st.get("label", ""), size=16, color=NAVY,
                          bold=True, anchor="b", min_size=11)
                self.text(s, cx + 3.15, ly + 0.02, 3.4, sh / 2 - 0.05, st.get("text", ""), size=13, color=GRAY_D,
                          min_size=9.5)
            else:
                self.line(s, cx - 2.9, ly, cx - edge - 0.1, ly, color=CARD_LINE, width=0.75)
                self.icon_dot(s, cx - 2.9, ly)
                self.text(s, cx - 6.55, ly - sh / 2 + 0.05, 3.4, 0.4, st.get("label", ""), size=16, color=NAVY,
                          bold=True, anchor="b", align="r", min_size=11)
                self.text(s, cx - 6.55, ly + 0.02, 3.4, sh / 2 - 0.05, st.get("text", ""), size=13, color=GRAY_D,
                          align="r", min_size=9.5)
        return s

    def icon_dot(self, s, cx, cy, d=0.16, color=CYAN):
        return self.circle(s, cx - d / 2, cy - d / 2, d, fill=color)

    def pyramid(self, title, levels, note=None, notes=None):
        """Пирамида (вершина — первый элемент): уровни от тёмно-синего к циану, подписи справа.
        levels: [{label, text?}] 3–6 шт."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(levels)
        x0, base_w = L + 0.3, 5.6
        apex = x0 + base_w / 2
        y0, H = TOP + 0.1, BOT - TOP - 0.15
        lh = H / n
        for i, lv in enumerate(levels):
            ya, yb = y0 + i * lh, y0 + (i + 1) * lh - 0.07
            wa = base_w * (ya - y0) / H
            wb = base_w * (yb - y0) / H
            c = _mix(NAVY_D, CYAN, i / max(1, n - 1))
            if i == 0:
                pts = [(apex, ya), (apex + wb / 2, yb), (apex - wb / 2, yb)]
            else:
                pts = [(apex - wa / 2, ya), (apex + wa / 2, ya), (apex + wb / 2, yb), (apex - wb / 2, yb)]
            self.poly(s, pts, grad=(c, _mix(c, NAVY_D, 0.25), 90))
            self.text(s, apex - 0.4, ya + (0.15 if i == 0 else 0), 0.8, yb - ya - (0.1 if i == 0 else 0),
                      f"{i + 1}", size=18, color=WHITE, align="c", anchor="m" if i else "b", fit=False)
            my = (ya + yb) / 2
            self.line(s, apex + (wa + wb) / 4 + 0.1, my, x0 + base_w + 0.55, my, color=CARD_LINE)
            self.icon_dot(s, x0 + base_w + 0.55, my)
            tx = x0 + base_w + 0.85
            self.text(s, tx, ya - 0.02, R - tx, 0.42, lv.get("label", ""), size=17, color=NAVY, bold=True,
                      anchor="m" if not lv.get("text") else "t", min_size=11)
            if lv.get("text"):
                self.text(s, tx, ya + 0.4, R - tx, lh - 0.45, lv["text"], size=13, color=GRAY_D, min_size=9.5)
        return s

    def matrix(self, title, quadrants, x_axis=None, y_axis=None, highlight=None, note=None, notes=None):
        """Матрица 2×2 (приоритизация, BCG и т.п.). quadrants: 4×{title, text|items} в порядке
        верх-лево, верх-право, низ-лево, низ-право. x_axis/y_axis: подписи осей («Низкая → Высокая»)."""
        s = self.new_slide(title, note=note, notes=notes)
        ox = L + (0.55 if y_axis else 0)
        oy = TOP
        w = R - ox
        h = BOT - oy - (0.45 if x_axis else 0)
        gw, gh = (w - 0.2) / 2, (h - 0.2) / 2
        for i, q in enumerate(quadrants[:4]):
            r, c = divmod(i, 2)
            x, y = ox + c * (gw + 0.2), oy + r * (gh + 0.2)
            on = highlight is not None and i == highlight
            if on:
                self.shape(s, MSO_SHAPE.RECTANGLE, x, y, gw, gh, grad=(CYAN, "0091DA", 45))
            else:
                self.card(s, x, y, gw, gh)
            self.plate(s, x, y + 0.2, min(gw * 0.75, 4.2), 0.5, q.get("title", ""),
                       color="navy", size=16, text_color=WHITE)
            body = q.get("items") or q.get("text", "")
            self.text(s, x + 0.25, y + 0.9, gw - 0.5, gh - 1.05, body, size=14, bullets=bool(q.get("items")),
                      color=WHITE if on else TEXT, bullet_color=WHITE if on else NAVY, min_size=10,
                      hl_color=WHITE if on else NAVY)
        if y_axis:
            self.line(s, L + 0.2, oy + h, L + 0.2, oy, color=NAVY, width=1.5, arrow=True)
            tb = self.text(s, L - 1.95, oy + h / 2 - 0.2, 4.3, 0.4, y_axis, size=13, color=NAVY, align="c",
                           fit=False)
            tb.rotation = -90
        if x_axis:
            self.line(s, ox, BOT - 0.2, R, BOT - 0.2, color=NAVY, width=1.5, arrow=True)
            self.text(s, ox, BOT - 0.15, w, 0.3, x_axis, size=13, color=NAVY, align="c", fit=False)
        return s

    def swot(self, title, strengths, weaknesses, opportunities, threats, labels=None, note=None, notes=None):
        """SWOT-анализ: 4 карточки с буквенными плашками S/W/O/T."""
        s = self.new_slide(title, note=note, notes=notes)
        labels = labels or ["Сильные стороны", "Слабые стороны", "Возможности", "Угрозы"]
        data = [strengths, weaknesses, opportunities, threats]
        letters = "SWOT"
        colors = ["navy", "mid", "cyan", "gray"]
        gw, gh = (CW - 0.3) / 2, (BOT - TOP - 0.3) / 2
        for i in range(4):
            r, c = divmod(i, 2)
            x, y = L + c * (gw + 0.3), TOP + r * (gh + 0.3)
            self.card(s, x, y, gw, gh)
            self.plate(s, x, y, 1.15, gh, "", color=colors[i], slant="right")
            self.text(s, x + 0.15, y + 0.12, 0.8, 0.8, letters[i], size=44, color=WHITE, anchor="t", fit=False,
                      bold=True)
            self.text(s, x + 1.25, y + 0.15, gw - 1.45, 0.4, labels[i], size=17, color=NAVY, bold=True)
            self.text(s, x + 1.25, y + 0.62, gw - 1.45, gh - 0.75, data[i], size=13, bullets=True, min_size=9.5,
                      para_space=3)
        return s

    def comparison(self, title, left, right, arrow=True, note=None, notes=None):
        """Сравнение двух вариантов / «Было → Стало». left/right: {title, items: [str], icon?}.
        Левая колонка серая, правая — фирменная синяя/циан."""
        s = self.new_slide(title, note=note, notes=notes)
        gap = 1.0 if arrow else 0.35
        w = (CW - gap) / 2
        for j, blk in enumerate((left, right)):
            x = L + j * (w + gap)
            self.card(s, x, TOP, w, BOT - TOP)
            self.plate(s, x, TOP, w, 0.7, blk.get("title", ""), color="gray" if j == 0 else "navy", size=19)
            if blk.get("icon"):
                self.icon(s, blk["icon"], x + w - 0.75, TOP + 0.95, 0.5, "gray" if j == 0 else "cyan")
            self.text(s, x + 0.3, TOP + 1.0, w - 1.2 if blk.get("icon") else w - 0.55, BOT - TOP - 1.2,
                      blk.get("items", []), size=16, bullets=True, bullet_color=GRAY if j == 0 else CYAN,
                      min_size=11, para_space=8)
        if arrow:
            ax = L + w + 0.2
            ay = (TOP + BOT) / 2
            self.poly(s, [(ax, ay - 0.45), (ax + 0.25, ay - 0.45), (ax + 0.6, ay), (ax + 0.25, ay + 0.45),
                          (ax, ay + 0.45), (ax + 0.35, ay)], grad=(CYAN, NAVY, 0))
        return s

    def pricing(self, title, plans, note=None, notes=None):
        """Тарифы / пакеты (Pricing Table): колонки с ценой и списком опций. highlight — рекомендованный.
        plans: [{name, price, period?, items: [str], highlight?}] 2–4 шт."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(plans)
        gap = 0.3
        w = (CW - gap * (n - 1)) / n
        for i, p in enumerate(plans):
            x = L + i * (w + gap)
            hl = p.get("highlight")
            y = TOP + (0 if hl else 0.2)
            self.card(s, x, y, w, BOT - y)
            self.plate(s, x, y, w, 0.65, p.get("name", ""), color="cyan" if hl else "navy", size=18)
            self.text(s, x + 0.3, y + 0.85, w - 0.6, 0.85, str(p.get("price", "")), size=36,
                      color=CYAN if hl else NAVY, anchor="b", min_size=20)
            self.text(s, x + 0.3, y + 1.72, w - 0.6, 0.35, p.get("period", ""), size=13, color=GRAY)
            self.hline(s, x + 0.3, y + 2.15, w - 0.6, color=CARD_LINE)
            self.text(s, x + 0.3, y + 2.35, w - 0.6, BOT - y - 2.5, p.get("items", []), size=14, bullets=True,
                      bullet_color=CYAN, min_size=10, para_space=6)
        return s

    def hub(self, title, center, items, note=None, notes=None):
        """Центр + спутники (слайды 56–58): круг-центр с градиентом, кольца, 4–8 блоков вокруг.
        items: [{title, text?, icon?}]"""
        s = self.new_slide(title, note=note, notes=notes)
        cx, cy = SW / 2, (TOP + BOT) / 2 + 0.05
        for d, a in ((4.3, LINE), (3.4, "CFE9FA")):
            self.circle(s, cx - d / 2, cy - d / 2, d, line=a, line_w=0.75)
        D = 2.35
        self.circle(s, cx - D / 2, cy - D / 2, D, grad=(CYAN, NAVY, 45), shadow=True)
        self.text(s, cx - D / 2 + 0.25, cy - D / 2 + 0.25, D - 0.5, D - 0.5, center, size=18, color=WHITE,
                  align="c", anchor="m", min_size=12, bold=True)
        n = len(items)
        left = items[: (n + 1) // 2]
        right = items[(n + 1) // 2:]
        for side, group in ((-1, left), (1, right)):
            m = len(group)
            for j, it in enumerate(group):
                t = (j + 0.5) / m
                ang = math.radians(-60 + 120 * t)
                r_ = 2.15
                nx = cx + side * r_ * math.cos(ang)
                ny = cy + r_ * math.sin(ang) * 1.05
                nd = 0.62
                self.circle(s, nx - nd / 2, ny - nd / 2, nd, fill=WHITE, line=CYAN, line_w=1.25, shadow=True)
                if it.get("icon"):
                    self.icon(s, it["icon"], nx - 0.19, ny - 0.19, 0.38, "navy")
                else:
                    self.text(s, nx - nd / 2, ny - nd / 2, nd, nd, str(j + 1 + (0 if side < 0 else len(left))),
                              size=16, color=NAVY, align="c", anchor="m", fit=False)
                tw = 3.15
                tx = nx + 0.5 if side > 0 else nx - 0.5 - tw
                al = "l" if side > 0 else "r"
                self.text(s, tx, ny - 0.42, tw, 0.4, it.get("title", ""), size=15, color=NAVY, bold=True,
                          align=al, anchor="b", min_size=10)
                if it.get("text"):
                    self.text(s, tx, ny + 0.02, tw, 0.6, it["text"], size=12, color=GRAY_D, align=al, min_size=9)
        return s

    def cycle(self, title, steps, center=None, note=None, notes=None):
        """Цикл (PDCA, жизненный цикл клиента): 3–6 узлов на кольце со стрелками, описания снаружи.
        steps: [{title, text?}]"""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(steps)
        cx, cy = SW / 2, (TOP + BOT) / 2 + (0.35 if n % 2 else 0.05)
        Rr = 1.55 if n % 2 else 1.75
        self.circle(s, cx - Rr, cy - Rr, Rr * 2, line=CYAN, line_w=2.5)
        start = -math.pi / 2 + (0 if n % 2 else math.pi / n)
        if center:
            self.text(s, cx - 1.2, cy - 0.6, 2.4, 1.2, center, size=18, color=NAVY, bold=True, align="c",
                      anchor="m", min_size=11)
        for i, st in enumerate(steps):
            ang = start + 2 * math.pi * i / n
            nx, ny = cx + Rr * math.cos(ang), cy + Rr * math.sin(ang)
            # стрелка на кольце между узлами
            a2 = ang + math.pi / n
            axp, ayp = cx + Rr * math.cos(a2), cy + Rr * math.sin(a2)
            tri = self.shape(s, MSO_SHAPE.ISOSCELES_TRIANGLE, axp - 0.13, ayp - 0.13, 0.26, 0.26, fill=CYAN)
            tri.rotation = math.degrees(a2) + 180
            nd = 0.85
            c = _mix(CYAN, NAVY, i / max(1, n - 1))
            self.circle(s, nx - nd / 2, ny - nd / 2, nd, grad=(c, _mix(c, NAVY_D, 0.3), 45), shadow=True)
            self.text(s, nx - nd / 2, ny - nd / 2, nd, nd, f"{i + 1:02d}", size=20, color=WHITE, align="c",
                      anchor="m", fit=False, bold=True)
            tw = 3.3
            cosv = math.cos(ang)
            if abs(cosv) < 0.3:  # верхний узел — подпись над ним
                tx, al, ty = nx - tw / 2, "c", ny - 1.25
            elif cosv > 0:
                tx, al, ty = nx + 0.6, "l", ny - 0.45
            else:
                tx, al, ty = nx - 0.6 - tw, "r", ny - 0.45
            self.text(s, tx, ty, tw, 0.4, st.get("title", ""), size=16, color=NAVY, bold=True, align=al,
                      min_size=11)
            if st.get("text"):
                self.text(s, tx, ty + 0.42, tw, 0.75, st["text"], size=12, color=GRAY_D, align=al, min_size=9)
        return s

    def venn(self, title, sets, center=None, text=None, note=None, notes=None):
        """Диаграмма Венна для 2–3 множеств. sets: [{title, text?}], center — подпись пересечения."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(sets)
        D = 3.4
        cx, cy = L + 3.7, (TOP + BOT) / 2
        cols = [NAVY, CYAN, BLUE3]
        if n == 2:
            pos = [(cx - 0.95, cy), (cx + 0.95, cy)]
        else:
            pos = [(cx, cy - 0.85), (cx - 0.98, cy + 0.75), (cx + 0.98, cy + 0.75)]
            D = 3.0
        for i, (px, py) in enumerate(pos):
            self.circle(s, px - D / 2, py - D / 2, D, fill=cols[i], alpha=72)
        for i, (px, py) in enumerate(pos):
            dx = (px - cx) * 1.0
            dy = (py - cy) * 1.0
            self.text(s, px + dx * 0.6 - 1.0, py + dy * 0.55 - 0.3, 2.0, 0.6, sets[i].get("title", ""), size=15,
                      color=WHITE, bold=True, align="c", anchor="m", min_size=10)
        if center:
            self.text(s, cx - 0.65, cy - 0.15 + (0.2 if n == 3 else 0), 1.3, 0.5, center, size=12, color=WHITE,
                      bold=True, align="c", anchor="m", min_size=9)
        tx = L + 7.9
        y = TOP + 0.2
        rh = (BOT - TOP - 0.2) / n
        for i, st in enumerate(sets):
            self.circle(s, tx, y + i * rh + 0.08, 0.28, fill=cols[i])
            self.text(s, tx + 0.45, y + i * rh, R - tx - 0.45, 0.45, st.get("title", ""), size=17, color=NAVY,
                      bold=True, min_size=11)
            self.text(s, tx + 0.45, y + i * rh + 0.45, R - tx - 0.45, rh - 0.55, st.get("text", ""), size=13,
                      color=GRAY_D, min_size=9.5)
        return s

    def org_chart(self, title, root, children, note=None, notes=None):
        """Оргструктура в 3 уровня. root: {name, role}; children: [{name, role, children?: [{name, role}]}] (2–5)."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(children)
        rw, rh = 3.8, 1.0
        rx = SW / 2 - rw / 2
        self.plate(s, rx, TOP, rw, rh, "", color="navy", slant="none")
        self.text(s, rx + 0.15, TOP + 0.12, rw - 0.3, 0.45, root.get("name", ""), size=18, color=WHITE, bold=True,
                  align="c", min_size=11, anchor="m")
        self.text(s, rx + 0.15, TOP + 0.58, rw - 0.3, 0.32, root.get("role", ""), size=13, color=CYAN_L, align="c",
                  min_size=9)
        gap = 0.2
        cw = (CW - gap * (n - 1)) / n
        y1 = TOP + rh + 0.7
        bus = TOP + rh + 0.35
        self.vline(s, SW / 2, TOP + rh, bus - TOP - rh, color=NAVY, width=1.25)
        xs = [L + i * (cw + gap) + cw / 2 for i in range(n)]
        if n > 1:
            self.hline(s, xs[0], bus, xs[-1] - xs[0], color=NAVY, width=1.25)
        for i, ch in enumerate(children):
            x = L + i * (cw + gap)
            self.vline(s, xs[i], bus, y1 - bus, color=NAVY, width=1.25)
            self.shape(s, MSO_SHAPE.RECTANGLE, x, y1, cw, 0.95, grad=(CYAN, "0091DA", 45))
            self.text(s, x + 0.1, y1 + 0.1, cw - 0.2, 0.45, ch.get("name", ""), size=16, color=WHITE, bold=True,
                      align="c", min_size=10, anchor="m")
            self.text(s, x + 0.1, y1 + 0.55, cw - 0.2, 0.32, ch.get("role", ""), size=12, color=WHITE, align="c",
                      min_size=9)
            subs = ch.get("children", [])
            if subs:
                sy = y1 + 1.25
                avail = BOT - sy
                sh = min(0.85, (avail - 0.15 * (len(subs) - 1)) / len(subs))
                self.vline(s, x + 0.18, y1 + 0.95, (sy + (len(subs) - 0.5) * (sh + 0.15)) - (y1 + 0.95) - 0.075,
                           color=CARD_LINE, width=1)
                for j, sb in enumerate(subs):
                    yy = sy + j * (sh + 0.15)
                    self.hline(s, x + 0.18, yy + sh / 2, 0.17, color=CARD_LINE, width=1)
                    self.card(s, x + 0.35, yy, cw - 0.35, sh)
                    self.text(s, x + 0.5, yy + 0.06, cw - 0.6, sh * 0.5, sb.get("name", ""), size=14, color=NAVY,
                              bold=True, anchor="m", min_size=9)
                    self.text(s, x + 0.5, yy + sh * 0.55, cw - 0.6, sh * 0.4, sb.get("role", ""), size=12,
                              color=GRAY_D, min_size=8.5)
        return s

    def fishbone(self, title, effect, causes, note=None, notes=None):
        """Диаграмма Исикавы: проблема справа, 4–6 групп причин сверху/снизу.
        causes: [{title, items: [str]}]"""
        s = self.new_slide(title, note=note, notes=notes)
        sy = (TOP + BOT) / 2
        hx = R - 2.6
        self.line(s, L, sy, hx, sy, color=NAVY, width=2.5)
        self.plate(s, hx, sy - 0.75, R - hx, 1.5, effect, color="cyan", slant="left", size=16, pad=0.25)
        n = len(causes)
        top = causes[: (n + 1) // 2]
        bot = causes[(n + 1) // 2:]
        for row, group in ((-1, top), (1, bot)):
            m = len(group)
            span = (hx - L - 0.3) / max(1, m)
            for j, c in enumerate(group):
                bx = L + 0.3 + span * (j + 1) - 0.1
                ey = TOP + 0.6 if row < 0 else BOT - 0.6
                ex = bx - 1.0
                self.line(s, ex, ey, bx, sy, color=CYAN, width=1.5)
                ly = ey - 0.5 if row < 0 else ey + 0.02
                self.plate(s, ex - 0.6, ly, min(2.6, span - 0.15), 0.46, c.get("title", ""), color="navy", size=13)
                items = c.get("items", [])
                for k, it in enumerate(items[:3]):
                    t = (k + 1) / (len(items[:3]) + 1)
                    px = ex + (bx - ex) * t
                    py = ey + (sy - ey) * t
                    self.hline(s, px - 0.25, py, 0.25, color=CARD_LINE)
                    self.text(s, px - 2.05, py - 0.18, 1.78, 0.36, it, size=10.5, color=GRAY_D, align="r",
                              anchor="m", min_size=8)
        return s

    def iceberg(self, title, above, below, labels=("Видимая часть", "Скрытая часть"), note=None, notes=None):
        """Айсберг: видимые факторы над «водой», скрытые — под ней. above/below: [str]."""
        s = self.new_slide(title, note=note, notes=notes)
        wy = TOP + 1.7
        self.shape(s, MSO_SHAPE.RECTANGLE, 0, wy, SW, BOT + 0.2 - wy, fill="E5F3FC")
        cx = L + 3.1
        self.poly(s, [(cx, TOP + 0.05), (cx + 1.0, wy), (cx - 1.15, wy)], grad=(CYAN_L, CYAN, 90))
        self.poly(s, [(cx - 1.15, wy), (cx + 1.0, wy), (cx + 1.9, wy + 0.8), (cx + 2.5, wy + 2.2),
                      (cx + 1.5, BOT - 0.15), (cx + 0.3, BOT), (cx - 1.4, BOT - 0.3), (cx - 2.4, wy + 2.4),
                      (cx - 2.2, wy + 1.0)], grad=(BLUE2, NAVY_D, 90))
        self.hline(s, 0, wy, SW, color=CYAN, width=1.5)
        tx = L + 6.3
        self.text(s, tx, TOP, R - tx, 0.4, labels[0], size=17, color=NAVY, bold=True)
        self.text(s, tx, TOP + 0.45, R - tx, wy - TOP - 0.55, above, size=14, bullets=True, min_size=10)
        self.text(s, tx, wy + 0.2, R - tx, 0.4, labels[1], size=17, color=NAVY, bold=True)
        self.text(s, tx, wy + 0.65, R - tx, BOT - wy - 0.75, below, size=14, bullets=True, min_size=10)
        return s

    # ---------------------------------------------------------------- люди
    def team(self, title, people, note=None, notes=None):
        """Команда / спикеры (слайд 39): карточки с круглым фото (или инициалами), имя, роль, описание.
        people: [{name, role, text?, photo?}] 2–5 шт."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(people)
        gap = 0.3
        w = (CW - gap * (n - 1)) / n
        for i, p in enumerate(people):
            x = L + i * (w + gap)
            self.card(s, x, TOP, w, BOT - TOP)
            d = min(1.6, w - 0.8)
            px = x + (w - d) / 2
            py = TOP + 0.35
            self.circle(s, px - 0.1, py - 0.1, d + 0.2, line=CYAN, line_w=1.25)
            if p.get("photo") and os.path.exists(str(photo_path(p["photo"]))):
                self.picture(s, p["photo"], px, py, d, d, circle=True)
            else:
                self.circle(s, px, py, d, grad=(NAVY, NAVY_D, 45))
                ini = "".join(w_[0] for w_ in p.get("name", "?").split()[:2]).upper()
                self.text(s, px, py, d, d, ini, size=d * 22, color=WHITE, align="c", anchor="m", fit=False)
            ty = py + d + 0.3
            self.text(s, x + 0.2, ty, w - 0.4, 0.7, p.get("name", ""), size=17, color=TEXT, align="c",
                      min_size=12, bold=True)
            self.hline(s, x + w / 2 - 0.3, ty + 0.8, 0.6, color=CYAN, width=1.5)
            self.text(s, x + 0.2, ty + 0.92, w - 0.4, 0.55, p.get("role", ""), size=14, color=CYAN, align="c",
                      min_size=10)
            if p.get("text"):
                self.text(s, x + 0.25, ty + 1.5, w - 0.5, BOT - ty - 1.6, p["text"], size=12, align="c",
                          color=GRAY_D, min_size=9)
        return s

    # ---------------------------------------------------------------- таблицы и графики
    def table(self, title, header, rows, highlight_col=None, highlight_row=None, col_widths=None,
              group_rows=None, note=None, notes=None, font_size=None):
        """Таблица (слайды 40–46): синяя шапка, тонкие разделители, выделенный столбец/строка циановой заливкой.
        group_rows: индексы строк-подзаголовков (заливка циан, белый текст)."""
        s = self.new_slide(title, note=note, notes=notes)
        nr, nc = len(rows) + 1, len(header)
        avail = BOT - TOP
        rh = min(0.5, avail / nr)
        fs = font_size or (14 if rh >= 0.45 else 12 if rh >= 0.36 else 10)
        gf = s.shapes.add_table(nr, nc, E(L), E(TOP), E(CW), E(rh * nr))
        tbl = gf.table
        tblPr = tbl._tbl.tblPr
        for a in ("bandRow", "firstRow"):
            tblPr.set(a, "0")
        sid = tblPr.find(qn("a:tableStyleId"))
        if sid is not None:
            sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # No Style, Table Grid → переопределяем
        if col_widths:
            widths = col_widths
        elif nc > 1:
            fw = max(_measure(str(r[0]).strip("*"), fs, True) for r in [header] + list(rows)) + 0.35
            fw = min(max(fw, 2.2), CW * 0.45)
            widths = [fw] + [(CW - fw) / (nc - 1)] * (nc - 1)
        else:
            widths = [CW]
        tot = sum(widths)
        for j, wd in enumerate(widths):
            tbl.columns[j].width = E(wd * CW / tot)
        for i in range(nr):
            tbl.rows[i].height = E(rh)
            for j in range(nc):
                cell = tbl.cell(i, j)
                val = header[j] if i == 0 else rows[i - 1][j]
                hdr = i == 0
                grp = group_rows and (i - 1) in group_rows
                hlc = highlight_col is not None and j == highlight_col and not hdr
                hlr = highlight_row is not None and (i - 1) == highlight_row
                if hdr:
                    fill = NAVY if not (highlight_col is not None and j == highlight_col) else CYAN
                elif grp:
                    fill = CYAN
                elif hlc or hlr:
                    fill = CYAN_BG
                else:
                    fill = WHITE
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(fill)
                cell.margin_left = cell.margin_right = E(0.1)
                cell.margin_top = cell.margin_bottom = E(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                txt = str(val)
                bold = txt.startswith("**") and txt.endswith("**")
                p.text = ""
                r_ = p.add_run()
                r_.text = txt.strip("*") if bold else txt
                r_.font.size = Pt(fs)
                r_.font.name = HEAD if (hdr or bold or grp) else BODY
                r_.font.bold = bool(hdr or bold or grp) and self.font_mode == "safe"
                r_.font.color.rgb = rgb(WHITE if (hdr or grp) else (NAVY if (hlc or bold or j == 0) else TEXT))
                p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
                _cell_borders(cell, bottom=LINE if not hdr else NAVY)
        return s

    def chart(self, title, kind, categories, series, side=None, header=None, number_format='0',
              legend=None, labels=True, note=None, notes=None, colors=None, conclusion=None):
        """Нативная диаграмма PowerPoint в фирменных цветах (слайды 48–55).
        kind: column | bar | line | pie | doughnut | stacked | stacked_bar | area.
        series: {name: [values]}; side: [{value, label, icon?}] (2–3 KPI справа) или list[str] (буллеты);
        header — плашка над графиком; conclusion — строка вывода снизу."""
        s = self.new_slide(title, note=note, notes=notes)
        cw = CW if not side else CW - 3.55
        y = TOP
        bot = BOT - (0.85 if conclusion else 0)
        self.card(s, L, y, cw, bot - y)
        if header:
            self.plate(s, L, y, min(cw * 0.7, 6.5), 0.5, header, color="navy", size=15)
            y += 0.55
        types = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                 "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT,
                 "stacked": XL_CHART_TYPE.COLUMN_STACKED, "stacked_bar": XL_CHART_TYPE.BAR_STACKED,
                 "area": XL_CHART_TYPE.AREA}
        cd = CategoryChartData()
        cd.categories = categories
        for name, vals in series.items():
            cd.add_series(name, vals, number_format=number_format)
        gf = s.shapes.add_chart(types[kind], E(L + 0.15), E(y + 0.1), E(cw - 0.3), E(bot - y - 0.2), cd)
        ch = gf.chart
        _style_chart(ch, kind, len(series), colors or CHART_COLORS, labels, number_format,
                     legend if legend is not None else (len(series) > 1 or kind in ("pie", "doughnut")))
        if side:
            sx = L + cw + 0.3
            sw = R - sx
            if isinstance(side[0], dict):
                k = len(side)
                gh = (bot - TOP - 0.2 * (k - 1)) / k
                for i, it in enumerate(side):
                    yy = TOP + i * (gh + 0.2)
                    self.card(s, sx, yy, sw, gh)
                    self.icon(s, it.get("icon"), sx + 0.2, yy + 0.2, 0.4, "cyan")
                    self.text(s, sx + (0.75 if it.get("icon") else 0.25), yy + 0.12, sw - 0.95, 0.6,
                              str(it.get("value", "")), size=28, color=CYAN if i == 0 else NAVY, min_size=16,
                              anchor="m")
                    self.text(s, sx + 0.25, yy + 0.78, sw - 0.45, gh - 0.9, it.get("label", ""), size=13,
                              min_size=9.5)
            else:
                self.shape(s, MSO_SHAPE.RECTANGLE, sx, TOP, sw, bot - TOP, grad=(CYAN, "0091DA", 45))
                self.text(s, sx + 0.25, TOP + 0.25, sw - 0.5, bot - TOP - 0.5, side, size=14, color=WHITE,
                          bullets=True, bullet_color=WHITE, min_size=10, para_space=8, hl_color=WHITE)
        if conclusion:
            self.plate(s, L, BOT - 0.65, 2.3, 0.65, "Вывод", color="cyan", size=17)
            self.text(s, L + 2.4, BOT - 0.65, CW - 2.4, 0.65, conclusion, size=14, anchor="m", min_size=10)
        return s

    def waterfall(self, title, categories, values, totals=(), header=None, number_format='0', side=None,
                  conclusion=None, note=None, notes=None):
        """Каскадная диаграмма (Waterfall / факторный анализ). values — приращения; индексы из totals —
        итоговые столбцы (значение = уровень, а не приращение). Рост — циан, снижение — серый, итоги — синий."""
        base, shown, cols = [], [], []
        run = 0
        for i, v in enumerate(values):
            if i in totals:
                base.append(0)
                shown.append(v)
                cols.append(NAVY)
                run = v
            else:
                lo = run + min(0, v)
                base.append(lo)
                shown.append(abs(v))
                cols.append(CYAN if v >= 0 else GRAY)
                run += v
        s = self.new_slide(title, note=note, notes=notes)
        cw = CW if not side else CW - 3.55
        y = TOP
        bot = BOT - (0.85 if conclusion else 0)
        self.card(s, L, y, cw, bot - y)
        if header:
            self.plate(s, L, y, min(cw * 0.7, 6.5), 0.5, header, color="navy", size=15)
            y += 0.55
        cd = CategoryChartData()
        cd.categories = categories
        cd.add_series("base", base, number_format=number_format)
        cd.add_series("value", shown, number_format=number_format)
        gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, E(L + 0.15), E(y + 0.1), E(cw - 0.3),
                                E(bot - y - 0.2), cd)
        ch = gf.chart
        _style_chart(ch, "stacked", 2, [WHITE, NAVY], False, number_format, False)
        plot = ch.plots[0]
        plot.gap_width = 50
        plot.series[0].format.fill.background()
        vs = plot.series[1]
        for i, c in enumerate(cols):
            pt = vs.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(c)
            dl = pt.data_label
            dl.has_text_frame = True
            v = values[i]
            sign = "" if i in totals else ("+" if v >= 0 else "−")
            dl.text_frame.text = f"{sign}{abs(v):g}"
            dl.position = XL_LABEL_POSITION.INSIDE_END
            r_ = dl.text_frame.paragraphs[0].runs[0]
            r_.font.size = Pt(11)
            r_.font.bold = True
            r_.font.color.rgb = rgb(WHITE)
        if side:
            sx = L + cw + 0.3
            self.shape(s, MSO_SHAPE.RECTANGLE, sx, TOP, R - sx, bot - TOP, grad=(CYAN, "0091DA", 45))
            self.text(s, sx + 0.25, TOP + 0.25, R - sx - 0.5, bot - TOP - 0.5, side, size=14, color=WHITE,
                      bullets=True, bullet_color=WHITE, min_size=10, para_space=8, hl_color=WHITE)
        if conclusion:
            self.plate(s, L, BOT - 0.65, 2.3, 0.65, "Вывод", color="cyan", size=17)
            self.text(s, L + 2.4, BOT - 0.65, CW - 2.4, 0.65, conclusion, size=14, anchor="m", min_size=10)
        return s

    def dashboard(self, title, kpis, chart, bullets=None, note=None, notes=None):
        """Дашборд: ряд KPI сверху + график + блок выводов. kpis: [{value, label, delta?}] 3–4;
        chart: {kind, categories, series, header?}; bullets: list[str]."""
        s = self.new_slide(title, note=note, notes=notes)
        n = len(kpis)
        gap = 0.25
        kw = (CW - gap * (n - 1)) / n
        kh = 1.25
        for i, k in enumerate(kpis):
            x = L + i * (kw + gap)
            self.card(s, x, TOP, kw, kh)
            self.shape(s, MSO_SHAPE.RECTANGLE, x, TOP, 0.09, kh, fill=CYAN if i == 0 else NAVY)
            self.text(s, x + 0.3, TOP + 0.1, kw - 0.5, 0.65, str(k.get("value", "")), size=30,
                      color=CYAN if i == 0 else NAVY, anchor="m", min_size=16)
            d = k.get("delta")
            if d:
                pos = not str(d).strip().startswith(("-", "−"))
                self.text(s, x + kw - 1.25, TOP + 0.18, 1.05, 0.4, str(d), size=13, color=GREEN if pos else RED,
                          align="r", bold=True, fit=False)
            self.text(s, x + 0.3, TOP + 0.75, kw - 0.45, 0.45, k.get("label", ""), size=12, color=GRAY_D,
                      min_size=9)
        y = TOP + kh + 0.3
        cw = CW if not bullets else CW - 3.9
        self.card(s, L, y, cw, BOT - y)
        yy = y
        if chart.get("header"):
            self.plate(s, L, y, min(cw * 0.7, 6), 0.48, chart["header"], color="navy", size=14)
            yy += 0.5
        types = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                 "line": XL_CHART_TYPE.LINE_MARKERS, "stacked": XL_CHART_TYPE.COLUMN_STACKED,
                 "doughnut": XL_CHART_TYPE.DOUGHNUT, "pie": XL_CHART_TYPE.PIE, "area": XL_CHART_TYPE.AREA}
        cd = CategoryChartData()
        cd.categories = chart["categories"]
        for nm, vals in chart["series"].items():
            cd.add_series(nm, vals, number_format=chart.get("number_format", "0"))
        gf = s.shapes.add_chart(types[chart.get("kind", "column")], E(L + 0.15), E(yy + 0.08), E(cw - 0.3),
                                E(BOT - yy - 0.15), cd)
        _style_chart(gf.chart, chart.get("kind", "column"), len(chart["series"]), CHART_COLORS,
                     chart.get("labels", True), chart.get("number_format", "0"),
                     len(chart["series"]) > 1 or chart.get("kind") in ("pie", "doughnut"))
        if bullets:
            sx = L + cw + 0.3
            self.shape(s, MSO_SHAPE.RECTANGLE, sx, y, R - sx, BOT - y, grad=(NAVY, NAVY_D, 45))
            self.text(s, sx + 0.25, y + 0.2, R - sx - 0.5, 0.4, "Выводы", size=17, color=CYAN_L, bold=True)
            self.text(s, sx + 0.25, y + 0.7, R - sx - 0.5, BOT - y - 0.9, bullets, size=13, color=WHITE,
                      bullets=True, bullet_color=CYAN, min_size=9.5, para_space=6, hl_color=CYAN_L)
        return s

    # ---------------------------------------------------------------- сборка
    def save(self, path):
        self.prs.save(path)
        if self.font_mode == "safe":
            _rewrite_theme_fonts(path, "Arial", "Arial")
        for w in self.warnings:
            print("WARNING", w)
        return path


# ---------------------------------------------------------------- module helpers
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NSM = {"a": _NS_A, "p": _NS_P}
_ICON_LIB = None


def icon_library():
    """Векторная библиотека иконок ВТБ: {id: {w, h, xml}} (4 354 шт.)."""
    global _ICON_LIB
    if _ICON_LIB is None:
        import gzip
        import json
        with gzip.open(os.path.join(ASSETS, "icons_lib.json.gz"), "rt", encoding="utf8") as f:
            _ICON_LIB = json.load(f)
    return _ICON_LIB


def list_icons(category=None, style=None):
    """ID иконок, фильтр по категории (finance, tech, …) и стилю ('o' контур / 'g' заливка)."""
    out = []
    for k in icon_library():
        st, rest = k.split("-", 1)
        cat = rest.rsplit("-", 1)[0]
        if (category is None or cat == category) and (style is None or st == style):
            out.append(k)
    return sorted(out)

def _mix(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{int(round(a[i] + (b[i] - a[i]) * t)):02X}" for i in range(3))


def _cell_borders(cell, bottom=LINE, width=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    for tag, col in (("a:lnL", None), ("a:lnR", None), ("a:lnT", None), ("a:lnB", bottom)):
        ln = etree.SubElement(tcPr, qn(tag))
        if col:
            ln.set("w", str(int(width * 12700)))
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", col)
        else:
            ln.set("w", "0")
            etree.SubElement(ln, qn("a:noFill"))
    # lnX must precede fill in tcPr
    fills = [el for el in tcPr if el.tag in (qn("a:solidFill"), qn("a:noFill"))]
    for f in fills:
        tcPr.remove(f)
        tcPr.append(f)


def _style_chart(ch, kind, nseries, colors, labels, number_format, legend):
    ch.font.size = Pt(11)
    ch.font.name = BODY
    ch.font.color.rgb = rgb(GRAY_D)
    ch.has_title = False
    ch.has_legend = bool(legend)
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM if kind not in ("pie", "doughnut") else XL_LEGEND_POSITION.RIGHT
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(11)
    plot = ch.plots[0]
    if kind in ("pie", "doughnut"):
        ser = plot.series[0]
        for i, pt in enumerate(ser.points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(colors[i % len(colors)])
            pt.format.line.color.rgb = rgb(WHITE)
        if kind == "doughnut":
            hole = plot._element.find(qn("c:holeSize"))
            if hole is not None:
                hole.set("val", "60")
    else:
        if kind in ("column", "bar", "stacked", "stacked_bar"):
            plot.gap_width = 60 if nseries > 1 else 80
            if kind in ("stacked", "stacked_bar"):
                plot.overlap = 100
        for i, ser in enumerate(plot.series):
            c = colors[i % len(colors)]
            if kind == "line":
                ser.format.line.color.rgb = rgb(c)
                ser.format.line.width = Pt(2.5)
                ser.smooth = False
                ser.marker.format.fill.solid()
                ser.marker.format.fill.fore_color.rgb = rgb(c)
                ser.marker.format.line.color.rgb = rgb(WHITE)
                ser.marker.size = 8
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = rgb(c)
        va = ch.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = rgb("E6E6E6")
        va.major_gridlines.format.line.width = Pt(0.5)
        va.format.line.fill.background()
        va.tick_labels.font.size = Pt(10)
        va.tick_labels.font.color.rgb = rgb(GRAY)
        va.tick_labels.number_format = number_format
        va.tick_labels.number_format_is_linked = False
        ca = ch.category_axis
        ca.format.line.color.rgb = rgb(LINE)
        ca.tick_labels.font.size = Pt(11)
        ca.tick_labels.font.color.rgb = rgb(GRAY_D)
        ca.has_major_gridlines = False
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(10 if nseries > 2 else 11)
        dl.number_format = number_format if kind not in ("pie", "doughnut") else number_format
        dl.number_format_is_linked = False
        if kind in ("pie", "doughnut"):
            dl.show_value = True
            dl.show_category_name = False
            dl.show_percentage = False
            dl.font.color.rgb = rgb(WHITE)
            dl.font.bold = True
        elif kind in ("stacked", "stacked_bar"):
            dl.position = XL_LABEL_POSITION.CENTER
            dl.font.color.rgb = rgb(WHITE)
        elif kind == "line":
            dl.position = XL_LABEL_POSITION.ABOVE
            dl.font.color.rgb = rgb(NAVY)
            for i, ser in enumerate(plot.series):  # чередуем над/под линией, чтобы подписи не слипались
                sdl = ser.data_labels
                sdl.show_value = True
                sdl.number_format = number_format
                sdl.number_format_is_linked = False
                sdl.font.size = Pt(10)
                sdl.font.color.rgb = rgb(colors[i % len(colors)] if colors[i % len(colors)] != CYAN else "0088CC")
                sdl.position = XL_LABEL_POSITION.ABOVE if i % 2 == 0 else XL_LABEL_POSITION.BELOW
        elif kind in ("column", "bar"):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
            dl.font.color.rgb = rgb(NAVY)


def _rewrite_theme_fonts(path, major, minor):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/theme/theme") and item.filename.endswith(".xml"):
                x = data.decode("utf8")
                x = re.sub(r'(<a:majorFont><a:latin typeface=")[^"]*', r"\g<1>" + major, x)
                x = re.sub(r'(<a:minorFont><a:latin typeface=")[^"]*', r"\g<1>" + minor, x)
                data = x.encode("utf8")
            if item.filename.startswith(("ppt/slideLayouts/", "ppt/slideMasters/", "ppt/slides/")):
                x = data.decode("utf8")
                x = re.sub(r'typeface="VTB Group[^"]*"', f'typeface="{minor}"', x)
                data = x.encode("utf8")
            zout.writestr(item, data)
    os.replace(tmp, path)
