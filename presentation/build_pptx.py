#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SILA IQ — Executive PPTX generator
==================================
Rebuilds the SILA IQ (صِلة دوائي) marketing strategy deck as a 16:9
executive / human-crafted .pptx, optimized for 100% Google Slides
compatibility.

Design system
-------------
- Canvas ............ 13.333" x 7.5" (true 16:9, EMU 12192000 x 6858000)
- Background ........ F8FAFC (off-white) — cover & closing: 0A192F navy
- Ink ............... 0F172A   Sub 64748B   Line E2E8F0
- Accent ............ 0D9488 (medical teal)  Tints F0FDFA / 99F6E4
- Secondary ......... E11D48 (rose)  F59E0B (amber)  Tints FFF1F2 / FFF7ED
- Font .............. Cairo (latin + complex-script + EA all set) — a Google
  Font, so Google Slides renders it natively.
- RTL ............... every paragraph gets a:pPr rtl="1", runs get lang="ar".
- Geometry .......... flat white rounded rectangles, 1pt E2E8F0 hairlines,
  no shadows, no gradients, generous margins (0.6" gutters).

Layout diversity (no bullet-only slides)
----------------------------------------
 1 Cover (navy minimalist)          8  90-day phase timeline (4 cols)
 2 Problem & opportunity (3 cards)  9  Two-sided engine (split columns)
 3 System triangle (diagram+cards) 10  Demand engine (channels + 5 cards)
 4 Positioning (5 pillars)         11  KPI big-number blocks (8)
 5 Segments (2x2 cards)            12  Decisions (numbered rows)
 6 Competition (native table)      13  7-day plan (executive table)
 7 Trust & compliance (3x2 cards) 14  Closing (navy bookend)

Run:  python3 build_pptx.py   ->  docs/silla-iq-executive-presentation.pptx
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn
from lxml import etree

# ---------------------------------------------------------------- tokens
FONT = "Cairo"
BG      = "F8FAFC"   # off-white
NAVY    = "0A192F"   # cover navy
NAVY2   = "16233B"   # cover surface
NAVYLN  = "243B55"   # cover hairline
INK     = "0F172A"   # headings
SUB     = "64748B"   # secondary text
FAINT   = "94A3B8"   # faint text
LINE    = "E2E8F0"   # hairlines
TEAL    = "0D9488"   # primary accent
TEALDK  = "0F766E"   # dark teal text
TEALSOFT= "F0FDFA"   # teal tint
TEALMD  = "CCFBF1"   # teal mid tint
TEALLT  = "5EEAD4"   # light teal (on navy)
ROSE    = "E11D48"   # secondary accent
ROSEDK  = "9F1239"
ROSESOFT= "FFF1F2"
AMBER   = "F59E0B"
AMBERDK = "B45309"
AMBERSOFT = "FFFBEB"
ZEBRA   = "F8FAFC"
WHITE   = "FFFFFF"

W, H = 12192000, 6858000           # EMU
MX = 0.6                           # side margin (in)
CW = 13.333 - 2 * MX               # content width
Y0 = 1.86                          # content top
FY = 7.06                          # footer line y

AL = {"r": PP_ALIGN.RIGHT, "c": PP_ALIGN.CENTER, "l": PP_ALIGN.LEFT}
ANCH = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


def C(hexstr):
    return RGBColor.from_string(hexstr)


# ---------------------------------------------------------------- helpers
def _fix_run(run, size, bold, color, italic=False):
    """Set font on a run incl. complex-script (a:cs) + (a:ea) typeface,
    so Arabic glyphs actually use Cairo in both PPT and Google Slides."""
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = C(color)
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", "ar")
    latin = rPr.find(qn("a:latin"))
    if latin is None:
        latin = etree.SubElement(rPr, qn("a:latin"))
    latin.set("typeface", FONT)
    for tag in ("a:ea", "a:cs", "a:sym"):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    prev = latin
    for tag in ("a:ea", "a:cs"):
        el = etree.Element(qn(tag))
        el.set("typeface", FONT)
        prev.addnext(el)
        prev = el


def _para(p, align, rtl=True, line=None, before=None, after=None):
    p.alignment = AL[align]
    pPr = p._p.get_or_add_pPr()
    if rtl:
        pPr.set("rtl", "1")
    if line is not None:
        p.line_spacing = line
    if before is not None:
        p.space_before = Pt(before)
    if after is not None:
        p.space_after = Pt(after)


def add_text(slide, x, y, w, h, paras, anchor="t", wrap=True):
    """paras: list of dicts {runs:[(text,{size,color,bold,italic})],
    align, line, before, after}"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = ANCH[anchor]
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    first = True
    for spec in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        _para(p, spec.get("align", "r"),
              line=spec.get("line"), before=spec.get("before"),
              after=spec.get("after"))
        for text, st in spec["runs"]:
            r = p.add_run()
            r.text = text
            _fix_run(r, st.get("size", 12), st.get("bold", False),
                     st.get("color", INK), st.get("italic", False))
    return tb


def P(runs, align="r", line=None, before=None, after=None):
    return {"runs": runs, "align": align, "line": line,
            "before": before, "after": after}


def R(text, size=12, color=INK, bold=False, italic=False):
    return (text, {"size": size, "color": color, "bold": bold, "italic": italic})


def add_shape(slide, x, y, w, h, fill=WHITE, line_color=LINE, line_w=1.0,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07, shadow=False):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = C(fill)
    if line_color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = C(line_color)
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE and radius is not None:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    tf = sp.text_frame
    tf.word_wrap = True
    pPr = tf.paragraphs[0]._p.get_or_add_pPr()
    pPr.set("rtl", "1")
    return sp


def add_line(slide, x1, y1, x2, y2, color=TEAL, w=1.5, arrow=False, dash=None):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = C(color)
    conn.line.width = Pt(w)
    conn.shadow.inherit = False
    ln = conn.line._get_or_add_ln()
    if dash:
        d = etree.SubElement(ln, qn("a:prstDash"))
        d.set("val", dash)
    if arrow:
        t = etree.SubElement(ln, qn("a:tailEnd"))
        t.set("type", "arrow")
        t.set("w", "med")
        t.set("len", "med")
    return conn


# ---------------------------------------------------------------- chrome
def new_slide(prs, bg=BG):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_shape(s, -0.02, -0.02, 13.38, 7.56, fill=bg, line_color=None,
              shape=MSO_SHAPE.RECTANGLE)
    return s


def header(slide, num, label, title, sub=None, tsize=25):
    add_text(slide, MX, 0.44, CW, 0.32,
             [P([R(f"{num}  •  {label}", 11, TEAL, True)])])
    add_text(slide, MX, 0.76, CW, 0.62,
             [P([R(title, tsize, INK, True)], line=1.05)])
    if sub:
        add_text(slide, MX, 1.40, CW, 0.34, [P([R(sub, 12.5, SUB)])])


def footer(slide, i, total, navy=False):
    c = "64748B" if navy else FAINT
    add_shape(slide, MX, FY, CW, 0.012,
              fill=NAVYLN if navy else LINE, line_color=None,
              shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, MX, FY + 0.08, CW * 0.5, 0.25,
             [P([R("صِلة دوائي  •  SILA IQ", 8.5, c, True)])])
    add_text(slide, MX + CW * 0.5, FY + 0.08, CW * 0.5, 0.25,
             [P([R(f"{i:02d} / {total:02d}", 8.5, c, True)], align="l")])


def card(slide, x, y, w, h, fill=WHITE, line=LINE, radius=0.07, lw=1.0):
    return add_shape(slide, x, y, w, h, fill=fill, line_color=line,
                     line_w=lw, radius=radius)


def chip(slide, x, y, w, h, text, tcolor=TEALDK, fill=TEALSOFT, line=None,
         size=10, bold=True, radius=0.5):
    add_shape(slide, x, y, w, h, fill=fill, line_color=line, radius=radius)
    add_text(slide, x + 0.08, y, w - 0.16, h,
             [P([R(text, size, tcolor, bold)], align="c")], anchor="m")


def dot_line(slide, x, y, w, text, size=10.5, color=SUB, dot=TEAL,
             line_spacing=1.12, before=4):
    add_text(slide, x, y, w, 0.34,
             [P([R("•  ", size, dot, True), R(text, size, color)], line=line_spacing,
                before=before)])


def tag(slide, x, y, w, text, fill=TEALSOFT, tcolor=TEALDK):
    """small section tag on top of a card"""
    add_shape(slide, x, y, w, 0.3, fill=fill, line_color=None, radius=0.5)
    add_text(slide, x + 0.05, y, w - 0.1, 0.3,
             [P([R(text, 9.5, tcolor, True)], align="c")], anchor="m")


# ---------------------------------------------------------------- table
NO_GRID_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def make_table(slide, x, y, w, h, rows, cols, col_widths, row_heights):
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y),
                                Inches(w), Inches(h))
    tbl = gf.table
    tblPr = gf._element.graphic.graphicData.tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = NO_GRID_STYLE
    for i, cw in enumerate(col_widths):
        tbl.columns[i].width = Inches(cw)
    for i, rh in enumerate(row_heights):
        tbl.rows[i].height = Inches(rh)
    return tbl


def cell_border_bottom(cell, color=LINE, w_pt=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    for el in tcPr.findall(qn("a:lnB")):
        tcPr.remove(el)
    ln = etree.Element(qn("a:lnB"))
    ln.set("w", str(int(Pt(w_pt))))
    ln.set("cap", "flat")
    sf = etree.SubElement(ln, qn("a:solidFill"))
    clr = etree.SubElement(sf, qn("a:srgbClr"))
    clr.set("val", color)
    tcPr.insert(0, ln)


def set_cell(cell, paras, fill=WHITE, anchor="m", ml=0.14, mr=0.14):
    cell.fill.solid()
    cell.fill.fore_color.rgb = C(fill)
    cell.vertical_anchor = ANCH[anchor]
    cell.margin_left = Inches(ml)
    cell.margin_right = Inches(mr)
    cell.margin_top = Inches(0.03)
    cell.margin_bottom = Inches(0.03)
    tf = cell.text_frame
    tf.word_wrap = True
    first = True
    for spec in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        _para(p, spec.get("align", "r"),
              line=spec.get("line"), before=spec.get("before"),
              after=spec.get("after"))
        for text, st in spec["runs"]:
            r = p.add_run()
            r.text = text
            _fix_run(r, st.get("size", 11), st.get("bold", False),
                     st.get("color", INK), st.get("italic", False))


# ================================================================ slides
TOTAL = 14


def s01_cover(prs):
    s = new_slide(prs, bg=NAVY)
    # quiet geometry (human touch, not decoration spam)
    ring = add_shape(s, 1.15, 1.0, 1.7, 1.7, fill=None, line_color=NAVYLN,
                     line_w=1.2, shape=MSO_SHAPE.OVAL, radius=None)
    add_shape(s, 10.9, 5.7, 0.14, 0.14, fill=TEAL, line_color=None,
              shape=MSO_SHAPE.OVAL, radius=None)
    for i in range(4):
        add_shape(s, 1.05 + i * 0.24, 6.35, 0.055, 0.055, fill=NAVYLN,
                  line_color=None, shape=MSO_SHAPE.OVAL, radius=None)

    # logo (top-right, RTL start)
    add_shape(s, 11.62, 0.55, 0.55, 0.55, fill=TEAL, line_color=None, radius=0.30)
    add_text(s, 11.62, 0.55, 0.55, 0.55,
             [P([R("+", 22, WHITE, True)], align="c")], anchor="m")
    add_text(s, 9.62, 0.52, 1.92, 0.34, [P([R("صِلة دوائي", 15, "F8FAFC", True)])])
    add_text(s, 9.62, 0.86, 1.92, 0.24, [P([R("S I L A  I Q", 9, FAINT, True)])])

    # kicker
    add_shape(s, 6.73, 1.98, 6.0, 0.44, fill=None, line_color=NAVYLN, line_w=1.0, radius=0.5)
    add_text(s, 6.73, 1.98, 6.0, 0.44,
             [P([R("الخطة التسويقية والاستراتيجية للنمو  •  السوق العراقي", 11, TEALLT, True)],
                align="c")], anchor="m")
    # accent rule
    add_shape(s, 12.13, 2.68, 0.55, 0.045, fill=TEAL, line_color=None,
              shape=MSO_SHAPE.RECTANGLE)
    # title
    add_text(s, 0.9, 2.95, 11.83, 1.9,
             [P([R("من تطبيق توصيل", 42, "F8FAFC", True)], line=1.12),
              P([R("إلى شبكة دواء موثوقة", 42, TEALLT, True)], line=1.12)])
    # promise
    add_text(s, 0.9, 5.02, 11.83, 0.4,
             [P([R("«دواؤك الموثوق، من صيدلية قريبة، وبالسعر والوقت الواضحين.»",
                  15, "94A3B8")])])
    # meta chips
    chips = [("قرار إطلاق  •  90 يوماً", 2.35),
             ("بغداد أولاً ثم البصرة / النجف / كربلاء", 3.6),
             ("أيلول 2026", 1.55)]
    cx = 12.73
    for text, w in chips:
        cx -= w
        add_shape(s, cx, 5.62, w, 0.42, fill=NAVY2, line_color=NAVYLN, line_w=1.0, radius=0.5)
        add_text(s, cx + 0.05, 5.62, w - 0.1, 0.42,
                 [P([R(text, 10.5, "B6C2D4", True)], align="c")], anchor="m")
        cx -= 0.22
    add_text(s, 0.9, 6.9, 6.0, 0.3,
             [P([R("SILA IQ — Executive Deck", 9, "475569", True)], align="l")])
    footer(s, 1, TOTAL, navy=True)


def s02_problem(prs):
    s = new_slide(prs)
    header(s, "01", "الفرصة", "رحلة المريض اليوم مشتّتة — والفجوة قابلة للامتلاك",
           sub="الاكتشاف يبدأ بالهاتف ووسائل التواصل، والإتمام بالنقد عند الاستلام — بينما الثقة والتنسيق لم يكتملا.")
    cw_, gap = 3.91, 0.2
    xs = [MX + CW - (3 - i) * (cw_ + gap) for i in range(3)]  # right→left order
    y, h = 1.92, 4.28

    # card A — problem journey
    card(s, xs[0], y, cw_, h)
    tag(s, xs[0] + cw_ - 1.75, y + 0.22, 1.5, "المشكلة")
    add_text(s, xs[0] + 0.26, y + 0.62, cw_ - 0.52, 0.4,
             [P([R("خمس حواجز في كل طلب دواء", 14.5, INK, True)])])
    steps = ["اتصال وانتظار أرقام متفرقة",
             "مشاوير متعددة بين الصيدليات",
             "«مو متوفر» — لا مخزون مضمون",
             "نقد عند الاستلام بلا وضوح مسبق",
             "قلق من مصدر الدواء وسلسلة الحفظ"]
    ly = y + 1.18
    for i, t in enumerate(steps, 1):
        add_shape(s, xs[0] + cw_ - 0.62, ly + 0.03, 0.3, 0.3, fill=TEALSOFT,
                  line_color=None, radius=0.5)
        add_text(s, xs[0] + cw_ - 0.62, ly + 0.03, 0.3, 0.3,
                 [P([R(str(i), 10.5, TEALDK, True)], align="c")], anchor="m")
        add_text(s, xs[0] + 0.26, ly, cw_ - 1.05, 0.34,
                 [P([R(t, 11.5, "475569")])], anchor="m")
        ly += 0.555

    # card B — why now
    card(s, xs[1], y, cw_, h)
    tag(s, xs[1] + cw_ - 1.75, y + 0.22, 1.5, "لماذا الآن")
    add_text(s, xs[1] + 0.26, y + 0.62, cw_ - 0.52, 0.4,
             [P([R("تقاطع ثلاث إشارات سوقية", 14.5, INK, True)])])
    why = [("سوق رقمي كثيف", "الاكتشاف يبدأ من فيسبوك وإنستغرام وتيك توك"),
           ("النقد مرساة الثقة", "السعر النهائي والوقت يجب أن يظهر قبل التأكيد"),
           ("فجوات مفتوحة", "التوفّر والتنسيق والثقة لم تُغلق بعد — لا حاجة لاختراع سوق")]
    ly = y + 1.22
    for lead, body in why:
        add_shape(s, xs[1] + cw_ - 0.5, ly + 0.06, 0.09, 0.09, fill=TEAL,
                  line_color=None, shape=MSO_SHAPE.OVAL, radius=None)
        add_text(s, xs[1] + 0.26, ly, cw_ - 0.72, 0.3,
                 [P([R(lead, 12, INK, True)])])
        add_text(s, xs[1] + 0.26, ly + 0.31, cw_ - 0.72, 0.5,
                 [P([R(body, 10.5, SUB)], line=1.15)])
        ly += 1.02

    # card C — gaps to own
    card(s, xs[2], y, cw_, h)
    tag(s, xs[2] + cw_ - 1.95, y + 0.22, 1.7, "نستطيع امتلاكها")
    add_text(s, xs[2] + 0.26, y + 0.62, cw_ - 0.52, 0.4,
             [P([R("خمس فجوات عراقية محددة", 14.5, INK, True)])])
    gaps = ["فجوة التوفّر — من يملك الدواء الآن؟",
            "فجوة التنسيق — الهاتف والواتساب تشتت الطلب",
            "فجوة الثقة — لا مرخّص ولا صيدلي ظاهر",
            "فجوة الوصول — راعي مريض مزمّن أو كبير",
            "فجوة الاختيار — مقارنة بلا مزاد على الصيدلية"]
    ly = y + 1.22
    for g in gaps:
        add_shape(s, xs[2] + 0.26, ly, cw_ - 0.52, 0.52, fill=BG,
                  line_color=LINE, line_w=0.75, radius=0.18)
        add_text(s, xs[2] + 0.42, ly, cw_ - 0.8, 0.52,
                 [P([R(g, 10.5, "475569", True)], line=1.1)], anchor="m")
        ly += 0.62

    # bottom strip
    add_shape(s, MX, 6.36, CW, 0.5, fill=ROSESOFT, line_color=None, radius=0.16)
    add_shape(s, MX + CW - 0.09, 6.42, 0.045, 0.38, fill=ROSE, line_color=None,
              shape=MSO_SHAPE.RECTANGLE)
    add_text(s, MX + 0.3, 6.36, CW - 0.7, 0.5,
             [P([R("الدواء فئة عالية الحساسية: ", 11.5, ROSEDK, True),
                 R("إثبات الترخيص والتحقق من الوصفة وسلسلة الحفظ شروط دخول قبل أي إعلان أداء.",
                   11.5, ROSEDK)], line=1.1)], anchor="m")
    footer(s, 2, TOTAL)


def s03_system(prs):
    s = new_slide(prs)
    header(s, "02", "المنصة", "نظام واحد يربط المريض والصيدلية والكابتن")

    # ---- triangle geometry
    P_x, P_y = 6.667, 2.44      # patient (top)
    F_x, F_y = 10.55, 4.52      # pharmacy (bottom-right)
    K_x, K_y = 2.75, 4.52       # captain (bottom-left)
    d = 0.98

    # edges (with arrows): patient→pharmacy, pharmacy→captain, captain→patient
    import math
    def edge(x1, y1, x2, y2, off=0.56):
        dx, dy = x2 - x1, y2 - y1
        L = math.hypot(dx, dy)
        ux, uy = dx / L, dy / L
        return (x1 + ux * off, y1 + uy * off, x2 - ux * off, y2 - uy * off)
    for (x1, y1, x2, y2) in ((P_x, P_y, F_x, F_y),
                             (F_x, F_y, K_x, K_y),
                             (K_x, K_y, P_x, P_y)):
        add_line(s, *edge(x1, y1, x2, y2), color="99F6E4", w=1.75, arrow=True)

    # core
    cx, cy, cd = 6.667, 3.78, 1.42
    add_shape(s, cx - cd / 2, cy - cd / 2, cd, cd, fill=WHITE, line_color=TEAL,
              line_w=1.5, shape=MSO_SHAPE.OVAL, radius=None)
    add_text(s, cx - 0.6, cy - 0.34, 1.2, 0.66,
             [P([R("قلب التشغيل", 10.5, INK, True)], align="c", line=1.05),
              P([R("والصيدلي", 10.5, INK, True)], align="c", line=1.05)])
    add_text(s, cx - 0.6, cy + 0.30, 1.2, 0.24,
             [P([R("طبقة الثقة", 8, TEALDK, True)], align="c")])

    # vertex circles
    verts = [(P_x, P_y, "المريض"), (F_x, F_y, "الصيدلية"), (K_x, K_y, "الكابتن")]
    for vx, vy, name in verts:
        add_shape(s, vx - d / 2, vy - d / 2, d, d, fill=WHITE, line_color=TEAL,
                  line_w=1.5, shape=MSO_SHAPE.OVAL, radius=None)
        add_text(s, vx - d / 2, vy - d / 2, d, d,
                 [P([R(name, 10.5, TEALDK, True)], align="c")], anchor="m")

    # edge labels (with bg masks)
    def elabel(x, y, text):
        add_shape(s, x - 0.85, y - 0.14, 1.7, 0.28, fill=BG, line_color=None,
                  shape=MSO_SHAPE.RECTANGLE)
        add_text(s, x - 0.85, y - 0.14, 1.7, 0.28,
                 [P([R(text, 8.5, SUB, True)], align="c")], anchor="m")
    elabel(8.62, 3.30, "طلب • وصفة")
    elabel(4.70, 3.30, "توصيل • COD")
    elabel(6.667, 4.78, "تجهيز • تسليم")

    # ---- value cards under vertices
    cw_, gap = 3.91, 0.2
    xs = [MX + CW - (3 - i) * (cw_ + gap) for i in range(3)]
    y, h = 5.18, 1.68
    data = [
        ("المريض / العائلة", "B2C", ["بحث ورفع الوصفة ومقارنة العروض", "اختيار بالسعر والمسافة والوقت", "تتبّع الطلب والدفع عند الاستلام"]),
        ("الصيدلية", "B2B", ["قناة طلبات محلية ولوحة إدارة", "تحكم بالسعر والعروض والموافقة", "سجل طلبات وتسويات شفافة"]),
        ("الكابتن", "التوصيل", ["طلبات ضمن نطاقه وخريطة واضحة", "سجل أرباح واعتماد ودعم", "تسوية يومية ودفع في وقته"]),
    ]
    # order: patient (center x), pharmacy (right), captain (left)
    card_x = {0: xs[1], 1: xs[0], 2: xs[2]}
    for i, (name, sub_, lines) in enumerate(data):
        x = card_x[i]
        card(s, x, y, cw_, h)
        add_shape(s, x + cw_ - 0.95, y + 0.17, 0.78, 0.28, fill=TEALSOFT,
                  line_color=None, radius=0.5)
        add_text(s, x + cw_ - 0.95, y + 0.17, 0.78, 0.28,
                 [P([R(sub_, 8.5, TEALDK, True)], align="c")], anchor="m")
        add_text(s, x + 0.24, y + 0.16, cw_ - 1.3, 0.3,
                 [P([R(name, 13, INK, True)])])
        ly = y + 0.58
        for t in lines:
            dot_line(s, x + 0.26, ly, cw_ - 0.5, t, size=10, before=0)
            ly += 0.34
    footer(s, 3, TOTAL)


def s04_positioning(prs):
    s = new_slide(prs)
    header(s, "03", "التموضع", "شبكة صيدليات موثّقة — لا متجر واحد ولا دليفري عام")
    add_text(s, MX, 1.42, CW, 0.34,
             [P([R("خمسة أعمدة تعنيها هذه العبارة، ويقيسها المستخدم في كل طلب.", 12.5, SUB)])])
    pillars = [
        ("ثقة", "اسم الصيدلية المرخّصة وتاريخ الاعتماد والصيدلي المتابع"),
        ("توفّر", "«متاح / يحتاج تأكيداً» مع وقت تحديث واضح لكل مخزون"),
        ("اختيار", "عروض متعددة مقارنة بالسعر والمسافة والوقت"),
        ("صيدلي", "يراجع الوصفة ولا يستبدل المادة دون موافقته"),
        ("تتبع", "توصيل قابل للتتبع وإيصال دقيق وتسوية شفافة"),
    ]
    gap = 0.22
    w = (CW - 4 * gap) / 5
    y, h = 2.05, 2.42
    for i, (name, desc) in enumerate(pillars):
        x = MX + CW - w - i * (w + gap)
        card(s, x, y, w, h)
        add_shape(s, x + w - 0.62, y + 0.22, 0.4, 0.4, fill=TEALSOFT,
                  line_color=None, radius=0.28)
        add_text(s, x + w - 0.62, y + 0.22, 0.4, 0.4,
                 [P([R(name[0], 15, TEALDK, True)], align="c")], anchor="m")
        add_text(s, x + 0.18, y + 0.24, w - 0.8, 0.34,
                 [P([R(name, 13.5, INK, True)])])
        add_shape(s, x + 0.2, y + 0.72, 0.34, 0.03, fill=TEAL, line_color=None,
                  shape=MSO_SHAPE.RECTANGLE)
        add_text(s, x + 0.18, y + 0.9, w - 0.36, h - 1.1,
                 [P([R(desc, 10, SUB)], line=1.25)])
    # banner
    add_shape(s, MX, 4.78, CW, 0.62, fill=TEALSOFT, line_color=None, radius=0.14)
    add_text(s, MX + 0.32, 4.78, CW - 0.64, 0.62,
             [P([R("نُظهر من يملك الدواء ", 13, INK, True),
                 R("الآن", 13, TEALDK, True),
                 R(" — مع وقت التحديث لكل مخزون.", 13, INK, True)])], anchor="m")
    # promise boundary
    add_text(s, MX, 5.62, CW, 0.5,
             [P([R("حدود الوعد:  ", 10.5, FAINT, True),
                 R("صِلة تُسهّل الوصول ولا تشخّص ولا تستبدل الطبيب أو الصيدلي، ولا تُسوَّق كخدمة إسعاف أو بديل للطوارئ.",
                   10.5, FAINT)], line=1.25)])
    add_text(s, MX, 6.18, CW, 0.4,
             [P([R("لا حرب «أسرع دائماً» أو «أرخص دائماً» قبل توفّر البيانات التشغيلية.",
                   10.5, SUB, True)])])
    footer(s, 4, TOTAL)


def s05_segments(prs):
    s = new_slide(prs)
    header(s, "04", "الجمهور", "أربع شرائح ذات أولوية — وأربع رسائل مختلفة",
           sub="لكل شريحة «وظيفة» تشتريها من صِلة، ومخاوفها وقنواتها الخاصة — لا حملة واحدة للجميع.")
    segs = [
        ("مدير / ة علاج الأسرة", "يرفع الدواء للأب والأم والأطفال مع سجل طلبات وعنوان محفوظ",
         "المخاوف: الاعتمادية", "فيسبوك + واتساب"),
        ("المريض المزمن / مقدم الرعاية", "إعادة طلب شهرية ثابتة وحساسية عالية للتوفّر والبديل الآمن",
         "المخاوف: التوفّر الدائم", "تذكّرات + واتساب"),
        ("المستخدم الرقمي الشاب", "OTC وعناية وفيتامينات — قابل للتجربة والتوصية والتفاعل",
         "المخاوف: تجربة سلسة", "تيك توك + إنستغرام"),
        ("الموظف / الطالب الحضري", "وقت محدود — طلب بسيط ودفع نقدي دون عناء",
         "المخاوف: السرعة والزمن", "جغرافي + بحث"),
    ]
    w, h, gap = 5.96, 2.14, 0.21
    pos = [(MX + CW - w, 1.95), (MX, 1.95), (MX + CW - w, 4.24), (MX, 4.24)]
    for (name, job, worry, chan), (x, y) in zip(segs, pos):
        card(s, x, y, w, h)
        add_shape(s, x + w - 0.16, y + 0.3, 0.05, h - 0.6, fill=TEAL,
                  line_color=None, shape=MSO_SHAPE.RECTANGLE)
        add_text(s, x + 0.26, y + 0.2, w - 0.6, 0.34,
                 [P([R(name, 14, INK, True)])])
        add_text(s, x + 0.26, y + 0.62, w - 0.62, 0.62,
                 [P([R(job, 11, SUB)], line=1.25)])
        chip(s, x + w - 2.62, y + h - 0.52, 1.5, 0.34, worry, size=9)
        chip(s, x + w - 4.28, y + h - 0.52, 1.56, 0.34, chan,
             tcolor=TEALDK, fill=TEALSOFT, size=9)
    add_shape(s, MX, 6.58, CW, 0.4, fill=BG, line_color=LINE, line_w=0.75, radius=0.5)
    add_text(s, MX + 0.3, 6.58, CW - 0.6, 0.4,
             [P([R("خارج مركز المدينة:  ", 10, TEALDK, True),
                 R("قيمة عالية لكن اقتصاديات التوصيل قد تكون سالبة — يؤجَّل لمرحلة لاحقة أو نقطة استلام.",
                   10, SUB)], align="c")], anchor="m")
    footer(s, 5, TOTAL)


def s06_competition(prs):
    s = new_slide(prs)
    header(s, "05", "المنافسة", "المنافس الحقيقي هو العادة الحالية — لا تطبيق واحد",
           sub="مصفوفة القدرات الست التي يقيسها المريض فعلاً عند اختياره مصدر الدواء.")
    cols = ["محور المقارنة", "هاتف + واتساب", "تطبيق صيدلية واحدة", "دليفري عام", "صِلة"]
    subs = ["", "صيدلية الحي", "نمط El Khabiry", "صيدليات Talabat", "شبكة صيدليات موثّقة"]
    rows = [
        ("ثقة صيدلية مرخّصة",      ["نعم", "جزئي", "—", "نعم"]),
        ("توفّر مؤكد لحظي",        ["—", "نعم", "—", "نعم"]),
        ("مقارنة عروض متعددة",      ["—", "—", "جزئي", "نعم"]),
        ("صيدلي متابع للطلب",       ["نعم", "جزئي", "—", "نعم"]),
        ("توصيل قابل للتتبع",       ["—", "جزئي", "نعم", "نعم"]),
        ("شبكة محلية أوسع",        ["—", "—", "جزئي", "نعم"]),
    ]
    cw_ = [3.13, 2.25, 2.25, 2.25, 2.25]
    rh = [0.66] + [0.52] * 6
    tbl = make_table(s, MX, 1.98, CW, sum(rh), 7, 5, cw_, rh)
    # header
    for j in range(5):
        cell = tbl.cell(0, j)
        fill = TEAL if j == 4 else INK
        tcolor = WHITE
        set_cell(cell,
                 [P([R(cols[j], 11, tcolor, True)], align="c"),
                  P([R(subs[j], 8, "CBD5E1" if j == 4 else "94A3B8")], align="c", before=1)],
                 fill=fill, ml=0.08, mr=0.08)
        cell_border_bottom(cell, TEAL, 1.5)
    # body
    for i, (axis, vals) in enumerate(rows, start=1):
        zebra = WHITE if i % 2 else ZEBRA
        c0 = tbl.cell(i, 0)
        set_cell(c0, [P([R(axis, 11, INK, True)])], fill=zebra)
        cell_border_bottom(c0)
        for j, v in enumerate(vals, start=1):
            us = (j == 4)
            fill = TEALSOFT if us else zebra
            if v == "نعم":
                colr, bold = (TEALDK, True)
            elif v == "جزئي":
                colr, bold = AMBERDK, True
            else:
                colr, bold = FAINT, False
            cj = tbl.cell(i, j)
            set_cell(cj, [P([R(v, 11, colr, bold)], align="c")], fill=fill)
            cell_border_bottom(cj, LINE, 0.75)
    # legend + takeaway
    add_text(s, MX, 5.94, CW, 0.28,
             [P([R("نعم = قوي      جزئي = محدود      — = غائب", 9.5, FAINT, True)])])
    add_shape(s, MX, 6.3, CW, 0.56, fill=TEALSOFT, line_color=None, radius=0.14)
    add_text(s, MX + 0.3, 6.3, CW - 0.6, 0.56,
             [P([R("المنافس المباشر (الصيدلية عبر الهاتف) نكسبه شريكاً لا نزاله — ", 11.5, INK, True),
                 R("الميزة الدفاعية: شبكة موثّقة + توفّر مؤكد + صيدلي في الحلقة.", 11.5, TEALDK, True)],
                line=1.15)], anchor="m")
    footer(s, 6, TOTAL)


def s07_compliance(prs):
    s = new_slide(prs)
    header(s, "06", "الثقة", "تحويل الخطر الصحي إلى أصل للعلامة",
           sub="ست ضوابط تشغيلية وتسويقية تُبنى قبل أول حملة — وتُعرض كجزء من القيمة لا كملاحظة قانونية.")
    items = [
        ("المصدر والأصالة", "صيدليات مرخّصة فقط، تصوير وثائق، تدقيق دوري، وبلاغ جودة وتصعيد للجهة المختصة"),
        ("الأدوية المقيدة", "لا صرف بلا وصفة حيث يلزم، مراجعة صيدلي، ومنع استبدال المادة أو التركيز دون موافقة"),
        ("الإعلان الطبي", "لا وعود علاج أو شفاء ولا إعلانات مضللة — محتوى تثقيفي بمراجعة صيدلي وتدقيق قانوني"),
        ("الخصوصية", "أقل قدر من البيانات، تشفير أثناء النقل، موافقة صريحة، وحذف الوصفة وفق سياسة معلنة"),
        ("COD والاحتيال", "تأكيد WhatsApp / OTP، تحديد محاولات، ضوابط للمستخدم عالي الإلغاء، وتسوية يومية"),
        ("سلسلة الحفظ", "عبوة مغلقة، حرارة وضوء حسب المنتج، حقيبة مناسبة، وتسليم للشخص الصحيح"),
    ]
    w, h, gap = 3.91, 1.98, 0.2
    for i, (name, desc) in enumerate(items):
        row, col = divmod(i, 3)
        x = MX + CW - (col + 1) * w - col * gap
        y = 1.95 + row * (h + gap)
        card(s, x, y, w, h)
        add_shape(s, x + w - 0.5, y + 0.24, 0.34, 0.34, fill=TEALSOFT,
                  line_color=None, radius=0.3)
        add_text(s, x + w - 0.5, y + 0.24, 0.34, 0.34,
                 [P([R(str(i + 1), 12, TEALDK, True)], align="c")], anchor="m")
        add_text(s, x + 0.24, y + 0.26, w - 0.9, 0.32,
                 [P([R(name, 13, INK, True)])])
        add_text(s, x + 0.24, y + 0.72, w - 0.48, h - 0.9,
                 [P([R(desc, 10.5, SUB)], line=1.3)])
    add_shape(s, MX, 6.28, CW, 0.58, fill=ROSESOFT, line_color=None, radius=0.14)
    add_shape(s, MX + CW - 0.09, 6.35, 0.045, 0.44, fill=ROSE, line_color=None,
              shape=MSO_SHAPE.RECTANGLE)
    add_text(s, MX + 0.32, 6.28, CW - 0.7, 0.58,
             [P([R("للخطر الطبي: ", 12, ROSEDK, True),
                 R("اتصل بالطوارئ أو توجه لأقرب مستشفى — التطبيق ليس إسعافاً، و«طلب عاجل» يعني أولوية تشغيل فقط.",
                   12, ROSEDK)], line=1.15)], anchor="m")
    footer(s, 7, TOTAL)


def s08_90days(prs):
    s = new_slide(prs)
    header(s, "07", "الـ 90 يوماً", "عنقود واحد في بغداد — ثم توسّع مشروط")
    # proportional phase pills (weeks: 2 / 4 / 4 / 2 of 12)
    weeks = [2, 4, 4, 2]
    total_w = sum(weeks)
    gap = 0.06
    usable = CW - 3 * gap
    x = MX
    fills = ["E6FAF7", "CCFBF1", "99F6E4", "5EEAD4"]
    tcols = [TEALDK, TEALDK, TEALDK, "065F46"]
    names = ["المرحلة 0  •  أسابيع 1–2", "المرحلة 1  •  أسابيع 3–6",
             "المرحلة 2  •  أسابيع 7–10", "المرحلة 3  •  أسابيع 11–13"]
    for i, wks in enumerate(weeks):
        w = usable * wks / total_w
        add_shape(s, x, 1.9, w, 0.44, fill=fills[i], line_color=None, radius=0.5)
        add_text(s, x + 0.05, 1.9, w - 0.1, 0.44,
                 [P([R(names[i], 10, tcols[i], True)], align="c")], anchor="m")
        x += w + gap
    # phase cards
    gap = 0.15
    w = (CW - 3 * gap) / 4
    y, h = 2.58, 3.62
    phases = [
        ("إثبات الجاهزية", ["تحديد عنقود بغداد: الكرادة / الجادرية / المنصور",
                            "تدقيق UX بالعربية العراقية والتهجئات البديلة",
                            "North Star: طلبات مكتملة وآمنة لكل صيدلية نشطة"]),
        ("كثافة العرض", ["15–20 صيدلية مرساة بساعات موثوقة ومخزون متنوع",
                          "تدريب 60 دقيقة + QR على الكاونتر + واتساب تشغيلي",
                          "كباتن ضمن نصف قطر صغير بحقيبة وهوية واضحة"]),
        ("طلب محلي مقاس", ["حملة «ابحث عن دوائك قبل أن تطلع» + QR في العيادات",
                           "Meta جغرافي + Search للنية العالية + TikTok للاعتياد",
                           "إحالة برصيد توصيل وإعادة استهداف بموافقة"]),
        ("تحسين وتوسيع", ["اختبار 3 رسائل وعرضين وأسلوبَي onboarding",
                          "توسيع منطقة واحدة فقط إذا تحققت الكثافة",
                          "Playbook للبصرة / النجف / كربلاء مع شريك محلي"]),
    ]
    for i, (title, lines) in enumerate(phases):
        x = MX + CW - (i + 1) * w - i * gap
        card(s, x, y, w, h)
        add_shape(s, x + w - 1.18, y + 0.2, 0.98, 0.3, fill=TEALSOFT,
                  line_color=None, radius=0.5)
        add_text(s, x + w - 1.18, y + 0.2, 0.98, 0.3,
                 [P([R(f"المرحلة {i}", 9.5, TEALDK, True)], align="c")], anchor="m")
        add_text(s, x + 0.2, y + 0.2, w - 1.45, 0.3,
                 [P([R(title, 13.5, INK, True)])])
        ly = y + 0.78
        for t in lines:
            dot_line(s, x + 0.22, ly, w - 0.44, t, size=10, line_spacing=1.18)
            ly += 0.88
    add_shape(s, MX, 6.42, CW, 0.5, fill=TEALSOFT, line_color=None, radius=0.14)
    add_text(s, MX + 0.32, 6.42, CW - 0.64, 0.5,
             [P([R("قاعدة التوسّع:  ", 11.5, INK, True),
                 R("ثلاثة عروض قابلة للمقارنة على الأقل لكل طلب + قبول 80% على الأقل خلال SLA — وإلا واصل بناء الكثافة.",
                   11.5, TEALDK, True)], line=1.1)], anchor="m")
    footer(s, 8, TOTAL)


def s09_engine(prs):
    s = new_slide(prs)
    header(s, "08", "السوق الثنائي", "لا طلب بلا عرض — ولا عرض بلا طلب",
           sub="آلة بيع ميدانية للصيدليات، وحزمة كباتن محلية، ورسالة تبيع الشراكة لا المنصة.")
    y, colh = 1.92, 3.62
    wcol = 5.96
    xr = MX + CW - wcol
    xl = MX

    # right: pharmacy sales machine
    card(s, xr, y, wcol, colh)
    add_text(s, xr + 0.26, y + 0.2, wcol - 0.5, 0.34,
             [P([R("آلة بيع الصيدليات", 14.5, INK, True)])])
    add_text(s, xr + 0.26, y + 0.56, wcol - 0.5, 0.26,
             [P([R("خطوات ميدانية موحّدة لكل زيارة", 10, FAINT)])])
    steps = [("قائمة", "خرائط + نقابة + إحالات الموزعين"),
             ("تأهيل", "ترخيص وساعات وتغطية وقبول COD وشخص القرار"),
             ("عرض 15 دقيقة", "طلب حي من هاتف المندوب — لا شرائح فقط"),
             ("تفعيل في الموقع", "QR + أعلى 200 SKU أو ربط مخزون + تدريب"),
             ("نجاح", "اتصال بعد 48 ساعة و7 أيام + زيارة أسبوعية")]
    ly = y + 0.98
    for i, (lead, body) in enumerate(steps, 1):
        add_shape(s, xr + wcol - 0.58, ly + 0.02, 0.32, 0.32, fill=TEAL,
                  line_color=None, radius=0.3)
        add_text(s, xr + wcol - 0.58, ly + 0.02, 0.32, 0.32,
                 [P([R(str(i), 11.5, WHITE, True)], align="c")], anchor="m")
        add_text(s, xr + 0.26, ly, wcol - 0.95, 0.36,
                 [P([R(lead + " — ", 11, INK, True), R(body, 10.5, SUB)], line=1.1)],
                 anchor="m")
        ly += 0.515

    # left top: quote
    qt_h = 1.5
    card(s, xl, y, wcol, qt_h)
    add_shape(s, xl + wcol - 0.075, y + 0.24, 0.05, qt_h - 0.48, fill=TEAL,
              line_color=None, shape=MSO_SHAPE.RECTANGLE)
    add_text(s, xl + 0.3, y + 0.22, wcol - 0.7, 0.8,
             [P([R("«لا نملك مخزونك ولا نفرض سعرك. نرسل لك مرضى يبحثون عن دواء متوفر لديك — وأنت تقبل أو ترفض من لوحة واضحة.»",
                   12, INK, True)], line=1.3)])
    add_text(s, xl + 0.3, y + qt_h - 0.42, wcol - 0.7, 0.28,
             [P([R("الرسالة البيعية لصيدليتك", 9.5, FAINT, True)])])

    # left bottom: partner package
    pk_y = y + qt_h + 0.14
    pk_h = colh - qt_h - 0.14
    card(s, xl, pk_y, wcol, pk_h)
    add_text(s, xl + 0.26, pk_y + 0.18, wcol - 0.5, 0.3,
             [P([R("باقة الشريك", 13, INK, True)])])
    perks = ["تسجيل ومراجعة مجانيان — عمولة على الطلب المكتمل فقط",
             "صفحة موثّقة وتقييمات بعد التسليم",
             "تحكم كامل بالسعر والعروض وسجل تسويات",
             "حوافز مرتبطة بالجودة — لا مقابل إدراج دواء"]
    ly = pk_y + 0.56
    for t in perks:
        dot_line(s, xl + 0.28, ly, wcol - 0.56, t, size=10.5, before=0)
        ly += 0.4

    # bottom: captain strip
    by = 5.72
    card(s, MX, by, CW, 0.78)
    add_text(s, MX + CW - 2.4, by, 2.1, 0.78,
             [P([R("محرك الكباتن", 13, INK, True)])], anchor="m")
    caps = ["تجنيد ضمن نصف قطر صغير", "حقيبة ولباس وهوية موثوقة",
            "سجل أرباح وخريطة واضحة", "تسوية يومية دقيقة", "قياس الدخل لكل ساعة"]
    seg_w = (CW - 2.6) / 5
    for i, t in enumerate(caps):
        x = MX + 0.25 + i * seg_w
        if i:
            add_shape(s, x - 0.03, by + 0.19, 0.012, 0.4, fill=LINE,
                      line_color=None, shape=MSO_SHAPE.RECTANGLE)
        add_text(s, x + 0.12, by, seg_w - 0.2, 0.78,
                 [P([R(t, 10, SUB, True)], line=1.15)], anchor="m")
    add_text(s, MX, 6.62, CW, 0.3,
             [P([R("اعتراضات شائعة:  ", 10, FAINT, True),
                 R("«عندي واتساب» → يضاف له لا يُلغى   •   «العمولة مرتفعة» → اختبر 30 يوماً واقس الصافي   •   «أخشى الوصفات» → الصيدلي يقرر مع سجل تصعيد",
                   10, SUB)], line=1.2)])
    footer(s, 9, TOTAL)


def s10_demand(prs):
    s = new_slide(prs)
    header(s, "09", "الطلب B2C", "خمس حملات — وخمس قنوات واضحة",
           sub="كل قناة لها وظيفة واحدة: جلب العائلات، بناء العادة، أو تأمين الخدمة — لا قناة تعمل بلا هدف.")
    chans = [("فيسبوك", "العائلة والأحياء + Click-to-WhatsApp"),
             ("إنستغرام", "Reels وقصص الصيدلي + إعادة استهداف"),
             ("تيك توك", "مواقف إنسانية 20–30 ث — بلا قبل/بعد طبي"),
             ("واتساب Business", "تأكيد وتحديثات وإعادة طلب بموافقة"),
             ("ميداني ومجتمعي", "QR في صيدليات وعيادات ومختبرات وجامعات"),
             ("ASO / SEO", "«توصيل دواء بغداد» • «رفع وصفة»")]
    w, h, gap = 3.91, 0.58, 0.16
    for i, (name, desc) in enumerate(chans):
        row, col = divmod(i, 3)
        x = MX + CW - (col + 1) * w - col * gap
        y = 1.9 + row * (h + gap)
        card(s, x, y, w, h, radius=0.5)
        add_text(s, x + w - 2.05, y, 1.9, h,
                 [P([R(name, 11, TEALDK, True)])], anchor="m")
        add_text(s, x + 0.18, y, 1.5, h,
                 [P([R(desc, 9.5, SUB)], line=1.1)], anchor="m")
    # campaigns
    camps = [("«لا تلف بغداد»", "مشهد عائلي: بحث، رفع وصفة، عروض، وصول", "صوّر وصفتك وخلي صِلة تدور لك"),
             ("«متوفر فعلاً»", "الصيدلية والمسافة والسعر قبل الخروج", "اعرف قبل لا تطلع"),
             ("«صِلة لأهلي»", "ابن يطلب لوالديه بعنوان محفوظ وتحديثات", "خلّي دواء أهلك أقرب"),
             ("«صيدليتك أقرب»", "B2B: الصيدلي يقبل طلباً حقيقياً من لوحته", "سجّل مجاناً — والاعتماد أولاً"),
             ("«الكابتن الموثوق»", "هوية وحقيبة وتوقيع تسليم — بلا تصوير دون موافقة", "الثقة توصل")]
    gap = 0.2
    w = (CW - 4 * gap) / 5
    y, h = 3.34, 2.42
    for i, (name, desc, cta) in enumerate(camps):
        x = MX + CW - w - i * (w + gap)
        card(s, x, y, w, h)
        add_shape(s, x + w / 2 - 0.17, y + 0.2, 0.34, 0.34, fill=TEAL,
                  line_color=None, radius=0.5)
        add_text(s, x + w / 2 - 0.17, y + 0.2, 0.34, 0.34,
                 [P([R(str(i + 1), 12, WHITE, True)], align="c")], anchor="m")
        add_text(s, x + 0.14, y + 0.68, w - 0.28, 0.34,
                 [P([R(name, 11.5, INK, True)], align="c")])
        add_text(s, x + 0.16, y + 1.06, w - 0.32, 0.72,
                 [P([R(desc, 9, SUB)], align="c", line=1.25)])
        add_shape(s, x + 0.16, y + h - 0.56, w - 0.32, 0.36, fill=TEALSOFT,
                  line_color=None, radius=0.5)
        add_text(s, x + 0.16, y + h - 0.56, w - 0.32, 0.36,
                 [P([R(cta, 8.5, TEALDK, True)], align="c", line=1.05)], anchor="m")
    add_shape(s, MX, 6.0, CW, 0.56, fill=AMBERSOFT, line_color=None, radius=0.14)
    add_shape(s, MX + CW - 0.09, 6.07, 0.045, 0.42, fill=AMBER, line_color=None,
              shape=MSO_SHAPE.RECTANGLE)
    add_text(s, MX + 0.32, 6.0, CW - 0.7, 0.56,
             [P([R("لهجة عراقية منضبطة — لا نقول أبداً:  ", 11, AMBERDK, True),
                 R("«يشفي»  •  «مضمون 100%»  •  «أسرع إسعاف»  •  «أرخص دواء»  •  تشجيع علاج بلا وصفة",
                   11, AMBERDK)], line=1.1)], anchor="m")
    footer(s, 10, TOTAL)


def s11_kpis(prs):
    s = new_slide(prs)
    header(s, "10", "الأداء", "أهداف الـ 90 يوماً — تُقاس، لا تُوعد")
    add_text(s, MX, 1.42, CW, 0.32,
             [P([R("North Star:  ", 12, TEALDK, True),
                 R("طلبات دوائية مكتملة وآمنة لكل صيدلية نشطة — لا تنزيلات.", 12, INK, True)])])
    kpis = [("60", "صيدلية معتمدة", "في عنقود الإطلاق — منها 40 نشطة أسبوعياً"),
            ("3,000", "طلب مكتمل", "خلال أول 90 يوماً"),
            ("85%", "حد أدنى لملء الطلب", "ضمن ساعات الخدمة"),
            ("5 دقائق", "حد أقصى لوسطي قبول الصيدلية", "يُقاس لا يُوعد به"),
            ("60 دقيقة", "حد أقصى للتوصيل الحضري", "للطلبات الجاهزة"),
            ("25%", "حد أدنى لإعادة الشراء", "بعد 30 يوماً — غير الطارئ"),
            ("12%", "حد أقصى لإلغاء COD", "مع تسوية يومية للكابتن"),
            ("0", "شكاوى دواء حرجة", "تطابق وسلامة — صفرية")]
    gap = 0.15
    w = (CW - 3 * gap) / 4
    h = 1.78
    ys = [2.14, 4.1]
    for i, (num, label, sub_) in enumerate(kpis):
        row, col = divmod(i, 4)
        x = MX + CW - w - col * (w + gap)
        y = ys[row]
        card(s, x, y, w, h)
        add_shape(s, x + w - 0.5, y + 0.22, 0.34, 0.34, fill=TEALSOFT,
                  line_color=None, radius=0.5)
        add_text(s, x + w - 0.5, y + 0.22, 0.34, 0.34,
                 [P([R(f"{i+1:02d}", 9, TEALDK, True)], align="c")], anchor="m")
        add_text(s, x + 0.22, y + 0.2, w - 0.85, 0.62,
                 [P([R(num, 30, TEAL, True)])])
        add_text(s, x + 0.22, y + 0.92, w - 0.44, 0.3,
                 [P([R(label, 11, INK, True)])])
        add_text(s, x + 0.22, y + 1.24, w - 0.44, 0.4,
                 [P([R(sub_, 9, FAINT)], line=1.15)])
    add_shape(s, MX, 6.12, CW, 0.56, fill="F1F5F9", line_color=LINE, line_w=0.75, radius=0.14)
    add_text(s, MX + 0.32, 6.12, CW - 0.64, 0.56,
             [P([R("قاعدة التوقف:  ", 11, INK, True),
                 R("لا توسّع جغرافياً إذا كانت المساهمة لكل طلب سالبة بعد الخصم، أو انخفضت السلامة والتسليم عن SLA.",
                   11, SUB, True)], line=1.15)], anchor="m")
    footer(s, 11, TOTAL)


def s12_decisions(prs):
    s = new_slide(prs)
    header(s, "11", "القرار", "خمسة قرارات تبدأ هذا الأسبوع",
           sub="العرض لا ينتهي بشعار — بل بطلبات قرار واضحة يمكن اتخاذها في نفس الاجتماع.")
    decisions = [
        ("اعتماد عنقود الإطلاق", "الكرادة / الجادرية / المنصور — حتى يصبح الوعد «متاحاً فعلاً»"),
        ("ميزانية اختبار Pilot", "14 يوماً للقياس لا للتوسع: time-to-first-offer و contribution/order قبل أي إنفاق إضافي"),
        ("مسؤول امتثال وسلامة", "صيدلي مرخّص بصلاحيات على الوصفات والفئات المحظورة والتصعيد وسجل التدقيق"),
        ("20 صيدلية مرساة", "بساعات موثوقة وكثافة مخزون — ثم تجلب الطلب حولها"),
        ("تاريخ Pilot محدد", "مع اجتماع أسبوعي: Growth / Operations / Safety على نفس لوحة المؤشرات"),
    ]
    y = 1.95
    for i, (title, desc) in enumerate(decisions, 1):
        h = 0.72
        card(s, MX, y, CW, h)
        add_shape(s, MX + CW - 0.62, y + h / 2 - 0.21, 0.42, 0.42, fill=TEAL,
                  line_color=None, shape=MSO_SHAPE.OVAL, radius=None)
        add_text(s, MX + CW - 0.62, y + h / 2 - 0.21, 0.42, 0.42,
                 [P([R(str(i), 14, WHITE, True)], align="c")], anchor="m")
        add_text(s, MX + 0.3, y, 4.3, h,
                 [P([R(title, 13.5, INK, True)])], anchor="m")
        add_shape(s, MX + 4.75, y + 0.17, 0.012, h - 0.34, fill=LINE,
                  line_color=None, shape=MSO_SHAPE.RECTANGLE)
        add_text(s, MX + 5.0, y, CW - 5.75, h,
                 [P([R(desc, 10.5, SUB)], line=1.2)], anchor="m")
        y += h + 0.13
    add_shape(s, MX, 6.42, CW, 0.5, fill=TEALSOFT, line_color=None, radius=0.14)
    add_text(s, MX + 0.32, 6.42, CW - 0.64, 0.5,
             [P([R("الخطوة الأولى:  ", 11.5, INK, True),
                 R("45 مقابلة ميدانية + صفحة هبوط للصيدليات بمؤشر «اعتماد صيدلية» قابل للإثبات + تحليلات موحّدة",
                   11.5, TEALDK, True)], line=1.1)], anchor="m")
    footer(s, 12, TOTAL)


def s13_sevendays(prs):
    s = new_slide(prs)
    header(s, "12", "التنفيذ", "قائمة تنفيذ فورية — خلال 7 أيام",
           sub="تسلسل مقصود: ما يُثبَت قبل أن يُعلن، وما يُقاس قبل أن يُوسَّع.")
    rows = [
        ("تأكيد نطاق العمل", "المحافظات والمناطق العاملة فعلياً، ساعات كل صيدلية، ومتجر التوزيع (App Store / Google Play)"),
        ("45 مقابلة", "10 مرضى، 10 مقدمي رعاية، 15 صيدلية، 10 كباتن — تسجيل الاقتباسات والعوائق لا الآراء العامة"),
        ("اختبار البحث", "20 اسماً دوائياً + 10 صور وصفات + اختلافات الإملاء واللهجة"),
        ("سياسات الامتثال", "الأدوية المقيدة، الرفض، الإرجاع، الخصوصية، الحوادث، وسلسلة الحرارة مع مستشار محلي"),
        ("صفحة هبوط", "صفحة / واتساب للصيدليات مع عرض تجريبي ومؤشر «اعتماد صيدلية» قابل للإثبات"),
        ("التحليلات الموحّدة", "Firebase / analytics مع أحداث القمع كاملة و UTM موحّدة قبل أول حملة"),
        ("Pilot مدفوع", "14 يوماً صغيرة ومقارنة ثلاث رسائل — لا ميزانية توسيع قبل قراءة القياسات"),
    ]
    cw_ = [0.85, 3.6, 7.68]
    rh = [0.5] + [0.545] * 7
    tbl = make_table(s, MX, 1.98, CW, sum(rh), 8, 3, cw_, rh)
    heads = ["الرقم", "عنصر التنفيذ", "التفاصيل"]
    for j, htxt in enumerate(heads):
        cell = tbl.cell(0, j)
        set_cell(cell, [P([R(htxt, 11, WHITE, True)],
                          align="c" if j == 0 else "r")], fill=INK,
                 ml=0.1, mr=0.16)
        cell_border_bottom(cell, TEAL, 1.5)
    for i, (item, detail) in enumerate(rows, start=1):
        zebra = WHITE if i % 2 else ZEBRA
        c0 = tbl.cell(i, 0)
        set_cell(c0, [P([R(str(i), 12, TEALDK, True)], align="c")], fill=zebra,
                 ml=0.05, mr=0.05)
        cell_border_bottom(c0)
        c1 = tbl.cell(i, 1)
        set_cell(c1, [P([R(item, 11, INK, True)])], fill=zebra, mr=0.16)
        cell_border_bottom(c1)
        c2 = tbl.cell(i, 2)
        set_cell(c2, [P([R(detail, 10, SUB)], line=1.12)], fill=zebra, mr=0.16)
        cell_border_bottom(c2)
    add_text(s, MX, 6.62, CW, 0.3,
             [P([R("المبدأ الحاكم:  ", 10.5, TEALDK, True),
                 R("لا صرف ميزانية توسيع قبل قراءة time-to-first-offer و contribution/order من بيانات الـ pilot.",
                   10.5, INK, True)])])
    footer(s, 13, TOTAL)


def s14_closing(prs):
    s = new_slide(prs, bg=NAVY)
    ring = add_shape(s, 10.9, 0.9, 1.7, 1.7, fill=None, line_color=NAVYLN,
                     line_w=1.2, shape=MSO_SHAPE.OVAL, radius=None)
    add_shape(s, 1.7, 5.9, 0.14, 0.14, fill=TEAL, line_color=None,
              shape=MSO_SHAPE.OVAL, radius=None)
    # logo
    add_shape(s, 11.62, 0.55, 0.55, 0.55, fill=TEAL, line_color=None, radius=0.30)
    add_text(s, 11.62, 0.55, 0.55, 0.55,
             [P([R("+", 22, WHITE, True)], align="c")], anchor="m")
    add_text(s, 9.62, 0.52, 1.92, 0.34, [P([R("صِلة دوائي", 15, "F8FAFC", True)])])
    add_text(s, 9.62, 0.86, 1.92, 0.24, [P([R("S I L A  I Q", 9, FAINT, True)])])

    add_text(s, 6.73, 2.1, 6.0, 0.34,
             [P([R("الخطوة التالية", 11, TEALLT, True)], align="c")])
    add_shape(s, 12.13, 2.62, 0.55, 0.045, fill=TEAL, line_color=None,
              shape=MSO_SHAPE.RECTANGLE)
    add_text(s, 0.9, 2.9, 11.83, 1.6,
             [P([R("هذا ليس شعاراً —", 38, "F8FAFC", True)], line=1.15),
              P([R("إنه قرار إطلاق.", 38, TEALLT, True)], line=1.15)])
    add_text(s, 0.9, 4.72, 11.83, 0.4,
             [P([R("دواؤك الموثوق، من صيدلية قريبة، وبالسعر والوقت الواضحين.",
                  14.5, "94A3B8")])])
    add_shape(s, 7.43, 5.5, 5.3, 0.42, fill=NAVY2, line_color=NAVYLN, line_w=1.0, radius=0.5)
    add_text(s, 7.43, 5.5, 5.3, 0.42,
             [P([R("صِلة دوائي  •  SILA IQ  —  أيلول 2026", 10.5, "B6C2D4", True)],
                align="c")], anchor="m")
    footer(s, 14, TOTAL, navy=True)


# ================================================================ main
def main():
    prs = Presentation()
    prs.slide_width = Emu(W)
    prs.slide_height = Emu(H)
    cp = prs.core_properties
    cp.title = "صِلة دوائي SILA IQ — الخطة التسويقية والاستراتيجية للنمو"
    cp.author = "SILA IQ"
    cp.subject = "Executive marketing strategy deck — 14 slides"
    cp.comments = "Generated with Open Design pipeline — Cairo font, RTL, Google Slides ready."

    builders = [s01_cover, s02_problem, s03_system, s04_positioning,
                s05_segments, s06_competition, s07_compliance, s08_90days,
                s09_engine, s10_demand, s11_kpis, s12_decisions,
                s13_sevendays, s14_closing]
    for b in builders:
        b(prs)

    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "docs", "silla-iq-executive-presentation.pptx")
    prs.save(out)
    print(f"OK  {out}")
    print(f"slides: {len(prs.slides.__iter__.__self__._sldIdLst)}")


if __name__ == "__main__":
    main()
