#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build "The Business Show" pitch deck as an editable PowerPoint (.pptx).

Design mirrors index.html: warm-white minimalist system, CNNBA red accent,
RTL Arabic type, cards, tables, and Morph/Fade slide transitions (motion).

Images: the script tries to download each concept visual and embed it.
- Run locally (where the CDN is reachable) -> real photos are embedded.
- Run in a network-restricted sandbox -> tasteful branded placeholders are
  drawn instead, and the image URL is written to the slide's speaker notes.

Usage:  pip install python-pptx pillow  &&  python build_pptx.py
Output: The_Business_Show.pptx
"""
import os, io, urllib.request
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from copy import deepcopy
from lxml import etree

# ---------------------------------------------------------------- palette
PAPER   = RGBColor(0xF6, 0xF2, 0xEC)
PAPER2  = RGBColor(0xEF, 0xE9, 0xDF)
INK     = RGBColor(0x1B, 0x1A, 0x18)
INKSOFT = RGBColor(0x4A, 0x47, 0x3F)
LINE    = RGBColor(0xDA, 0xD2, 0xC4)
WOOD    = RGBColor(0xC7, 0xA5, 0x6B)
STONE   = RGBColor(0xE5, 0xDD, 0xCE)
RED     = RGBColor(0xB0, 0x1E, 0x2E)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
GREEN   = RGBColor(0x00, 0xA6, 0x7A)

HEAD = "Cairo"       # heading font (falls back gracefully if not installed)
BODY = "Tajawal"     # body font

CDN = "https://d8j0ntlcm91z4.cloudfront.net/user_3DIduyBmzwYgrJ6fxIBny2g7uu3/"
IMG = {
    "hero":     CDN + "hf_20260701_060546_696d30aa-45c3-4fb7-b212-b9c3cd81c514.png",
    "portrait": CDN + "hf_20260701_060550_7d48016f-3dad-4a2e-b323-1cc867552a83.png",
    "twoshot":  CDN + "hf_20260701_060553_15484b5a-9c34-485c-95ce-5043676707c5.png",
    "vertical": CDN + "hf_20260701_060555_fde81cf9-2958-4cc7-99d7-020198cccf29.png",
    "opposing": CDN + "hf_20260701_053411_86db6605-555f-492b-b349-f04ed09707a3.png",
    "lower":    CDN + "hf_20260701_053426_bd2f7b1b-2a22-4f31-ac4c-92d3822e90d7.png",
    "palette":  CDN + "hf_20260701_053429_9828122b-3bf2-400f-96af-189ded280e78.png",
}
_imgcache = {}
def fetch(key):
    if key in _imgcache:
        return _imgcache[key]
    try:
        req = urllib.request.Request(IMG[key], headers={"User-Agent": "Mozilla/5.0"})
        data = urllib.request.urlopen(req, timeout=20).read()
        _imgcache[key] = io.BytesIO(data)
    except Exception:
        _imgcache[key] = None
    return _imgcache[key]

# ---------------------------------------------------------------- deck
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
EMU = 914400
SW, SH = 13.333, 7.5

def slide(bg=PAPER):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = bg; r.line.fill.background()
    r.shadow.inherit = False
    # send bg to back
    sp = r._element; sp.getparent().remove(sp); s.shapes._spTree.insert(2, sp)
    return s

def _set_cs(run, font):
    """Ensure Arabic (complex-script) + latin typeface are both set."""
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", font)

def _rtl(p, align=PP_ALIGN.RIGHT):
    p.alignment = align
    pPr = p._p.get_or_add_pPr(); pPr.set("rtl", "1")

def txt(s, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(2); tf.margin_top = tf.margin_bottom = Pt(2)
    return tb, tf

def para(tf, runs, size=18, color=INK, bold=False, font=BODY,
         align=PP_ALIGN.RIGHT, rtl=True, space_after=6, space_before=0, line=1.15, first=False):
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    if rtl: _rtl(p, align)
    else:   p.alignment = align
    p.space_after = Pt(space_after); p.space_before = Pt(space_before); p.line_spacing = line
    if isinstance(runs, str): runs = [(runs, color, bold)]
    for item in runs:
        t_, c_, b_ = (item + (color, bold))[:3] if isinstance(item, tuple) else (item, color, bold)
        r = p.add_run(); r.text = t_
        r.font.size = Pt(size); r.font.bold = b_; r.font.name = font
        r.font.color.rgb = c_; _set_cs(r, font)
    return p

def rect(s, l, t, w, h, fill=None, lc=None, lw=1.0, rounded=False, dash=None):
    shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                             Inches(l), Inches(t), Inches(w), Inches(h))
    shp.shadow.inherit = False
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if lc is None: shp.line.fill.background()
    else:
        shp.line.color.rgb = lc; shp.line.width = Pt(lw)
        if dash: shp.line._get_or_add_ln().append(etree.SubElement(shp.line._get_or_add_ln(), qn('a:prstDash'), {'val': dash}))
    if rounded:
        try: shp.adjustments[0] = 0.06
        except Exception: pass
    return shp

def hline(s, l, t, w, color=LINE, weight=1.0):
    ln = s.shapes.add_connector(2, Inches(l), Inches(t), Inches(l + w), Inches(t))
    ln.line.color.rgb = color; ln.line.width = Pt(weight); return ln

def kicker(s, text, t=0.55):
    rect(s, SW-0.62-0.0, t+0.09, 0.42, 0.028, fill=RED)  # short rule
    _, tf = txt(s, SW-6.5-0.75, t-0.06, 6.5, 0.4)
    para(tf, text, size=12.5, color=RED, bold=True, font=HEAD, first=True, space_after=0)

def corner_logo(s, dark=True):
    """Simplified, editable CNN Business mark: 'CNN business' wordmark + red box."""
    x, y, hh = 0.75, 0.42, 0.30
    # red box with white 'business'
    bw = 1.15
    box = rect(s, x, y, bw, hh, fill=RED, rounded=False)
    _, tf = txt(s, x, y-0.045, bw, hh+0.09, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, "business", size=13, color=WHITE, bold=True, font=HEAD,
         align=PP_ALIGN.CENTER, rtl=False, first=True, space_after=0)
    # CNN wordmark to the left of the box
    _, tf2 = txt(s, x-1.35, y-0.045, 1.3, hh+0.09, anchor=MSO_ANCHOR.MIDDLE)
    para(tf2, "CNN", size=17, color=(WHITE if not dark else INK), bold=True, font=HEAD,
         align=PP_ALIGN.RIGHT, rtl=False, first=True, space_after=0)
    # green corner tick
    tri = s.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(x-1.5), Inches(y+hh-0.10), Inches(0.14), Inches(0.10))
    tri.rotation = 180; tri.fill.solid(); tri.fill.fore_color.rgb = GREEN; tri.line.fill.background(); tri.shadow.inherit = False

def snum(s, n):
    _, tf = txt(s, 0.6, 0.5, 1.2, 0.3)
    para(tf, f"{n:02d}", size=11, color=INKSOFT, bold=True, font=HEAD,
         align=PP_ALIGN.LEFT, rtl=False, first=True, space_after=0)

def notes(s, text):
    s.notes_slide.notes_text_frame.text = text

def picture(s, key, l, t, w, h, caption=None):
    """Embed the real image if reachable, else a branded placeholder frame."""
    bio = fetch(key)
    if bio is not None:
        bio.seek(0)
        try:
            s.shapes.add_picture(bio, Inches(l), Inches(t), Inches(w), Inches(h))
            if caption:
                cap = rect(s, l, t+h-0.42, w, 0.42, fill=INK); cap.fill.fore_color.rgb = INK
                _, tf = txt(s, l+0.15, t+h-0.42, w-0.3, 0.42, anchor=MSO_ANCHOR.MIDDLE)
                para(tf, caption, size=11, color=WHITE, bold=True, font=HEAD, first=True, space_after=0)
            return
        except Exception:
            pass
    # placeholder
    ph = rect(s, l, t, w, h, fill=PAPER2, lc=LINE, lw=1.2, rounded=True)
    _, tf = txt(s, l+0.2, t, w-0.4, h, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, "◱  صورة", size=15, color=INKSOFT, bold=True, font=HEAD,
         align=PP_ALIGN.CENTER, first=True, space_after=2)
    if caption:
        para(tf, caption, size=11.5, color=INKSOFT, font=BODY, align=PP_ALIGN.CENTER, space_after=0)
    notes(s, (s.notes_slide.notes_text_frame.text + "\n" if s.has_notes_slide else "") +
             f"[صورة] {caption or key}: {IMG[key]}")

# ---------- motion: Morph transition (fallback Fade) on every slide ----------
TRANS = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
         '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">'
         '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med" p14:dur="700">'
         '<p14:morph option="byObject"/></p:transition></mc:Choice>'
         '<mc:Fallback>'
         '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med">'
         '<p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
def add_transition(s):
    sld = s._element
    node = etree.fromstring(TRANS)
    cmo = sld.find(qn('p:clrMapOvr'))
    if cmo is not None: cmo.addnext(node)
    else: sld.find(qn('p:cSld')).addnext(node)

# ---------- entrance animation: fade-in of a shape, auto after load ----------
def fade_in(s, shape, delay=200, dur=500):
    spid = shape.shape_id
    xml = f'''<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
      <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
      <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
      <p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="afterEffect">
      <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>
      <p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
      <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>
      <p:to><p:strVal val="visible"/></p:to></p:set>
      <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="{dur}"/>
      <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>
      </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn>
      <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
      <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
      </p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>
      <p:bldLst><p:bldP spid="{spid}" grpId="0"/></p:bldLst></p:timing>'''
    node = etree.fromstring(xml)
    sld = s._element
    old = sld.find(qn('p:timing'))
    if old is not None: sld.remove(old)
    sld.append(node)

def card(s, l, t, w, h, num, title, body, fill=WHITE, lc=LINE):
    rect(s, l, t, w, h, fill=fill, lc=lc, lw=1.0, rounded=True)
    _, tf = txt(s, l+0.22, t+0.16, w-0.44, h-0.3)
    if num: para(tf, num, size=13, color=RED, bold=True, font=HEAD, first=True, space_after=2)
    para(tf, title, size=15.5, color=INK, bold=True, font=HEAD, first=(num is None), space_after=4, line=1.05)
    para(tf, body, size=11.5, color=INKSOFT, font=BODY, space_after=0, line=1.15)

def title_block(s, lines, l, t, w, size=34, gap=0.0):
    _, tf = txt(s, l, t, w, 2.2)
    for i, ln in enumerate(lines):
        para(tf, ln, size=size, color=INK, bold=True, font=HEAD, first=(i == 0),
             space_after=2, line=1.02)
    return tf

# ================================================================ SLIDES
def std(n, kick):
    s = slide(); corner_logo(s, dark=True); snum(s, n); kicker(s, kick); return s

# --- 01 COVER ---
s = slide(INK)
if fetch("hero"):
    fetch("hero").seek(0); s.shapes.add_picture(fetch("hero"), 0, 0, prs.slide_width, prs.slide_height)
ov = rect(s, 0, 0, SW, SH, fill=INK); ov.fill.fore_color.rgb = INK
ov.fill.transparency = 0  # solid; approximate scrim
try:
    # set ~55% transparency on overlay
    sp = ov.fill.fore_color._xFill
except Exception:
    pass
# simpler scrim: a semi rectangle via alpha
ov._element.spPr.find(qn('a:solidFill')).find(qn('a:srgbClr')).append(
    etree.fromstring('<a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="55000"/>'))
_, tf = txt(s, 1.2, 2.0, SW-2.4, 3.6, anchor=MSO_ANCHOR.MIDDLE)
para(tf, "CNN BUSINESS ARABIC", size=14, color=WHITE, bold=True, font=HEAD, align=PP_ALIGN.CENTER, rtl=False, first=True, space_after=8)
para(tf, "The Business Show", size=58, color=WHITE, bold=True, font=HEAD, align=PP_ALIGN.CENTER, rtl=False, space_after=4)
para(tf, "البرنامج الاقتصادي الأسبوعي الرئيسي", size=22, color=WOOD, bold=True, font=HEAD, align=PP_ALIGN.CENTER, space_after=14)
para(tf, "ليس مجرد برنامج حواري آخر… بل منصّة تُطرح فيها أهم القرارات الاقتصادية في المنطقة، وتُناقَش، وتُختبر، ثم تُترجم إلى لغة يفهمها كل مشاهد.",
     size=15, color=WHITE, font=BODY, align=PP_ALIGN.CENTER, space_after=16, line=1.35)
para(tf, "يتضمّن فقرة «غرفة القرار»  ·  الحلقة التجريبية: القطاع العقاري",
     size=13.5, color=WOOD, bold=True, font=HEAD, align=PP_ALIGN.CENTER, space_after=0)
notes(s, "الغلاف — The Business Show. صورة الخلفية: " + IMG["hero"])

# --- 02 THE IDEA ---
s = std(2, "الفكرة")
title_block(s, ["شخصٌ واحد.", "قرارٌ واحد.", "حوارٌ لا يتنازل."], SW-6.2-0.75, 2.2, 6.2, size=40)
_, tf = txt(s, 0.75, 2.35, 5.7, 3.6)
para(tf, "شخصٌ يملك رأس مال حقيقيًّا أو صلاحية قرار فعلية، في حوار عربي معدٍّ وغير متسرّع، يُضغط فيه على الخيارات التي يتخذها فعلًا بشأن المال — لا على نقاط حديثه الإعلامية المُعلَّبة.",
     size=15.5, color=INKSOFT, font=BODY, first=True, space_after=12, line=1.4)
hline(s, 0.75, 4.5, 5.7)
para(tf, [("الوعد للمشاهد:  ", INK, True), ("هذه هي الغرفة التي يُفسِّر فيها أصحاب القرار خياراتهم، ويطرح عليهم الأسئلة الصعبة مَن يملك من الثقل ما يُخوِّله ذلك.", INKSOFT, False)],
     size=13, font=BODY, space_before=8, space_after=0, line=1.35)

# --- 03 WHY THE HOST ---
s = std(3, "لماذا المقدِّمة")
picture(s, "portrait", 0.75, 1.55, 5.7, 5.2, caption="المقدِّمة — في الاستوديو")
title_block(s, ["الميزة الحقيقية", "ليست الطاولة…", "بل مَن يجلس خلفها."], SW-6.0-0.75, 2.0, 6.0, size=32)
_, tf = txt(s, SW-6.0-0.75, 4.2, 6.0, 2.6)
para(tf, [("خلفيتها المتعددة التخصصات هي الميزة: ليست المتشددة الاقتصادية ولا المدافعة عن فئة بعينها. لذلك يثق بها الضيوف من مختلف القطاعات في أنها ستكون منصفة، ويراها الجمهور ", INKSOFT, False),
          ("صوتًا مؤسسيًّا لا صوتًا فئويًّا.", INK, True)], size=14.5, font=BODY, first=True, line=1.4, space_after=0)

# --- 04 ANATOMY ---
s = std(4, "البنية والتشريح")
title_block(s, ["أربع حركات متأنية"], SW-9.0-0.75, 1.15, 9.0, size=32)
_, tf = txt(s, 0.75, 1.65, 6.0, 0.4)
para(tf, "حلقة ~24 دقيقة، بأسلوب رقمي أولًا", size=13, color=INKSOFT, font=BODY, first=True, space_after=0)
moves = [("90 ثانية","الموجز","تُحدِّد المقدِّمة القرار الواحد الذي تتمحور حوله الحلقة — لا السيرة الذاتية، بل الخيار."),
         ("16–18 دقيقة","غرفة القرار","الفقرة الجوهرية. ضيف واحد، عمق حقيقي، رحلة عبر ثلاثة أو أربعة قرارات فعلية."),
         ("3–4 دقائق","الكرسي المقابل","يُواجَه الضيف بأقوى رأي معارض، مصدره شخص حقيقي."),
         ("90 ثانية","الخلاصة","تتجه المقدِّمة نحو الكاميرا وتُترجم الحوار إلى ما يعنيه لمال المشاهد.")]
cw, gap, x0, y0, ch = 2.86, 0.12, 0.75, 2.45, 3.3
hline(s, x0, y0-0.12, cw*4+gap*3, color=INK, weight=2)
for i,(tm,h4,body) in enumerate(moves):
    x = x0 + i*(cw+gap)
    _, tf = txt(s, x, y0+0.1, cw, ch)
    para(tf, tm.split()[0], size=30, color=RED, bold=True, font=HEAD, align=PP_ALIGN.RIGHT, first=True, space_after=0)
    para(tf, " ".join(tm.split()[1:]), size=12, color=INKSOFT, font=BODY, space_after=8)
    para(tf, h4, size=17, color=INK, bold=True, font=HEAD, space_after=5)
    para(tf, body, size=11.5, color=INKSOFT, font=BODY, line=1.2, space_after=0)
    if i: hline(s, x-gap/2, y0+0.1, 0.001)  # subtle divider skipped
_, tf = txt(s, 0.75, 6.35, SW-1.5, 0.5)
para(tf, [("القاعدة الذهبية:  ", INK, True), ("لا ينتهي البرنامج أبدًا بالضيف — بل ينتهي بك أنت.", RED, True)],
     size=14, font=HEAD, first=True, space_after=0)

# --- 05 OPPOSING CHAIR ---
s = std(5, "الميكانيكية المميّزة")
picture(s, "opposing", SW-5.3-0.75, 1.7, 5.3, 4.9, caption="الكرسي المقابل — لحظة متأصّلة في البرنامج")
title_block(s, ["الكرسي المقابل"], 0.75, 1.5, 6.2, size=38)
_, tf = txt(s, 0.75, 2.5, 6.2, 1.8)
para(tf, "قبل كل تسجيل، يجمع الفريق تحديًّا حادًّا موجَّهًا للضيف من طرف مقابل حقيقي — منافس، محلل متشكك، مؤسِّس احترق، أو مواطن تأثر بالقرار. في منتصف المقابلة تطرحه المقدِّمة مباشرةً وتُلزم الضيف بالرد أمام الكاميرا.",
     size=13.5, color=INKSOFT, font=BODY, first=True, line=1.35, space_after=0)
cy = 4.35
for i,(n,h4,b) in enumerate([("01","احتكاك حقيقي","اللحظة التي تجعل المقاطع تنتشر."),
                             ("02","استقلالية تحريرية","مؤشر يثبت أن CNNBA غير مُستتبَعة."),
                             ("03","إيقاع مميّز","«ماذا قالوا هذا الأسبوع؟» سببٌ للمتابعة.")]):
    card(s, 0.75+i*2.06, cy, 1.94, 2.05, n, h4, b, fill=PAPER2, lc=STONE)

# --- 06 MAYA MOMENT ---
s = std(6, "The Maya Moment · ماذا يعني ذلك بالنسبة لك؟")
q = rect(s, 0.75, 1.5, SW-1.5, 3.3, fill=INK, rounded=True)
_, tf = txt(s, 1.2, 1.85, SW-2.4, 2.7, anchor=MSO_ANCHOR.MIDDLE)
para(tf, "«في الاقتصاد، ليست كل زيادة في الأسعار خبراً جيداً، وليست كل موجة صعود فقاعة. التحدي الحقيقي هو التوازن بين حماية قيمة الأصول، والحفاظ على القدرة على السكن والاستثمار. فهذا التوازن هو ما يصنع سوقاً عقارياً مستداماً، لا مجرد سوق يحقق أرقاماً قياسية.»",
     size=19, color=WHITE, bold=True, font=HEAD, first=True, line=1.5, space_after=10)
para(tf, "— لحظة مايا، تُختم دائمًا بالمشاهد لا بالضيف", size=13, color=WOOD, bold=True, font=HEAD, space_after=0)
_, tf = txt(s, 0.75, 5.05, SW-1.5, 1.8)
para(tf, "بعد المقابلة، تتوجه المقدِّمة مباشرة إلى الكاميرا — دون ضيف، دون نقاش — لتشرح بلغة بسيطة الأثر العملي على المستثمر، رائد الأعمال، صاحب الشركة، المستهلك، والاقتصاد عمومًا. لا تنحاز، ولا تكرّر كلام الضيف، بل تربط النقاش بفكرة اقتصادية أكبر.",
     size=13, color=INKSOFT, font=BODY, first=True, line=1.35, space_after=0)

# --- 07 PILOT INTRO ---
s = std(7, "الحلقة التجريبية · Pilot")
pill = rect(s, SW-3.1-0.75, 1.5, 3.1, 0.42, fill=RED, rounded=True)
_, tf = txt(s, SW-3.1-0.75, 1.5, 3.1, 0.42, anchor=MSO_ANCHOR.MIDDLE)
para(tf, "الحلقة الأولى — القطاع العقاري", size=12, color=WHITE, bold=True, font=HEAD, align=PP_ALIGN.CENTER, first=True, space_after=0)
title_block(s, ["من دبي إلى الرياض…", "ومن أبوظبي إلى القاهرة."], SW-6.0-0.75, 2.15, 6.0, size=27)
_, tf = txt(s, SW-6.0-0.75, 3.7, 6.0, 2.8)
para(tf, [("العقار لم يعد يهم المستثمرين فقط، بل كل شخص يبحث عن منزل، أو يدفع إيجارًا، أو يفكّر في أول تملّك. السؤال ليس «ماذا يحدث في السوق؟» بل: ", INKSOFT, False),
          ("كيف يؤثر ذلك على بيتك، وإيجارك، وفرصك؟", RED, True)], size=14, font=BODY, first=True, line=1.4, space_after=10)
para(tf, "الضيف المحتمل: العبار / بن غاطي / سجواني — أحد أبرز قادة القطاع.", size=11.5, color=INKSOFT, font=BODY, space_after=0)
stats = [("Q1","حجم الاستثمار العقاري في الربع الأول"),("٪","نسبة القروض العقارية من تمويلات البنوك"),("↑↓","ارتفاع الإيجارات أو تراجعها")]
for i,(big,lbl) in enumerate(stats):
    y = 2.2+i*1.5
    rect(s, 0.75, y, 5.5, 1.35, fill=PAPER2, lc=STONE, rounded=True)
    _, tf = txt(s, 1.0, y+0.12, 5.0, 1.1, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, [(big+"   ", RED, True), (lbl, INKSOFT, False)], size=14, font=HEAD, first=True, space_after=0, line=1.1)

# --- 08 RUNSHEET (table) ---
s = std(8, "رحلة الحلقة · Runsheet")
title_block(s, ["تشريح الحلقة التجريبية"], 0.75, 1.15, 8, size=28)
rows = [("المحتوى","الفقرة","المدة"),
        ("بنبرة موجَّهة لمشاهد بعينه: القطاع يحرّك البنوك ويكشف أين تتجه الأموال.","المقدّمة","1:30"),
        ("محتوى مرئي: حجم الاستثمار، نسبة القروض العقارية، حركة الإيجارات.","غرفة القرار (افتتاح)","~1:00"),
        ("ملفات ساخنة + سرد لكواليس القطاع + تبسيط + أرقام وتوقّعات حصرية.","محاور الحوار","16–18:00"),
        ("خبير/مستأجر/مستثمر يطرح التحدّي، فيردّ الضيف الرئيسي.","الكرسي المقابل","3–4:00"),
        ("المقدِّمة إلى الكاميرا: الأثر المباشر على المستثمر والمستهلك.","ماذا يعني لك؟","1:30"),
        ("ربط الخبر بحياة الناس (أسعار النفط، حوالات المغتربين…).","اقتصاد الناس","segment"),
        ("فكرة اقتصادية أكبر تُختم بالمشاهد: هل يرتفع المعروض بما يكفي؟","لحظة مايا","1:00–1:30")]
tb = s.shapes.add_table(len(rows), 3, Inches(0.75), Inches(1.9), Inches(SW-1.5), Inches(4.9)).table
tb.columns[0].width = Inches(7.0); tb.columns[1].width = Inches(2.9); tb.columns[2].width = Inches(1.93)
tb.first_row = False
for ri,row in enumerate(rows):
    tb.rows[ri].height = Inches(0.2)
    for ci,val in enumerate(row):
        cell = tb.cell(ri, ci); cell.margin_left=Pt(6); cell.margin_right=Pt(6); cell.margin_top=Pt(3); cell.margin_bottom=Pt(3)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid(); cell.fill.fore_color.rgb = (INK if ri==0 else (PAPER if ri%2 else WHITE))
        tf = cell.text_frame; tf.word_wrap = True
        col = WHITE if ri==0 else (RED if ci==2 else (INK if ci==1 else INKSOFT))
        bold = (ri==0) or (ci in (1,2))
        sz = 12 if ri==0 else (11 if ci==0 else 12)
        para(tf, val, size=sz, color=col, bold=bold, font=(HEAD if ci in(1,2) or ri==0 else BODY), first=True, space_after=0, line=1.1)

# --- 09 ECONOMY OF PEOPLE ---
s = std(9, "فقرة أساسية · اقتصاد الناس")
title_block(s, ["«ما الذي يعنيه", "هذا الخبر بالنسبة لي؟»"], SW-6.0-0.75, 1.5, 6.0, size=28)
_, tf = txt(s, SW-6.0-0.75, 3.1, 6.0, 1.2)
para(tf, "فقرة تربط الأخبار الاقتصادية بحياة الناس اليومية، وتشرح كيف تصل القرارات العالمية إلى كل بيت — سواء ارتبطت بحلقة اليوم أو بقضية أخرى.",
     size=13.5, color=INKSOFT, font=BODY, first=True, line=1.35, space_after=0)
rect(s, SW-6.0-0.75, 4.5, 6.0, 2.15, fill=PAPER2, lc=STONE, rounded=True)
_, tf = txt(s, SW-6.0-0.55, 4.68, 5.6, 1.9, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [("مثال:  ", INK, True), ("ارتفعت أسعار النفط… لكن هذا لا يؤثر على أصحاب السيارات فقط، بل يمتد إلى الشحن، وأسعار السلع، وتشغيل المصانع، وفواتير الكهرباء، وربما سعر تذكرة السفر. حتى لو لم تملك سيارة، قد يصل الارتفاع إلى جيبك.", INKSOFT, False)],
     size=12.5, font=BODY, first=True, line=1.35, space_after=0)
_, tf = txt(s, 0.75, 1.7, 5.7, 0.5)
para(tf, "أفكار أخرى قابلة للتناول", size=16, color=INK, bold=True, font=HEAD, first=True, space_after=0)
chips = ["حوالات المغتربين","البطالة","أسعار الفائدة","التضخّم","سعر الصرف"]
cx, cyy = 0.75, 2.35
for ch_ in chips:
    w_ = 0.5 + len(ch_)*0.13
    if cx + w_ > 6.5: cx = 0.75; cyy += 0.6
    rect(s, cx, cyy, w_, 0.45, fill=RED if ch_=="حوالات المغتربين" else WHITE, lc=None if ch_=="حوالات المغتربين" else LINE, rounded=True)
    _, tf = txt(s, cx, cyy, w_, 0.45, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, ch_, size=11.5, color=WHITE if ch_=="حوالات المغتربين" else INK, bold=True, font=HEAD, align=PP_ALIGN.CENTER, first=True, space_after=0)
    cx += w_ + 0.15
_, tf = txt(s, 0.75, 3.7, 5.7, 3.0)
para(tf, [("حوالات المغتربين ", INK, True), ("ليست مجرد أموال يرسلها شخص لعائلته، بل شريان اقتصادي يربط بلدين: يوفّر دخلًا لملايين الأسر ويضخّ عملة أجنبية في بلد الاستقبال، ويعكس قوة سوق العمل في بلد المصدر. لذلك يقرأها الاقتصاديون كمؤشر لا كتحويل مالي.", INKSOFT, False)],
     size=12.5, font=BODY, first=True, line=1.4, space_after=0)

# --- 10 CONTENT ECOSYSTEM ---
s = std(10, "منظومة المحتوى · Content Ecosystem")
title_block(s, ["مقابلة واحدة… تُغذّي أسبوعًا كاملًا."], 0.75, 1.2, SW-1.5, size=30)
_, tf = txt(s, 0.75, 2.05, SW-1.5, 0.6)
para(tf, "المقابلة ليست المنتج النهائي، بل المادة الخام التي تُبنى عليها منظومة تحريرية متكاملة عبر كل المنصّات.",
     size=14, color=INKSOFT, font=BODY, first=True, space_after=0)
items = ["الحلقة الكاملة","Hero Clip — أبرز مقطع","فقرة الكرسي المقابل","اقتصاد الناس","لحظة مايا",
         "فيديو عمودي × 3–4","تصاميم الاقتباسات","فيديوهات شرح المفاهيم","مقال تحليلي (SEO)",
         "فقرة النشرة البريدية","نسخة بودكاست","LinkedIn","Instagram Reels","TikTok","YouTube Shorts"]
cx, cyy = SW-0.75, 2.95
row_r = SW-0.75
for it in items:
    w_ = 0.5 + len(it)*0.115
    if cx - w_ < 0.75: cx = SW-0.75; cyy += 0.62
    red = it in ("الحلقة الكاملة","Hero Clip — أبرز مقطع")
    wood = it == "فيديو عمودي × 3–4"
    rect(s, cx-w_, cyy, w_, 0.46, fill=RED if red else (WOOD if wood else WHITE), lc=None if (red or wood) else LINE, rounded=True)
    _, tf = txt(s, cx-w_, cyy, w_, 0.46, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, it, size=11, color=WHITE if red else INK, bold=True, font=HEAD, align=PP_ALIGN.CENTER, rtl=not red, first=True, space_after=0)
    cx -= (w_ + 0.14)
hline(s, 0.75, 6.35, SW-1.5)
_, tf = txt(s, 0.75, 6.45, SW-1.5, 0.5)
para(tf, [("جلسة واحدة ← أسبوع كامل من التوزيع. ", INK, True), ("الثقل التحريري يُعيد ملء القمع دون تسجيل يومي.", INKSOFT, False)],
     size=13, font=BODY, first=True, space_after=0)

# --- 11 PRODUCTION + RULES ---
s = std(11, "نموذج الإنتاج والقواعد التحريرية")
_, tf = txt(s, SW-6.0-0.75, 1.35, 6.0, 0.5)
para(tf, "الإيقاع والإنتاج", size=17, color=INK, bold=True, font=HEAD, first=True, space_after=6)
for b in ["حلقة رئيسية أسبوعية — بث في ليلة ثابتة (الأحد/الاثنين).",
          "تسجيل حلقتين كل أسبوعين في يوم استوديو واحد.",
          "جدول نشر ثابت ومنتظم لتكريس عادة المشاهدة.",
          "حلقات «سريعة» عند حدث استثنائي لتبدو الفرنشايز حيّة."]:
    para(tf, "▪  "+b, size=13, color=INKSOFT, font=BODY, space_after=8, line=1.25)
_, tf = txt(s, 0.75, 1.35, 5.7, 0.5)
para(tf, "ثلاث قواعد تحمي الفرنشايز", size=17, color=INK, bold=True, font=HEAD, first=True, space_after=0)
for i,(n,h4,b) in enumerate([("01","ضيف واحد، قرار واحد","لا لجان نقاش، ولا استعراض مسيرة مهنية."),
                             ("02","الكرسي المقابل غير قابل للتفاوض","حتى مع أرفع الضيوف — وإلا مات مؤشر الاستقلالية."),
                             ("03","تنتهي كل حلقة بمال المشاهد","لا بمجد الضيف — هذا الفارق الجوهري.")]):
    card(s, 0.75, 2.0+i*1.55, 5.7, 1.4, n, h4, b)

# --- 12 SEASON 1 (table) ---
s = std(12, "الموسم الأول · 12 حلقة مبنيّة على قصص حيّة")
srows = [("#","الضيف / القرار","الكرسي المقابل")] + [
 ("1","قيادة صندوق الاستثمارات العامة — التحوّل 2026-30 نحو 80% محلي","مدير تخصيص عالمي"),
 ("2","صندوق سيادي / مبادلة — الأسواق الخاصة المُرمَّزة لصغار المستثمرين","متشكك من الأسهم الخاصة"),
 ("3","مؤسِّس يونيكورن / رأس مال مغامر — أول رهان كبير في الخليج","مؤسِّس لم يجمع تمويلًا"),
 ("4","شركة عائلية وورثتها — نقل الثروة (الحلقة المحورية)","الأجيال تتحدّى بعضها"),
 ("5","البنك المركزي / VARA — قواعد الأصول الرقمية والدرهم المستقر","رائد أعمال تشفير"),
 ("6","رئيس شركة مدرجة — تخصيص رأسمال (إعادة شراء، توسّع، أرباح)","مستثمر ناشط أو محلل"),
 ("7","تمويل المنشآت واللوجستيات — الائتمان وتكاليف التجارة","صاحب منشأة صغيرة"),
 ("8","قائد في العقارات / التكنولوجيا العقارية — العقار المُرمَّز","وسيط تقليدي أو منظِّم"),
 ("9","صانع سياسات / مصرفي مصري — استقرار الجنيه والمدّخر المغترب","مدّخر خسر من العملة"),
 ("10","مدير صندوق للمغتربين — توجيه رأس المال العربي نحو المنطقة","مستثمر مغترب احترق"),
 ("11","امرأة تدير مكتب عائلة أو صندوقًا — قرار تخصيص حقيقي","نظير أو عضو مجلس إدارة"),
 ("12","ختام الموسم: وزير / شخصية بمستوى FII — الرهان الكبير للعام","اقتصادي مغاير للسائد")]
tb = s.shapes.add_table(len(srows), 3, Inches(0.75), Inches(1.55), Inches(SW-1.5), Inches(5.3)).table
tb.columns[0].width = Inches(0.55); tb.columns[1].width = Inches(8.05); tb.columns[2].width = Inches(3.23)
tb.first_row=False
for ri,row in enumerate(srows):
    tb.rows[ri].height=Inches(0.1)
    for ci,val in enumerate(row):
        cell=tb.cell(ri,ci); cell.margin_left=Pt(6);cell.margin_right=Pt(6);cell.margin_top=Pt(1);cell.margin_bottom=Pt(1)
        cell.vertical_anchor=MSO_ANCHOR.MIDDLE
        cell.fill.solid(); cell.fill.fore_color.rgb=(INK if ri==0 else (PAPER if ri%2 else WHITE))
        tf=cell.text_frame; tf.word_wrap=True
        col=WHITE if ri==0 else (RED if ci==0 else (INK if ci==1 else INKSOFT))
        para(tf,val,size=10.5 if ri else 11, color=col, bold=(ri==0 or ci==0),
             font=(HEAD if (ri==0 or ci==0) else BODY),
             align=PP_ALIGN.CENTER if ci==0 else PP_ALIGN.RIGHT, first=True, space_after=0, line=1.05)

# --- 13 SET DESIGN ---
s = std(13, "الديكور والهوية البصرية · Set Design")
picture(s, "twoshot", 0.75, 1.6, 5.9, 5.0, caption="الطاولة نفسها عند حضور الضيف — طاولة وكراسٍ عالية")
title_block(s, ["بساطة نقيّة…", "ودفء طبيعي."], SW-5.7-0.75, 1.7, 5.7, size=32)
_, tf = txt(s, SW-5.7-0.75, 3.1, 5.7, 3.4)
para(tf, "استوديو أبيض بالكامل، بسيط ونظيف، طاولة عالية وكراسٍ عالية. المقدِّمة تنظر إلى الكاميرا في فقراتها المنفردة، وتنتقل إلى الطاولة نفسها عند وصول الضيف.",
     size=13.5, color=INKSOFT, font=BODY, first=True, line=1.4, space_after=10)
for b in [("القاعدة: ","أبيض دافئ + عنصر طبيعي واحد (خشب/حجر) يمنح الدفء دون ازدحام."),
          ("الإضاءة: ","ألواح ضوء معمارية ناعمة تمنح العمق بدل الجدران المزخرفة."),
          ("المبدأ: ","«رقيّ مع راهنية» — ثابت وفاخر لا موضة عابرة.")]:
    para(tf, [("▪  "+b[0], INK, True), (b[1], INKSOFT, False)], size=12.5, font=BODY, space_after=7, line=1.3)

# --- 14 PALETTE ---
s = std(14, "لوحة الخامات · Material Palette")
title_block(s, ["دفء طبيعي على قاعدة بيضاء"], 0.75, 1.5, 5.7, size=24)
_, tf = txt(s, 0.75, 2.35, 5.7, 1.4)
para(tf, "قاعدة بيضاء دافئة يكسرها عنصر طبيعي واحد، مع لمسة حمراء مقتصَدة تربط الهوية بـ CNN Business Arabic دون أن تسيطر.",
     size=13.5, color=INKSOFT, font=BODY, first=True, line=1.4, space_after=0)
sw = [("أبيض دافئ",PAPER,INK),("حجر",STONE,INK),("سنديان",WOOD,INK),("أحمر CNNBA",RED,WHITE),("حبر",INK,WHITE)]
for i,(name,c,tc) in enumerate(sw):
    x=0.75+i*1.12
    rect(s, x, 3.9, 1.0, 1.4, fill=c, lc=LINE, rounded=True)
    _, tf = txt(s, x, 5.35, 1.0, 0.4)
    para(tf, name, size=10.5, color=INK, bold=True, font=HEAD, align=PP_ALIGN.CENTER, first=True, space_after=0)
picture(s, "palette", SW-5.9-0.75, 1.7, 5.9, 4.9, caption="لوحة الخامات")

# --- 15 GRAPHICS ---
s = std(15, "هوية الشاشة · On-screen Graphics")
title_block(s, ["رسومات رقمية أولًا."], 0.75, 1.6, 5.7, size=30)
_, tf = txt(s, 0.75, 2.5, 5.7, 3.6)
para(tf, "لا رسومات تلفزيونية مضافة في المونتاج، بل نظام بصري مصمَّم للمنصّات: كتابات متحرّكة (kinetic captions) بالعربية، وبطاقات بيانات نظيفة تُقرأ في لمحة.",
     size=13.5, color=INKSOFT, font=BODY, first=True, line=1.4, space_after=10)
for b in [("Lower-third ","بخط عربي راقٍ + لمسة حمراء واحدة."),
          ("بطاقات بيانات ","بمساحة بيضاء واسعة وأرقام كبيرة."),
          ("ترجمة متحرّكة ","مدمجة من التصوير — أقوى إشارة على الطابع الرقمي.")]:
    para(tf, [("▪  "+b[0], INK, True), (b[1], INKSOFT, False)], size=12.5, font=BODY, space_after=7, line=1.3)
picture(s, "lower", SW-5.9-0.75, 1.7, 5.9, 4.4, caption="نموذج Lower-third + بطاقة بيانات")

# --- 16 DIGITAL / VERTICAL ---
s = std(16, "الإطار الرقمي · Digital-first")
picture(s, "vertical", 0.9, 1.55, 3.0, 5.2, caption="كادر عمودي 9:16")
title_block(s, ["مُصمَّم للعمودي…", "لا مُقتطَع منه."], SW-6.3-0.75, 1.8, 6.3, size=30)
_, tf = txt(s, SW-6.3-0.75, 3.2, 6.3, 1.3)
para(tf, "الإضاءة والكادر يخدمان الأفقي والعمودي معًا — مناطق آمنة لقصّ 9:16 من التصوير، لا كادر تلفزيوني يُبتر لاحقًا.",
     size=13.5, color=INKSOFT, font=BODY, first=True, line=1.4, space_after=8)
plats=["TikTok","Reels","Shorts","YouTube","LinkedIn","Podcast"]
cx=SW-0.75
for p_ in plats:
    w_=0.4+len(p_)*0.13
    if cx-w_<SW-6.3-0.75: cx=SW-0.75
    red=p_ in("TikTok","Reels","Shorts")
    rect(s, cx-w_, 4.5, w_, 0.44, fill=RED if red else WHITE, lc=None if red else LINE, rounded=True)
    _, tf=txt(s, cx-w_,4.5,w_,0.44,anchor=MSO_ANCHOR.MIDDLE)
    para(tf,p_,size=11,color=WHITE if red else INK,bold=True,font=HEAD,align=PP_ALIGN.CENTER,rtl=False,first=True,space_after=0)
    cx-=(w_+0.13)
_, tf=txt(s, SW-6.3-0.75, 5.2, 6.3, 1.4)
para(tf, [("تعديلات رقمية مقترحة:  ", INK, True), ("ضبط زمن الحلقة الكاملة عند ~22–24 دقيقة، واعتماد «الكرسي المقابل» و«لحظة مايا» كمقاطع البطولة العمودية.", INKSOFT, False)],
     size=12.5, font=BODY, first=True, line=1.35, space_after=0)

# --- 17 WHY IT WINS ---
s = std(17, "لماذا سينجح")
title_block(s, ["يُركِّز أندر ما تملكه CNNBA"], 0.75, 1.3, SW-1.5, size=30)
wins=[("الوصول","إلى أبرز صنّاع القرار الاقتصادي في المنطقة."),
      ("مصداقية تحريرية","راسخة وموثوقة."),
      ("أسلوب حواري مهني","يجمع الصرامة والإنصاف."),
      ("تبسيط القضايا","الاقتصادية بلغة واضحة."),
      ("سرد قصصي","مصمَّم أولًا للمنصّات الرقمية."),
      ("منظومة محتوى","قابلة للتوسّع وإعادة التوظيف.")]
for i,(h,b) in enumerate(wins):
    col=i%2; row=i//2
    x=0.75+col*6.05; y=2.3+row*0.95
    rect(s, x+5.8, y+0.05, 0.09, 0.55, fill=RED)  # right marker
    _, tf=txt(s, x, y, 5.7, 0.9)
    para(tf, [(h+"  ", INK, True),(b, INKSOFT, False)], size=14, font=BODY, first=True, line=1.25, space_after=0)
hline(s, 0.75, 5.55, SW-1.5)
_, tf=txt(s, 0.75, 5.75, SW-1.5, 0.9)
para(tf, "«لأن كل قرار اقتصادي كبير… يصل في النهاية إلى محفظة شخص، أو شركته، أو مستقبله.»",
     size=17, color=RED, bold=True, font=HEAD, first=True, line=1.3, space_after=0)

# --- 18 NEXT STEPS ---
s = std(18, "الخطوات التالية")
title_block(s, ["من الورق إلى الشاشة"], 0.75, 1.2, 8, size=30)
steps=[("01","إقرار الهوية","اعتماد الاسم، الشعار، لوحة الخامات ونظام الرسومات."),
       ("02","بناء الاستوديو","تنفيذ الديكور الأبيض والإضاءة وخطة الكاميرات (أفقي + عمودي)."),
       ("03","تصوير البايلوت","الحلقة العقارية + تأمين ضيف «الكرسي المقابل»."),
       ("04","تفعيل المنظومة","قصّ المقاطع، النشر عبر المنصّات، وقياس الأداء.")]
for i,(n,h4,b) in enumerate(steps):
    card(s, 0.75+i*3.02, 2.4, 2.85, 2.6, n, h4, b)

# --- 19 CLOSING ---
s = slide(INK)
_, tf = txt(s, 1.5, 2.2, SW-3.0, 3.2, anchor=MSO_ANCHOR.MIDDLE)
para(tf, "CNN BUSINESS ARABIC", size=13, color=WOOD, bold=True, font=HEAD, align=PP_ALIGN.CENTER, rtl=False, first=True, space_after=12)
para(tf, "The Business Show", size=52, color=WHITE, bold=True, font=HEAD, align=PP_ALIGN.CENTER, rtl=False, space_after=6)
para(tf, "البرنامج الاقتصادي الأسبوعي الرئيسي — يتضمّن فقرة «غرفة القرار»", size=16, color=WOOD, bold=True, font=HEAD, align=PP_ALIGN.CENTER, space_after=18)
para(tf, "مساحة تُطرح فيها أهم القرارات الاقتصادية في المنطقة، وتُناقش، وتُختبر، ثم تُترجم إلى لغة يفهمها كل مشاهد.",
     size=15, color=RGBColor(0xE9,0xE4,0xDA), font=BODY, align=PP_ALIGN.CENTER, space_after=20, line=1.4)
para(tf, "تنتهي كل حلقة بالمشاهد… وليس بالضيف.", size=15, color=WOOD, bold=True, font=HEAD, align=PP_ALIGN.CENTER, space_after=0)

# ---------------------------------------------------------------- motion
# Morph transition (with Fade fallback) between every slide — the main "motion".
for i, s in enumerate(prs.slides):
    add_transition(s)

out = "The_Business_Show.pptx"
prs.save(out)
embedded = sum(1 for k in IMG if _imgcache.get(k) is not None)
print(f"Saved {out}  ·  slides={len(prs.slides._sldIdLst)}  ·  images embedded={embedded}/{len(IMG)}")
if embedded < len(IMG):
    print("Note: some images used placeholders (CDN unreachable here). Run this script on a "
          "network-connected machine to embed the real photos automatically.")
