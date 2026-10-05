"""Arabic (RTL) PowerPoint builder on top of python-pptx.

Same principle as Word, different vocabulary
--------------------------------------------
PowerPoint has a full bidirectional text engine, so slides must be given
**logical-order, unshaped Unicode**. Never run `arabic_reshaper` on text headed
for a .pptx — PowerPoint would shape the presentation forms a second time,
producing broken glyphs and text that cannot be searched or edited.

Direction is expressed in DrawingML, not WordprocessingML:

| Level | Attribute | Without it |
|---|---|---|
| paragraph `<a:pPr>` | `rtl="1"` | Mixed Arabic/Latin reorders wrongly |
| paragraph `<a:pPr>` | `algn="r"` | Text hugs the left edge of its box |
| run `<a:rPr>` | child `<a:cs typeface="…"/>` | Arabic falls back to a default face |
| table `<a:tblPr>` | `rtl="1"` | First column renders on the left |

`python-pptx` exposes none of these, so they are set on the XML directly here.

Usage
-----
    from arabic_pptx import ArabicPptx, preview

    deck = ArabicPptx(cs_font="Arial")
    deck.title_slide("عنوان العرض", "سطر فرعي")
    deck.bullets_slide("المحاور", ["نقطة أولى", "نقطة ثانية"])
    deck.table_slide("التقنيات", ["الطبقة", "التقنية"], [["الواجهة", "Nuxt 3"]])
    deck.save("deck.pptx")
    preview("deck.pptx")
"""

from __future__ import annotations

import os
from copy import deepcopy
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Cm, Pt

ARABIC_RE = re.compile(
    r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]"
)
LATIN_RE = re.compile(r"[A-Za-z0-9]")
ARABIC_INDIC = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")

# Mirrored characters: rendered flipped when they land in a right-to-left run,
# so both halves of a pair have to share one run — see segment().
BRACKETS = {"(": ")", "[": "]", "{": "}", "<": ">", "«": "»", "‹": "›",
            "〈": "〉", "⟨": "⟩"}
CLOSERS = {close: open_ for open_, close in BRACKETS.items()}

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


# Neutral characters that belong to a Latin token when written flush against it:
# `+963`, `.NET`, `$5`, `@user` on the leading side; `C#`, `C++`, `50%` on the
# trailing side. Left as neutrals they resolve Arabic, land in the RTL run and
# are drawn on the far side of the token — `C#` renders as `#C`.
# `%` is deliberately absent: Arabic convention puts it on the left of the
# number (٪٥٠ / %50), which is what the bidi algorithm already does.
AFFIX_BEFORE = set("+-.#@$/~\\")
AFFIX_AFTER = set("+#")
LRE, PDF = "\u202a", "\u202c"   # LEFT-TO-RIGHT EMBEDDING … POP DIRECTIONAL FORMATTING


def _guard(chunk):
    """Fence a Latin run that starts or ends with a flush affix.

    Word and PowerPoint decide a neutral's direction from its run, so `C#` in
    its own `en-US` run is already right there. LibreOffice, PDF viewers and
    browsers ignore run boundaries and run the plain Unicode bidi algorithm
    across the paragraph, where `C#` at the edge of a Latin island renders
    `#C` and `.NET` renders `NET.`. An LRE…PDF pair around the token fixes it
    in both families. Measured on Word, PowerPoint and LibreOffice:

        plain run       Word ✓  PowerPoint ✓  LibreOffice ✗
        LRM on the edge Word ✗  PowerPoint ✓  LibreOffice ✓
        LRE … PDF       Word ✓  PowerPoint ✓  LibreOffice ✓
        LRI … PDI       correct, but Word draws the marks as visible boxes

    Only tokens that need it are wrapped; the marks are invisible.
    """
    if chunk and (chunk[0] in AFFIX_BEFORE or chunk[-1] in AFFIX_AFTER):
        return LRE + chunk + PDF
    return chunk


def _attach_affixes(text, kinds, resolved):
    """Pull flush prefixes/suffixes into the Latin run they are attached to.

    Sentence punctuation is deliberately not in the sets: the full stop in
    `مع Nuxt 3.` ends the Arabic sentence and must stay on the Arabic side.
    """
    n = len(text)
    for i in range(n):
        if kinds[i] != "L":
            continue
        j = i - 1
        while j >= 0 and kinds[j] == "N" and text[j] in AFFIX_BEFORE:
            j -= 1
        if j < i - 1 and (j < 0 or text[j].isspace() or kinds[j] == "L"
                          or text[j] in BRACKETS):
            resolved[j + 1:i] = ["L"] * (i - 1 - j)
        k = i + 1
        while k < n and kinds[k] == "N" and text[k] in AFFIX_AFTER:
            k += 1
        if k > i + 1:
            resolved[i + 1:k] = ["L"] * (k - i - 1)


def _bracket_pairs(text):
    """Return (open_index, close_index) for every matched mirrored pair."""
    stack, pairs = [], []
    for i, ch in enumerate(text):
        if ch in BRACKETS:
            stack.append((ch, i))
        elif ch in CLOSERS:
            for k in range(len(stack) - 1, -1, -1):
                if stack[k][0] == CLOSERS[ch]:
                    pairs.append((stack[k][1], i))
                    del stack[k:]
                    break
    return pairs


def segment(text):
    """Split into (chunk, is_arabic). Neutrals join the Arabic side unless they
    sit between two Latin tokens — see arabic-docx for the full reasoning.

    A matched bracket pair is then pulled wholly into the Latin run whenever
    either half resolved Latin. Otherwise `DECIMAL(5,2)` splits: the inner `(`
    stays Latin while the trailing `)` joins the Arabic run, where PowerPoint
    mirrors it and moves it to the far edge, rendering `(DECIMAL(5,2`. Pairs
    that resolved Arabic on both sides — `(API)`, `(1.25)` — are left alone;
    there both are mirrored and their positions swap, which cancels out.
    """
    if not text:
        return []

    def classify(ch):
        if ARABIC_RE.match(ch):
            return "A"
        return "L" if LATIN_RE.match(ch) else "N"

    kinds = [classify(ch) for ch in text]
    n = len(kinds)

    prev_strong, last = [None] * n, None
    for i, k in enumerate(kinds):
        prev_strong[i] = last
        if k != "N":
            last = k
    next_strong, nxt = [None] * n, None
    for i in range(n - 1, -1, -1):
        next_strong[i] = nxt
        if kinds[i] != "N":
            nxt = kinds[i]

    resolved = [
        k if k != "N"
        else ("L" if prev_strong[i] == "L" and next_strong[i] == "L" else "A")
        for i, k in enumerate(kinds)
    ]
    _attach_affixes(text, kinds, resolved)

    # Skip the span if it contains Arabic — that is an unpaired `<`/`>` used as
    # a comparison operator, not a bracket.
    for i, j in _bracket_pairs(text):
        if (resolved[i] == "L" or resolved[j] == "L") and "A" not in kinds[i:j + 1]:
            resolved[i:j + 1] = ["L"] * (j - i + 1)

    out, buf, cur = [], text[0], resolved[0]
    for ch, k in zip(text[1:], resolved[1:]):
        if k == cur:
            buf += ch
        else:
            out.append((buf, cur == "A"))
            buf, cur = ch, k
    out.append((buf, cur == "A"))
    return [(t, a) for t, a in out if t]


def arabic_digits(text, keep_latin_context=False):
    """Convert Western digits to Arabic-Indic (٠–٩).

    With `keep_latin_context=True`, a number stays Western when the nearest
    word before or after it is Latin — `Nuxt 3`, `Python 3.12` — and only
    numbers in Arabic context are converted. Use one digit system per deck.
    """
    if not keep_latin_context:
        return text.translate(ARABIC_INDIC)

    def repl(m):
        before = re.search(r"(\S+)\s*$", text[:m.start()])
        after = re.match(r"\s*(\S+)", text[m.end():])
        for word in (before, after):
            if word and re.search(r"[A-Za-z]", word.group(1)) \
                    and not ARABIC_RE.search(word.group(1)):
                return m.group(0)
        if m.start() and text[m.start() - 1].isalpha() \
                and not ARABIC_RE.match(text[m.start() - 1]):
            return m.group(0)
        return m.group(0).translate(ARABIC_INDIC)

    return re.sub(r"\d+(?:[.,]\d+)*", repl, text)


def _q(tag):
    """'a:rPr' / 'p:par' -> Clark notation."""
    prefix, name = tag.split(":")
    return f"{{{A_NS if prefix == 'a' else P_NS}}}{name}"


def rtl_positions(n, left, width, item_width, gap=None):
    """x offsets for `n` items in one row, item 0 on the RIGHT.

    Numbered or sequential content — steps, phases, timelines, cards labelled
    01, 02, 03 — is read right to left in Arabic. Laying it out left to right
    is the most common layout defect in Arabic decks, and nothing in the text
    engine fixes it: the order is decided by these coordinates. Pass EMU (or
    python-pptx Length) values; the row is centred inside [left, left+width].
    """
    if gap is None:
        gap = (width - n * item_width) / max(n - 1, 1) if n > 1 else 0
    total = n * item_width + (n - 1) * gap
    start = left + (width - total) / 2
    return [int(start + (n - 1 - i) * (item_width + gap)) for i in range(n)]


def make_rtl_defaults(prs, rtl_lang="ar-SA"):
    """Make text added later in PowerPoint itself default to RTL.

    The builder marks every paragraph it writes, but text the presenter types
    into the finished deck takes its direction from the presentation's
    `defaultTextStyle` and the slide master's `txStyles`. Both default to LTR
    in every stock template, so new text boxes start at the left. Alignment is
    only changed where it was left-aligned or unset; centred styles stay
    centred.
    """
    roots = [prs.part._element.find(_q("p:defaultTextStyle"))]
    for master in prs.slide_masters:
        tx = master._element.find(_q("p:txStyles"))
        if tx is not None:
            roots.extend(tx)
    for root in roots:
        if root is None:
            continue
        for lvl in root:
            if lvl.tag.endswith("pPr") and "lvl" in lvl.tag:
                lvl.set("rtl", "1")
                if lvl.get("algn") in (None, "l"):
                    lvl.set("algn", "r")
    # Language of new text: an Arabic default stops PowerPoint from flagging
    # every Arabic word as an English spelling error.
    for root in roots:
        if root is None:
            continue
        for rpr in root.iter(_q("a:defRPr")):
            rpr.set("lang", rtl_lang)


_TRANSITIONS = {"fade", "dissolve", "cut"}


def add_transition(slide, kind="fade", speed="med"):
    """Slide transition, inserted at its schema position.

    Only direction-neutral transitions are offered. Directional ones (wipe,
    push, cover) carry an LTR default and have to be mirrored by hand for an
    RTL deck — and checked visually — so they are left out on purpose.
    """
    if kind not in _TRANSITIONS:
        raise ValueError(f"transition {kind!r}: use one of {sorted(_TRANSITIONS)}")
    sld = slide._element
    old = sld.find(_q("p:transition"))
    if old is not None:
        sld.remove(old)
    tr = sld.makeelement(_q("p:transition"), {"spd": speed})
    tr.append(tr.makeelement(_q(f"p:{kind}"), {}))
    for successor in ("p:timing", "p:extLst"):
        el = sld.find(_q(successor))
        if el is not None:
            el.addprevious(tr)
            return tr
    sld.append(tr)
    return tr


def reading_order(shapes):
    """Sort shapes into Arabic reading order: top band first, then right to left.

    Shapes whose tops are within a quarter of the tallest shape's height are
    treated as one row.
    """
    shapes = list(shapes)
    if not shapes:
        return []
    band = max(s.height for s in shapes) / 4
    rows = []
    for s in sorted(shapes, key=lambda s: s.top):
        if rows and abs(s.top - rows[-1][0].top) <= band:
            rows[-1].append(s)
        else:
            rows.append([s])
    return [s for row in rows for s in sorted(row, key=lambda s: -(s.left + s.width))]


def set_build(slide, steps, dur_ms=350, stagger_ms=60):
    """Entrance animations: one click per step, fade in, items staggered.

    steps  list of steps; each step is a list of targets revealed together on
           one click. A target is a shape, or (shape, paragraph_index) to
           reveal one paragraph of a text box — the usual bullet-by-bullet
           build. Order targets with `reading_order()` for RTL content.

    Fade is used throughout because it has no direction: fly-in and wipe
    default to entering from the left, which reads backwards in RTL. Replaces
    any existing animation on the slide. The XML is the same structure
    PowerPoint writes for "Fade, On Click / With Previous".
    """
    sld = slide._element
    old = sld.find(_q("p:timing"))
    if old is not None:
        sld.remove(old)
    ids = iter(range(3, 100000))
    built, para_built = [], set()

    def tgt(target):
        if isinstance(target, tuple):
            shape, idx = target
            para_built.add(shape.shape_id)
            return (f'<p:spTgt spid="{shape.shape_id}"><p:txEl>'
                    f'<p:pRg st="{idx}" end="{idx}"/></p:txEl></p:spTgt>')
        if target.shape_id not in built:
            built.append(target.shape_id)
        return f'<p:spTgt spid="{target.shape_id}"/>'

    def effect(target, first, delay):
        t = tgt(target)
        node = "clickEffect" if first else "withEffect"
        return (f'<p:par><p:cTn id="{next(ids)}" presetID="10" presetClass="entr" '
                f'presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
                f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f'<p:tgtEl>{t}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility'
                f'</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/>'
                f'</p:to></p:set><p:animEffect transition="in" filter="fade"><p:cBhvr>'
                f'<p:cTn id="{next(ids)}" dur="{dur_ms}"/><p:tgtEl>{t}</p:tgtEl></p:cBhvr>'
                f'</p:animEffect></p:childTnLst></p:cTn></p:par>')

    clicks = []
    for step in steps:
        if not step:
            continue
        outer, inner = next(ids), next(ids)
        effects = "".join(effect(t, i == 0, i * stagger_ms) for i, t in enumerate(step))
        clicks.append(f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst>'
                      f'<p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par>'
                      f'<p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/>'
                      f'</p:stCondLst><p:childTnLst>{effects}</p:childTnLst></p:cTn></p:par>'
                      f'</p:childTnLst></p:cTn></p:par>')
    if not clicks:
        return

    shapes_by_id = {s.shape_id: s for s in slide.shapes}
    bld = "".join(f'<p:bldP spid="{sid}" grpId="0" build="p"/>' for sid in sorted(para_built))
    bld += "".join(f'<p:bldP spid="{sid}" grpId="0"/>' for sid in built
                   if sid not in para_built and sid in shapes_by_id
                   and shapes_by_id[sid].has_text_frame)
    xml = (f'<p:timing xmlns:p="{P_NS}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" '
           f'restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" '
           f'nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
           f'<p:childTnLst>{"".join(clicks)}</p:childTnLst></p:cTn><p:prevCondLst>'
           f'<p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
           f'</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl>'
           f'<p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst>'
           f'</p:cTn></p:par></p:tnLst>' + (f'<p:bldLst>{bld}</p:bldLst>' if bld else "")
           + '</p:timing>')
    timing = etree.fromstring(xml)
    ext = sld.find(_q("p:extLst"))
    if ext is not None:
        ext.addprevious(timing)
    else:
        sld.append(timing)
    return timing


class ArabicPptx:
    """Build a right-to-left slide deck.

    Slides are laid out with explicit text boxes rather than placeholders,
    because placeholder geometry varies between templates and the goal here is
    predictable RTL output.

    transition  "fade" (or "dissolve") applies a slide transition to every
                slide at save time; None leaves slides without one.
    accent      hex colour used by steps_slide().
    """

    def __init__(self, cs_font="Arial", latin_font="Arial", template=None,
                 widescreen=True, base_size=18, rtl_lang="ar-SA", ltr_lang="en-US",
                 transition=None, accent="1F4E79"):
        self.prs = Presentation(template) if template else Presentation()
        if template is None:
            self.prs.slide_width = Cm(33.87 if widescreen else 25.4)
            self.prs.slide_height = Cm(19.05)
        self.cs_font, self.latin_font, self.base = cs_font, latin_font, base_size
        # rtl_lang drives PowerPoint's bidi engine: "fa-IR", "ur-PK", "he-IL".
        self.rtl_lang, self.ltr_lang = rtl_lang, ltr_lang
        self.transition, self.accent = transition, RGBColor.from_string(accent)
        self.W = self.prs.slide_width
        self.H = self.prs.slide_height
        self.blank = self._blank_layout()
        make_rtl_defaults(self.prs, rtl_lang)

    def _blank_layout(self):
        """The layout with no placeholders — by content, not by index.

        `slide_layouts[6]` is "Blank" only in python-pptx's own default
        template; in a corporate template index 6 can be anything.
        """
        layouts = list(self.prs.slide_layouts)
        for layout in layouts:
            if layout.name.strip().lower() == "blank":
                return layout
        return min(layouts, key=lambda l: len(l.placeholders))

    # -- internals ---------------------------------------------------------

    def _rtl_paragraph(self, p, align="r"):
        pPr = p._p.get_or_add_pPr()
        pPr.set("rtl", "1")
        pPr.set("algn", align)
        return p

    def _fill(self, p, text, size, bold=False, color=None):
        """Add runs to a paragraph, splitting mixed scripts and tagging each
        run with its language.

        `lang` on <a:rPr> is the switch that decides everything. PowerPoint uses
        it to pick the shaping and bidi behaviour for the run, and python-pptx
        never sets it. Measured on the same sentence:

            no lang at all      → "PostgreSQLو"  (boundary spaces collapse)
            lang="ar-SA" on all → "3Nuxt"        (Latin numbers reorder)
            lang per run        → correct

        So the split is not cosmetic: it exists so each fragment can carry the
        right language. This is the DrawingML counterpart of `w:lang w:bidi` in
        Word, and it fails just as silently.

        xml:space="preserve" is set as well, since OOXML otherwise strips
        leading and trailing whitespace from text elements.
        """
        for chunk, is_arabic in segment(text):
            run = p.add_run()
            run.text = chunk if is_arabic else _guard(chunk)
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.name = self.latin_font
            if color:
                run.font.color.rgb = color

            t = run._r.find(f"{{{A_NS}}}t")
            if t is not None:
                t.set(f"{{{XML_NS}}}space", "preserve")

            rPr = run._r.get_or_add_rPr()
            rPr.set("lang", self.rtl_lang if is_arabic else self.ltr_lang)
            # The complex-script slot is a separate child of <a:rPr>; without it
            # Arabic ignores font.name entirely.
            cs = rPr.makeelement(f"{{{A_NS}}}cs", {"typeface": self.cs_font})
            rPr.append(cs)
        self._end_props(p, size)
        return p

    def _end_props(self, p, size):
        """`a:endParaRPr` sets what the presenter types at the end of the line.

        Without it, text added in PowerPoint after the generated text comes in
        as English at the template's default size and font.
        """
        end = p._p.find(_q("a:endParaRPr"))
        if end is None:
            end = p._p.makeelement(_q("a:endParaRPr"), {})
            p._p.append(end)
        end.set("lang", self.rtl_lang)
        end.set("sz", str(int(round(size * 100))))
        if end.find(_q("a:cs")) is None:
            end.append(end.makeelement(_q("a:latin"), {"typeface": self.latin_font}))
            end.append(end.makeelement(_q("a:cs"), {"typeface": self.cs_font}))

    def _bullet(self, p, char="•", indent_cm=0.9):
        """A real bullet: `a:buChar` with a hanging indent.

        In an `rtl="1"` paragraph PowerPoint applies `marL`/`indent` on the
        right, so the bullet sits on the right and wrapped lines align under
        the text, not under the bullet — which a typed "•  " cannot do.
        """
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(Cm(indent_cm))))
        pPr.set("indent", str(-int(Cm(indent_cm))))
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
            for el in pPr.findall(_q(tag)):
                pPr.remove(el)
        font = pPr.makeelement(_q("a:buFont"), {"typeface": "Arial"})
        bu = pPr.makeelement(_q("a:buChar"), {"char": char})
        anchor = next((el for el in pPr if el.tag in (_q("a:tabLst"), _q("a:defRPr"),
                                                       _q("a:extLst"))), None)
        if anchor is not None:
            anchor.addprevious(font)
            anchor.addprevious(bu)
        else:
            pPr.append(font)
            pPr.append(bu)

    def _textbox(self, slide, left, top, width, height):
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        return tf

    def _slide(self):
        return self.prs.slides.add_slide(self.blank)

    def _fit(self, text, size, width, max_lines=1, floor=None):
        """Largest size <= `size` at which `text` fits in `max_lines`.

        An estimate (average Arabic glyph ≈ 0.55 em), used so a long title
        shrinks instead of overflowing into the body. The floor keeps it
        readable; past the floor the text wraps.
        """
        floor = floor or max(self.base + 4, 24)
        width_pt = width / 12700
        while size > floor and len(text) * size * 0.55 > width_pt * max_lines:
            size -= 1
        return size

    def _heading(self, slide, text, size=None):
        width = self.W - Cm(3)
        size = self._fit(text, size or self.base + 10, width)
        tf = self._textbox(slide, Cm(1.5), Cm(1.2), width, Cm(2.4))
        p = tf.paragraphs[0]
        self._rtl_paragraph(p)
        self._fill(p, text, size, bold=True)
        return tf

    # -- public API --------------------------------------------------------

    def title_slide(self, title, subtitle=""):
        slide = self._slide()
        tf = self._textbox(slide, Cm(2), self.H / 3, self.W - Cm(4), Cm(6))
        p = tf.paragraphs[0]
        self._rtl_paragraph(p, align="ctr")
        self._fill(p, title, self.base + 20, bold=True)
        if subtitle:
            sp = tf.add_paragraph()
            self._rtl_paragraph(sp, align="ctr")
            self._fill(sp, subtitle, self.base + 2)
        return self

    def bullets_slide(self, title, items, marker="•", reveal=False):
        """Heading plus a bulleted list.

        reveal  True builds the list one bullet per click (fade).
        """
        slide = self._slide()
        self._heading(slide, title)
        tf = self._textbox(slide, Cm(1.5), Cm(4.5), self.W - Cm(3),
                           self.H - Cm(6))
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            self._rtl_paragraph(p)
            p.space_after = Pt(14)
            self._bullet(p, marker)
            self._fill(p, item, self.base)
        if reveal:
            body = slide.shapes[-1]
            set_build(slide, [[(body, i)] for i in range(len(items))])
        return self

    def text_slide(self, title, paragraphs):
        slide = self._slide()
        self._heading(slide, title)
        tf = self._textbox(slide, Cm(1.5), Cm(4.5), self.W - Cm(3),
                           self.H - Cm(6))
        for i, text in enumerate(paragraphs):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            self._rtl_paragraph(p)
            p.space_after = Pt(12)
            self._fill(p, text, self.base)
        return self

    def steps_slide(self, title, steps, reveal=False):
        """A row of numbered steps — 01 on the RIGHT — joined by ← arrows.

        steps   2–5 (heading, text) tuples, or plain strings.
        reveal  True reveals one step per click, in reading order.

        Arrows are drawn as the character ← because bidi mirroring applies to
        brackets, not arrows: → typed in an RTL deck still points right, i.e.
        backwards.
        """
        slide = self._slide()
        self._heading(slide, title)
        n = len(steps)
        gap = Cm(1.4)
        card_w = int((self.W - Cm(3) - gap * (n - 1)) / n)
        top, card_h = Cm(5), min(self.H - Cm(7.5), Cm(7))
        xs = rtl_positions(n, Cm(1.5), self.W - Cm(3), card_w, gap)
        groups = []
        for i, (x, step) in enumerate(zip(xs, steps)):
            head, body = step if isinstance(step, tuple) else (step, "")
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top,
                                          card_w, card_h)
            card.adjustments[0] = 0.08
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(0xF3, 0xF6, 0xFA)
            card.line.color.rgb = self.accent
            card.line.width = Pt(1.25)
            tf = card.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP
            tf.margin_left = tf.margin_right = Cm(0.4)
            tf.margin_top = Cm(0.5)
            p = tf.paragraphs[0]
            self._rtl_paragraph(p, align="ctr")
            self._fill(p, f"{i + 1:02d}", self.base + 14, bold=True, color=self.accent)
            p2 = tf.add_paragraph()
            self._rtl_paragraph(p2, align="ctr")
            p2.space_before = Pt(6)
            self._fill(p2, head, self.base + 2, bold=True,
                       color=RGBColor(0x1F, 0x1F, 0x1F))
            if body:
                p3 = tf.add_paragraph()
                self._rtl_paragraph(p3, align="ctr")
                p3.space_before = Pt(6)
                self._fill(p3, body, self.base - 2, color=RGBColor(0x40, 0x40, 0x40))
            group = [card]
            if i < n - 1:
                ax = x - gap
                arrow = self._textbox(slide, ax, top + card_h / 2 - Cm(1), gap, Cm(2))
                ap = arrow.paragraphs[0]
                self._rtl_paragraph(ap, align="ctr")
                self._fill(ap, "←", self.base + 10, bold=True, color=self.accent)
                group.append(slide.shapes[-1])
            groups.append(group)
        if reveal:
            set_build(slide, groups)
        return self

    def image_slide(self, title, image_path, caption=None):
        """Heading plus one picture scaled to fit the body area, centred."""
        slide = self._slide()
        self._heading(slide, title)
        box_w, box_h = self.W - Cm(3), self.H - Cm(6.5 if caption else 5.5)
        pic = slide.shapes.add_picture(str(image_path), Cm(1.5), Cm(4.2))
        ratio = min(box_w / pic.width, box_h / pic.height, 1.0)
        pic.width, pic.height = int(pic.width * ratio), int(pic.height * ratio)
        pic.left = int((self.W - pic.width) / 2)
        if caption:
            tf = self._textbox(slide, Cm(1.5), pic.top + pic.height + Cm(0.3),
                               self.W - Cm(3), Cm(1.2))
            p = tf.paragraphs[0]
            self._rtl_paragraph(p, align="ctr")
            self._fill(p, caption, self.base - 2, color=RGBColor(0x40, 0x40, 0x40))
        return self

    def table_slide(self, title, header, rows, widths_cm=None):
        """Add a slide with an RTL table.

        Columns are passed in natural reading order — the first element is the
        rightmost column. `rtl="1"` on <a:tblPr> performs the flip, so unlike
        the PDF builder no manual reversal is needed.
        """
        slide = self._slide()
        self._heading(slide, title)

        cols, n_rows = len(header), len(rows) + 1
        left, top = Cm(1.5), Cm(4.5)
        width = self.W - Cm(3)
        height = Cm(1.2) * n_rows

        shape = slide.shapes.add_table(n_rows, cols, left, top, width, height)
        table = shape.table
        table._tbl.tblPr.set("rtl", "1")

        if widths_cm:
            for col, w in zip(table.columns, widths_cm):
                col.width = Cm(w)

        def cell(r, c, text, size, bold=False):
            p = table.cell(r, c).text_frame.paragraphs[0]
            self._rtl_paragraph(p)
            if not ARABIC_RE.search(text):
                # Latin-only cell: right-aligned like its neighbours, but LTR,
                # so "v2.0." or "(beta)" keep their punctuation in place.
                p._p.get_or_add_pPr().set("rtl", "0")
            self._fill(p, text, size, bold=bold)

        for c, text in enumerate(header):
            cell(0, c, str(text), self.base, bold=True)
        for r, row in enumerate(rows, start=1):
            for c, text in enumerate(row):
                cell(r, c, str(text), max(self.base - 2, 18))
        return self

    def save(self, path):
        if self.transition:
            for slide in self.prs.slides:
                add_transition(slide, self.transition)
        self.prs.save(str(path))
        return str(path)


# --------------------------------------------------------------------------
# Verification
# --------------------------------------------------------------------------

def find_powerpoint():
    """Locate POWERPNT.EXE across Office versions. Windows only — the conversion
    below drives PowerPoint through COM, which does not exist elsewhere."""
    if sys.platform != "win32":
        return None
    roots = [os.environ.get("ProgramFiles", r"C:\Program Files"),
             os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")]
    subdirs = ["root/Office16", "Office16", "root/Office15", "Office15", "Office14"]
    for root in roots:
        for sub in subdirs:
            cand = Path(root) / "Microsoft Office" / sub / "POWERPNT.EXE"
            if cand.exists():
                return str(cand)
    return None


def find_soffice():
    """Locate LibreOffice — PATH first, then the usual install locations on
    Windows, macOS and Linux, since it is rarely on PATH on the first two."""
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice", "/usr/local/bin/soffice", "/snap/bin/libreoffice",
    ]
    return next((c for c in candidates if Path(c).exists()), None)


NO_CONVERTER = (
    "Cannot render a preview: neither Microsoft PowerPoint nor LibreOffice was "
    "found. Install LibreOffice (https://www.libreoffice.org) — headless "
    "conversion is enough — or open the .pptx on a machine that has PowerPoint. "
    "The deck itself is already written; only the visual check needs a converter."
)


PPT_FAILED = (
    "Microsoft PowerPoint was found but failed to convert the deck, and there "
    "is no LibreOffice install to fall back to. PowerPoint reported:\n{}"
)

# Pure ASCII, and paths arrive through the environment rather than being
# interpolated in. Windows PowerShell 5.1 reads a BOM-less script as the system
# ANSI codepage, so an Arabic path baked into the text comes out mangled,
# PowerPoint cannot find the file, and the conversion fails for every
# Arabic-named deck. Environment variables are passed as Unicode, so the script
# never has to carry a non-ASCII character.
PPT_SCRIPT = r'''
$ErrorActionPreference = "Stop"
$src = $env:ARABIC_PPTX_SRC
$pdf = $env:ARABIC_PPTX_PDF
$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($src, $true, $false, $false)
    $pres.SaveAs($pdf, 32)    # 32 = ppSaveAsPDF
    $pres.Close()
} finally {
    $app.Quit()
}
'''


def _console_text(raw):
    """Decode PowerShell output, which is console-codepage bytes, not UTF-8."""
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(enc).strip()
        except UnicodeDecodeError:
            continue
    return repr(raw)


def _ppt_to_pdf(pptx_path, pdf_path):
    """Drive PowerPoint through COM. None on success, else the failure text."""
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                     encoding="utf-8-sig") as fh:
        fh.write(PPT_SCRIPT)
        ps = fh.name
    env = dict(os.environ, ARABIC_PPTX_SRC=str(pptx_path),
               ARABIC_PPTX_PDF=str(pdf_path))
    try:
        proc = subprocess.run(["powershell", "-NoProfile", "-NonInteractive",
                               "-ExecutionPolicy", "Bypass", "-File", ps],
                              capture_output=True, env=env)
    finally:
        Path(ps).unlink(missing_ok=True)
    if proc.returncode == 0 and Path(pdf_path).exists():
        return None
    return (_console_text(proc.stderr) or _console_text(proc.stdout)
            or f"powershell exited with {proc.returncode} and wrote no PDF")


def pptx_to_pdf(pptx_path, pdf_path=None):
    """Convert through PowerPoint (or LibreOffice) so slides can be rendered.

    A PowerPoint failure is carried, not swallowed. Reporting "no converter
    found" when PowerPoint is installed and merely errored sends the caller off
    installing LibreOffice for a problem that has nothing to do with it.
    """
    pptx_path = Path(pptx_path).resolve()
    pdf_path = Path(pdf_path).resolve() if pdf_path else pptx_path.with_suffix(".pdf")
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    ppt_error = None
    if find_powerpoint():
        ppt_error = _ppt_to_pdf(pptx_path, pdf_path)
        if ppt_error is None:
            return str(pdf_path)

    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(PPT_FAILED.format(ppt_error) if ppt_error
                           else NO_CONVERTER)

    try:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                        "--outdir", str(pdf_path.parent), str(pptx_path)],
                       check=True, capture_output=True)
    except subprocess.CalledProcessError as exc:
        detail = _console_text(exc.stderr) or _console_text(exc.stdout)
        raise RuntimeError(
            f"LibreOffice failed to convert {pptx_path.name}: {detail}"
            + (f"\nPowerPoint was tried first and failed: {ppt_error}" if ppt_error else "")
        ) from exc
    # LibreOffice names the output after the *source* file, ignoring pdf_path.
    produced = pdf_path.parent / f"{pptx_path.stem}.pdf"
    if produced != pdf_path and produced.exists():
        produced.replace(pdf_path)
    return str(pdf_path)


def preview(pptx_path, out_dir=None, scale=1.5):
    """Convert to PDF and render every slide to PNG. View these images."""
    import pypdfium2 as pdfium

    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir) if out_dir else pptx_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(pptx_to_pdf(pptx_path, out_dir / f"{pptx_path.stem}.pdf"))
    paths = []
    for i in range(len(pdf)):
        png = out_dir / f"{pptx_path.stem}_s{i + 1}.png"
        pdf[i].render(scale=scale).to_pil().save(png)
        paths.append(str(png))
    # Release the file handle: on Windows an open PdfDocument locks the PDF,
    # and the next conversion to the same path fails inside Word/PowerPoint
    # with a bare E_FAIL.
    pdf.close()
    return paths


# --------------------------------------------------------------------------
# Audit and repair of existing decks
# --------------------------------------------------------------------------

PRESENTATION_FORMS_RE = re.compile(r"[ﭐ-﷿ﹰ-﻿]")


def _para_text(p):
    return "".join(t.text or "" for t in p.iter(_q("a:t")))


def _iter_slide_paragraphs(slide):
    """(shape_element, a:p) for every paragraph on a slide, tables and groups included."""
    for p in slide._element.iter(_q("a:p")):
        sp = p
        while sp is not None and sp.tag not in (_q("p:sp"), _q("p:graphicFrame")):
            sp = sp.getparent()
        yield sp, p


def _shape_name(sp):
    c = sp.find(".//" + _q("p:cNvPr")) if sp is not None else None
    return c.get("name", "") if c is not None else ""


def _lvl_attr(container, lvl, attr):
    """attr of <a:lvlNpPr> inside a list-style container, or None."""
    if container is None:
        return None
    el = container.find(_q(f"a:lvl{lvl + 1}pPr"))
    return el.get(attr) if el is not None else None


def _effective_rtl(prs, slide, sp, p):
    """Resolve a paragraph's rtl through the inheritance chain.

    Direct pPr → the shape's own lstStyle → for placeholders, the master's
    title/body text style → the presentation's defaultTextStyle. A deck whose
    defaults are RTL needs no flag on each paragraph.
    """
    pPr = p.find(_q("a:pPr"))
    lvl = int(pPr.get("lvl", "0")) if pPr is not None else 0
    if pPr is not None and pPr.get("rtl") is not None:
        return pPr.get("rtl")
    if sp is not None:
        lst = sp.find(".//" + _q("a:lstStyle"))
        v = _lvl_attr(lst, lvl, "rtl")
        if v is not None:
            return v
        ph = sp.find(".//" + _q("p:ph"))
        if ph is not None:
            tx = slide.slide_layout.slide_master._element.find(_q("p:txStyles"))
            kind = "titleStyle" if ph.get("type") in ("title", "ctrTitle") else "bodyStyle"
            v = _lvl_attr(tx.find(_q(f"p:{kind}")) if tx is not None else None, lvl, "rtl")
            if v is not None:
                return v
    v = _lvl_attr(prs.part._element.find(_q("p:defaultTextStyle")), lvl, "rtl")
    return v or "0"


def audit_pptx(path, min_pt=18):
    """Report the RTL defects in a deck that do not show up in its text.

    Returns a list of findings ("slide 3, 'Title 1': …"), empty when clean.
    Direction is resolved through the inheritance chain, so a deck with RTL
    defaults is not flagged paragraph by paragraph.

    `min_pt` flags explicitly sized text below a readable size, one summary
    line per slide. 18 pt is a floor, not a target: titles of 40+ pt and
    lists of 24+ pt read from the back of a hall. Footers and page numbers
    will show up here too — judge those by eye.
    """
    prs = Presentation(str(path))
    issues = []
    for n, slide in enumerate(prs.slides, 1):
        small = []
        for sp, p in _iter_slide_paragraphs(slide):
            text = _para_text(p)
            if not text.strip():
                continue
            where = f"slide {n}, {_shape_name(sp)!r}: «{text.strip()[:40]}»"
            if PRESENTATION_FORMS_RE.search(text):
                issues.append(f"{where} — pre-shaped Arabic presentation forms "
                              "(arabic_reshaper on PowerPoint text): renders broken, "
                              "not searchable")
            rtl = _effective_rtl(prs, slide, sp, p)
            has_ar = bool(ARABIC_RE.search(text))
            has_latin = bool(re.search(r"[A-Za-z]", text))
            if has_ar and rtl not in ("1", "true"):
                issues.append(f"{where} — Arabic paragraph that is not RTL: "
                              "punctuation and Latin terms land on the wrong side")
            edge = re.search(r"^[^\w\s]|[^\w\s]$", text.strip())
            if not has_ar and has_latin and rtl in ("1", "true") and edge:
                issues.append(f"{where} — Latin-only paragraph marked rtl=\"1\": "
                              "edge punctuation jumps to the other end")
            for r in p.iter(_q("a:r")):
                rt = "".join(t.text or "" for t in r.iter(_q("a:t")))
                rpr = r.find(_q("a:rPr"))
                if ARABIC_RE.search(rt) and re.search(r"[A-Za-z]", rt):
                    issues.append(f"{where} — one run mixes Arabic and Latin: "
                                  "spaces collapse (PostgreSQLو) or numbers flip (3Nuxt)")
                lang = rpr.get("lang") if rpr is not None else None
                if has_ar and has_latin and rt.strip() and lang is None:
                    issues.append(f"{where} — mixed paragraph with runs that carry "
                                  "no lang: boundary spaces collapse (PostgreSQLو)")
                if ARABIC_RE.search(rt) and lang is not None and lang.startswith("en"):
                    issues.append(f"{where} — Arabic run tagged lang={lang!r}: "
                                  "PowerPoint applies LTR rules to it")
                if rpr is not None and rpr.get("sz") and int(rpr.get("sz")) < min_pt * 100:
                    small.append((int(rpr.get("sz")) / 100, text.strip()[:30]))
        if small:
            size, sample = min(small)
            issues.append(f"slide {n} — {len(small)} run(s) below {min_pt} pt "
                          f"(smallest {size:g} pt: «{sample}»)")
        for tbl in slide._element.iter(_q("a:tbl")):
            text = "".join(t.text or "" for t in tbl.iter(_q("a:t")))
            tblPr = tbl.find(_q("a:tblPr"))
            if ARABIC_RE.search(text) and (tblPr is None or tblPr.get("rtl") != "1"):
                issues.append(f"slide {n} — table with Arabic but no rtl=\"1\": "
                              "first column renders on the left")
    # De-duplicate identical messages (a long paragraph repeats per run).
    return list(dict.fromkeys(issues))


def _split_run(r, rtl_lang, ltr_lang, cs_font):
    t = r.find(_q("a:t"))
    pieces = segment(t.text or "") if t is not None else []
    if len(pieces) < 2:
        pieces = [(t.text if t is not None else "", bool(ARABIC_RE.search(
            t.text or "")))]
    out = []
    for chunk, is_arabic in pieces:
        new = deepcopy(r)
        nt = new.find(_q("a:t"))
        if nt is not None:
            nt.text = chunk if is_arabic else _guard(chunk)
            nt.set(f"{{{XML_NS}}}space", "preserve")
        rpr = new.find(_q("a:rPr"))
        if rpr is None:
            rpr = new.makeelement(_q("a:rPr"), {})
            new.insert(0, rpr)
        rpr.set("lang", rtl_lang if is_arabic else ltr_lang)
        latin = rpr.find(_q("a:latin"))
        if rpr.find(_q("a:cs")) is None and (cs_font or latin is not None):
            cs = rpr.makeelement(_q("a:cs"), {
                "typeface": cs_font or latin.get("typeface")})
            # schema order: latin, ea, cs, sym, hlinkClick …
            after = next((el for el in rpr if el.tag in (
                _q("a:sym"), _q("a:hlinkClick"), _q("a:hlinkMouseOver"),
                _q("a:rtl"), _q("a:extLst"))), None)
            if after is not None:
                after.addprevious(cs)
            else:
                rpr.append(cs)
        out.append(new)
    for new in out:
        r.addprevious(new)
    r.getparent().remove(r)


def fix_pptx(src, dst, rtl_lang="ar-SA", ltr_lang="en-US", cs_font=None,
             tables=False):
    """Retrofit RTL onto an existing deck; returns `dst`.

    For decks made elsewhere — a template, an export, slides built LTR. The
    original structure, positions and formatting are kept:

      - paragraphs containing Arabic get rtl="1"; alignment is changed to
        right only where it was left or unset, so centred titles stay centred;
      - Latin-only paragraphs that carry rtl="1" and start or end with
        punctuation get rtl="0", which moves "Hello, world!" back from
        "!Hello, world" (others are left alone, so list bullets stay on the
        right);
      - mixed-script runs are split, every run gets its own `lang`, and an
        `a:cs` font is added where the run names a Latin font but no
        complex-script one;
      - presentation defaults and master text styles become RTL, so text the
        presenter adds later starts on the right.

    `tables=True` also sets rtl="1" on tables containing Arabic. It is off by
    default because it flips the visual column order: right for a table typed
    in reading order, wrong for one whose columns were reversed by hand.
    Layout — the left-to-right order of cards, steps and timelines — is not
    touched; `audit_pptx` cannot see it either. Check those slides by eye.
    """
    prs = Presentation(str(src))
    make_rtl_defaults(prs, rtl_lang)
    for slide in prs.slides:
        for sp, p in list(_iter_slide_paragraphs(slide)):
            text = _para_text(p)
            if not text.strip():
                continue
            pPr = p.find(_q("a:pPr"))
            if pPr is None:
                pPr = p.makeelement(_q("a:pPr"), {})
                p.insert(0, pPr)
            if ARABIC_RE.search(text):
                pPr.set("rtl", "1")
                if pPr.get("algn") in (None, "l"):
                    pPr.set("algn", "r")
                for r in list(p.findall(_q("a:r"))):
                    _split_run(r, rtl_lang, ltr_lang, cs_font)
                end = p.find(_q("a:endParaRPr"))
                if end is not None:
                    end.set("lang", rtl_lang)
            elif re.search(r"[A-Za-z]", text) \
                    and _effective_rtl(prs, slide, sp, p) in ("1", "true") \
                    and re.search(r"^[^\w\s]|[^\w\s]$", text.strip()):
                pPr.set("rtl", "0")
        if tables:
            for tbl in slide._element.iter(_q("a:tbl")):
                if ARABIC_RE.search("".join(t.text or "" for t in tbl.iter(_q("a:t")))):
                    tbl.find(_q("a:tblPr")).set("rtl", "1")
    prs.save(str(dst))
    return str(dst)
