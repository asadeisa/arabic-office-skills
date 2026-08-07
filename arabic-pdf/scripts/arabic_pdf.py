"""Arabic (RTL) PDF builder on top of reportlab.

Why this exists
---------------
reportlab has no bidirectional text engine. Producing correct Arabic requires
solving four separate problems, and getting any one of them wrong silently
produces a document that *looks* plausible but reads as gibberish:

1. **Glyph shaping.** Arabic letters change form depending on their neighbours.
   `arabic_reshaper` converts logical characters into presentation forms.

2. **Visual reordering.** `python-bidi` reorders the string. It must be told
   `base_dir="R"` explicitly — otherwise it infers direction from the first
   strong character, so a line starting with "PostgreSQL" is laid out
   left-to-right and every Arabic run after it lands in the wrong place.

3. **Mirroring and Latin islands.** `python-bidi` reorders but never applies
   rule L4, so a bracket moved to the other end of an Arabic span keeps its
   original glyph and `(20 فأكثر)` prints as `)20 فأكثر(`. And `<`/`>` are not
   Unicode *paired brackets*, so the algorithm's bracket rule does not hold
   `Array<String>` together. `prepare()` handles both.

4. **Line wrapping order.** After reshaping, the string is in *visual* order.
   If reportlab wraps it, it moves what it thinks are trailing words to the
   next line — but in visual order those are the *leading* words, so lines come
   out shuffled. The fix is to wrap manually before reshaping, reshape each
   line on its own, and join with <br/> so reportlab never re-wraps.

Tables need a fourth fix: reportlab lays columns out left-to-right, so the
column order is reversed here to make the first logical column appear rightmost.

Usage
-----
    from arabic_pdf import ArabicPDF, preview

    pdf = ArabicPDF("report.pdf", doc_title="تقرير")
    pdf.title("عنوان المستند")
    pdf.heading("أولاً — المقدمة")
    pdf.para("نص عربي مع مصطلحات مثل PostgreSQL و FastAPI.")
    pdf.table(["العمود الأول", "العمود الثاني"], [["قيمة", "قيمة"]], [5, 11])
    pdf.bullets(["نقطة أولى", "نقطة ثانية"])
    pdf.save()
    preview("report.pdf")   # PNG per page, for visual verification
"""

from __future__ import annotations

import os
import re
from pathlib import Path

import arabic_reshaper
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

try:  # python-bidi >= 0.6 moved get_display to the package root
    from bidi import get_display
except ImportError:  # pragma: no cover
    from bidi.algorithm import get_display


# --------------------------------------------------------------------------
# Fonts
# --------------------------------------------------------------------------

# Ordered by preference. Each entry is (regular, bold). The first pair whose
# regular face exists on disk wins; a missing bold face falls back to regular.
FONT_CANDIDATES = [
    (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
    (r"C:\Windows\Fonts\tahoma.ttf", r"C:\Windows\Fonts\tahomabd.ttf"),
    ("/System/Library/Fonts/Supplemental/Arial.ttf",
     "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    ("/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
     "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Bold.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
]

_REGISTERED = False
REG, BOLD = "ArabicRegular", "ArabicBold"


def register_fonts(regular: str | None = None, bold: str | None = None) -> tuple[str, str]:
    """Register Arabic-capable TTFs. Pass explicit paths to override discovery."""
    global _REGISTERED
    if _REGISTERED and regular is None:
        return REG, BOLD

    if regular is None:
        for cand_reg, cand_bold in FONT_CANDIDATES:
            if os.path.exists(cand_reg):
                regular, bold = cand_reg, (cand_bold if os.path.exists(cand_bold) else cand_reg)
                break
        else:
            raise FileNotFoundError(
                "No Arabic-capable font found. Install one (e.g. Noto Naskh Arabic or "
                "Amiri) and pass its path: register_fonts('/path/Regular.ttf', '/path/Bold.ttf')"
            )
    if bold is None or not os.path.exists(bold):
        bold = regular

    pdfmetrics.registerFont(TTFont(REG, regular))
    pdfmetrics.registerFont(TTFont(BOLD, bold))
    _REGISTERED = True
    return REG, BOLD


# --------------------------------------------------------------------------
# Text shaping
# --------------------------------------------------------------------------

ARABIC_RE = re.compile(
    r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]"
)
LATIN_RE = re.compile(r"[A-Za-z0-9]")

# Mirrored characters and their partners, for rule L4 below.
BRACKETS = {"(": ")", "[": "]", "{": "}", "<": ">", "«": "»", "‹": "›",
            "〈": "〉", "⟨": "⟩"}
CLOSERS = {close: open_ for open_, close in BRACKETS.items()}
MIRROR = str.maketrans({**BRACKETS, **CLOSERS})

LRI, PDI = "⁦", "⁩"   # left-to-right isolate / pop directional isolate


def _bracket_pairs(text: str) -> list[tuple[int, int]]:
    """Return (open_index, close_index) for every matched mirrored pair."""
    stack: list[tuple[str, int]] = []
    pairs: list[tuple[int, int]] = []
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


def prepare(text: str) -> str:
    """Mirror right-to-left brackets and fence off the Latin spans.

    Two repairs `get_display` does not make on its own:

    * **Rule L4 — mirroring.** python-bidi reorders but never swaps a mirrored
      character for its partner, so `(20 فأكثر)` comes back with the parens in
      exchanged positions and their original glyphs, and renders `)20 فأكثر(`.
      Substituting here puts the right glyph where the reorder will move it.
    * **Latin islands.** `<` and `>` are mirrored but are not Unicode *paired
      brackets*, so the algorithm's bracket rule — the one that keeps
      `DECIMAL(5,2)` intact — does not cover them. The `>` of `Array<String>`
      simply takes the direction of the Arabic after it and is thrown to the
      far side. Wrapping each Latin span in LRI…PDI makes it one left-to-right
      island, which is the same thing a run boundary buys in Word.

    A bracket is right-to-left unless it sits between two Latin tokens, and a
    matched pair takes a side together — otherwise `DECIMAL(5,2)`, whose inner
    paren is Latin and whose trailing one is not, would be half-mirrored.
    """
    if not text:
        return text

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
    for i, j in _bracket_pairs(text):
        if (resolved[i] == "L" or resolved[j] == "L") and "A" not in kinds[i:j + 1]:
            resolved[i:j + 1] = ["L"] * (j - i + 1)

    out = []
    for i, ch in enumerate(text):
        if resolved[i] == "A":
            out.append(ch.translate(MIRROR))
            continue
        if i == 0 or resolved[i - 1] != "L":
            out.append(LRI)
        out.append(ch)
        if i == n - 1 or resolved[i + 1] != "L":
            out.append(PDI)
    return "".join(out)


def shape(text: str) -> str:
    """Reshape Arabic glyphs and reorder to visual RTL order.

    base_dir="R" is not optional: without it a line beginning with a Latin word
    is treated as left-to-right and the Arabic that follows is misplaced.

    Reshaping runs *first* so the isolates are never inserted between two
    Arabic letters, where they would break the joining. The isolates are then
    dropped once the reorder has used them, so no invisible control character
    reaches the PDF — some fonts draw those as empty boxes. Removing them
    cannot disturb the order of what is left.
    """
    visual = get_display(prepare(arabic_reshaper.reshape(text)), base_dir="R")
    return visual.replace(LRI, "").replace(PDI, "")


ARABIC_DIGITS = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def arabic_digits(text: str) -> str:
    """Convert Western digits to Arabic-Indic. Purely cosmetic."""
    return text.translate(ARABIC_DIGITS)


def _escape(text: str) -> str:
    """Escape for reportlab's mini-XML — *after* shaping, never before.

    Escaping first sends `&lt;` through the bidi pass as four separate
    characters. Its trailing `;` is a neutral, so it resolves to the
    surrounding Arabic and is ejected to the far side of the Latin island:
    `Array<String>` renders as `Array<String> ;`.
    """
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _wrap(text: str, style: ParagraphStyle, max_width: float) -> str:
    """Wrap manually, reshape per line, join with <br/>.

    The 8pt safety margin absorbs small differences between stringWidth here and
    reportlab's own measurement; without it an occasional line overflows and
    reportlab re-wraps it, which reverses that line's word order.
    """
    limit = max_width - 8
    lines: list[str] = []
    for part in text.split("\n"):
        cur: list[str] = []
        for word in part.split():
            trial = " ".join(cur + [word])
            fits = pdfmetrics.stringWidth(shape(trial), style.fontName, style.fontSize) <= limit
            if not cur or fits:
                cur.append(word)
            else:
                lines.append(" ".join(cur))
                cur = [word]
        lines.append(" ".join(cur))
    return "<br/>".join(_escape(shape(l)) for l in lines)


def rtl_paragraph(text: str, style: ParagraphStyle, max_width: float) -> Paragraph:
    """Build a Paragraph whose Arabic reads correctly at the given width."""
    return Paragraph(_wrap(text, style, max_width), style)


# --------------------------------------------------------------------------
# Document builder
# --------------------------------------------------------------------------

class ArabicPDF:
    """Incremental builder for simple right-to-left documents.

    Widths are given in centimetres because that is how page layouts are
    usually reasoned about; they are converted to points internally.
    """

    def __init__(self, path, doc_title="", pagesize=A4, margin_cm=2.2,
                 top_cm=2.0, bottom_cm=2.0, base_size=10.5, page_numbers=False):
        register_fonts()
        self.path = str(path)
        self.pagesize = pagesize
        self.page_numbers = page_numbers
        self.margin = margin_cm * cm
        self.top, self.bottom = top_cm * cm, bottom_cm * cm
        self.body_width = pagesize[0] - 2 * self.margin
        self.doc_title = doc_title
        self.story: list = []

        b = base_size
        self.s_title = ParagraphStyle("title", fontName=BOLD, fontSize=b + 5.5,
                                      leading=b + 13, alignment=TA_CENTER,
                                      textColor=colors.black)
        self.s_subtitle = ParagraphStyle("subtitle", fontName=REG, fontSize=b,
                                         leading=b + 7, alignment=TA_CENTER,
                                         textColor=colors.black)
        self.s_heading = ParagraphStyle("heading", fontName=BOLD, fontSize=b + 2,
                                        leading=b + 11, alignment=TA_RIGHT,
                                        textColor=colors.black,
                                        spaceBefore=14, spaceAfter=4)
        self.s_body = ParagraphStyle("body", fontName=REG, fontSize=b,
                                     leading=b + 8.5, alignment=TA_RIGHT,
                                     textColor=colors.black)
        self.s_cell = ParagraphStyle("cell", fontName=REG, fontSize=b - 0.5,
                                     leading=b + 4.5, alignment=TA_RIGHT,
                                     textColor=colors.black)
        self.s_cell_head = ParagraphStyle("cellhead", fontName=BOLD, fontSize=b - 0.5,
                                          leading=b + 4.5, alignment=TA_RIGHT,
                                          textColor=colors.black)

    # -- content -----------------------------------------------------------

    def title(self, text):
        self.story.append(rtl_paragraph(text, self.s_title, self.body_width))
        return self

    def subtitle(self, text):
        self.story.append(rtl_paragraph(text, self.s_subtitle, self.body_width))
        return self

    def heading(self, text):
        self.story.append(rtl_paragraph(text, self.s_heading, self.body_width))
        return self

    def para(self, text):
        self.story.append(rtl_paragraph(text, self.s_body, self.body_width))
        return self

    def bullets(self, items, marker="–"):
        for item in items:
            self.story.append(
                rtl_paragraph(f"{marker}  {item}", self.s_body, self.body_width))
        return self

    def table(self, header, rows, widths_cm, grid=True):
        """Add a table. `header` and each row are in natural reading order
        (first element = rightmost column); columns are reversed internally.

        Cell text may contain "\\n" to force a line break.
        """
        widths = [w * cm for w in widths_cm]
        pad = 16  # left + right cell padding, subtracted before wrapping

        def row_cells(values, style):
            return [rtl_paragraph(str(v), style, w - pad)
                    for v, w in zip(reversed(values), reversed(widths))]

        data = [row_cells(header, self.s_cell_head)]
        data += [row_cells(r, self.s_cell) for r in rows]

        t = Table(data, colWidths=list(reversed(widths)), hAlign="CENTER")
        style = [
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ]
        if grid:
            style.append(("GRID", (0, 0), (-1, -1), 0.6, colors.black))
        t.setStyle(TableStyle(style))
        # KeepTogether stops a single orphan row from being pushed to a new page.
        self.story.append(KeepTogether(t))
        return self

    def spacer(self, points=10):
        self.story.append(Spacer(1, points))
        return self

    def page_break(self):
        self.story.append(PageBreak())
        return self

    # -- output ------------------------------------------------------------

    def _stamp(self, canvas, doc):
        if not self.page_numbers:
            return
        canvas.saveState()
        canvas.setFont(REG, 9)
        canvas.setFillColor(colors.black)
        canvas.drawCentredString(self.pagesize[0] / 2, self.bottom / 2,
                                 shape(arabic_digits(str(doc.page))))
        canvas.restoreState()

    def save(self):
        doc = SimpleDocTemplate(
            self.path, pagesize=self.pagesize,
            leftMargin=self.margin, rightMargin=self.margin,
            topMargin=self.top, bottomMargin=self.bottom,
            title=self.doc_title or None,
        )
        doc.build(self.story, onFirstPage=self._stamp, onLaterPages=self._stamp)
        return self.path


# --------------------------------------------------------------------------
# Verification
# --------------------------------------------------------------------------

def preview(pdf_path, out_dir=None, scale=2):
    """Render each page to PNG so the result can actually be looked at.

    Arabic defects (reversed lines, orphaned words, missing glyphs shown as
    boxes) are invisible in extracted text — they only show up in the render.
    Always view these images before delivering a document.
    """
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:  # pragma: no cover
        raise ImportError("preview() needs pypdfium2:  pip install pypdfium2") from exc

    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir) if out_dir else pdf_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(str(pdf_path))
    paths = []
    for i in range(len(pdf)):
        png = out_dir / f"{pdf_path.stem}_p{i + 1}.png"
        pdf[i].render(scale=scale).to_pil().save(png)
        paths.append(str(png))
    return paths
