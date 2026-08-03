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

import re
import subprocess
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Cm, Pt

ARABIC_RE = re.compile(
    r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]"
)
LATIN_RE = re.compile(r"[A-Za-z0-9]")
ARABIC_INDIC = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def segment(text):
    """Split into (chunk, is_arabic). Neutrals join the Arabic side unless they
    sit between two Latin tokens — see arabic-docx for the full reasoning.
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

    out, buf, cur = [], text[0], resolved[0]
    for ch, k in zip(text[1:], resolved[1:]):
        if k == cur:
            buf += ch
        else:
            out.append((buf, cur == "A"))
            buf, cur = ch, k
    out.append((buf, cur == "A"))
    return [(t, a) for t, a in out if t]


def arabic_digits(text):
    return text.translate(ARABIC_INDIC)


class ArabicPptx:
    """Build a right-to-left slide deck.

    Slides are laid out with explicit text boxes rather than placeholders,
    because placeholder geometry varies between templates and the goal here is
    predictable RTL output.
    """

    def __init__(self, cs_font="Arial", latin_font="Arial", template=None,
                 widescreen=True, base_size=18, rtl_lang="ar-SA", ltr_lang="en-US"):
        self.prs = Presentation(template) if template else Presentation()
        if template is None:
            self.prs.slide_width = Cm(33.87 if widescreen else 25.4)
            self.prs.slide_height = Cm(19.05)
        self.cs_font, self.latin_font, self.base = cs_font, latin_font, base_size
        # rtl_lang drives PowerPoint's bidi engine: "fa-IR", "ur-PK", "he-IL".
        self.rtl_lang, self.ltr_lang = rtl_lang, ltr_lang
        self.W = self.prs.slide_width
        self.H = self.prs.slide_height
        self.blank = self.prs.slide_layouts[6]   # 6 = completely blank

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
            run.text = chunk
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
        return p

    def _textbox(self, slide, left, top, width, height):
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        return tf

    def _slide(self):
        return self.prs.slides.add_slide(self.blank)

    def _heading(self, slide, text, size=None):
        tf = self._textbox(slide, Cm(1.5), Cm(1.2), self.W - Cm(3), Cm(2.4))
        p = tf.paragraphs[0]
        self._rtl_paragraph(p)
        self._fill(p, text, size or self.base + 10, bold=True)
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

    def bullets_slide(self, title, items, marker="•"):
        slide = self._slide()
        self._heading(slide, title)
        tf = self._textbox(slide, Cm(1.5), Cm(4.5), self.W - Cm(3),
                           self.H - Cm(6))
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            self._rtl_paragraph(p)
            p.space_after = Pt(14)
            self._fill(p, f"{marker}  {item}", self.base)
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

        for c, text in enumerate(header):
            p = table.cell(0, c).text_frame.paragraphs[0]
            self._rtl_paragraph(p)
            self._fill(p, str(text), self.base - 2, bold=True)

        for r, row in enumerate(rows, start=1):
            for c, text in enumerate(row):
                p = table.cell(r, c).text_frame.paragraphs[0]
                self._rtl_paragraph(p)
                self._fill(p, str(text), self.base - 3)
        return self

    def save(self, path):
        self.prs.save(str(path))
        return str(path)


# --------------------------------------------------------------------------
# Verification
# --------------------------------------------------------------------------

PPT_PATHS = [
    r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
]


def pptx_to_pdf(pptx_path, pdf_path=None):
    """Convert through PowerPoint (or LibreOffice) so slides can be rendered."""
    pptx_path = Path(pptx_path).resolve()
    pdf_path = Path(pdf_path).resolve() if pdf_path else pptx_path.with_suffix(".pdf")

    if any(Path(p).exists() for p in PPT_PATHS):
        script = f'''
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open("{pptx_path}", $true, $false, $false)
$pres.SaveAs("{pdf_path}", 32)
$pres.Close()
$app.Quit()
'''
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                         encoding="utf-8") as fh:
            fh.write(script)
            ps = fh.name
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive",
                        "-ExecutionPolicy", "Bypass", "-File", ps],
                       check=True, capture_output=True)
        Path(ps).unlink(missing_ok=True)
        return str(pdf_path)

    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", str(pdf_path.parent), str(pptx_path)],
                   check=True, capture_output=True)
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
    return paths
