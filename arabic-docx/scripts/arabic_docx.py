"""Arabic (RTL) Word document builder on top of python-docx.

The critical difference from PDF generation
-------------------------------------------
Do NOT reshape text with `arabic_reshaper` here. reportlab and Pillow have no
text engine, so they must be handed final visual-order glyphs. Word has a full
bidirectional text engine and shapes at render time — feeding it pre-shaped
presentation forms makes it shape them a second time, producing broken glyphs
and text that cannot be searched, copied, or spell-checked.

Word needs **logical-order, unshaped Unicode** plus the right direction flags.
python-docx exposes almost none of those flags, so they are written directly
into the underlying XML here.

What actually has to be set (each one matters independently):

| Where | Element | Without it |
|---|---|---|
| settings.xml | `<w:themeFontLang w:bidi="ar-SA"/>` | Word never engages its RTL layout engine |
| styles.xml docDefaults | `<w:lang w:bidi="ar-SA"/>`, `<w:bidi/>`, `<w:jc w:val="start"/>` | New paragraphs revert to LTR |
| each `<w:p>` | `<w:bidi/>` + `<w:jc w:val="start"/>` | Paragraph indents and alignment mirror wrongly |
| each `<w:r>` | `<w:rtl/>` | Runs reorder incorrectly in mixed text |
| each `<w:tbl>` | `<w:bidiVisual/>` | Column order flips — first column lands on the left |
| each `<w:sectPr>` | `<w:bidi/>` | Margins, columns and page furniture stay LTR |

Use `w:jc w:val="start"` rather than `right`: `start` is direction-relative and
stays correct if a paragraph turns out to be LTR, whereas hardcoded `right`
breaks mixed documents.

Fonts occupy three independent slots per run. Arabic is drawn from the
complex-script slot, so setting only `w:ascii`/`w:hAnsi` silently leaves Arabic
on a default face at the wrong size:

    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>   + <w:szCs w:val="24"/>

Mixed Arabic/Latin text is split automatically into separate runs, with `w:rtl`
set only on the Arabic ones. A single run holding both scripts is the most
common defect in generated Arabic documents.

Usage
-----
    from arabic_docx import ArabicDocx, docx_to_pdf

    doc = ArabicDocx(cs_font="Arial", size=12)
    doc.title("عنوان المستند")
    doc.heading("أولاً — المقدمة")
    doc.para("نص عربي مع مصطلحات مثل PostgreSQL و FastAPI.")
    doc.table(["البند", "التفصيل"], [["قاعدة البيانات", "PostgreSQL"]], [5, 11])
    doc.bullets(["نقطة أولى", "نقطة ثانية"])
    doc.save("out.docx")
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

# Arabic-script ranges: Arabic, Supplement, Extended-A, Presentation Forms A/B.
ARABIC_RE = re.compile(
    r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]"
)
LATIN_RE = re.compile(r"[A-Za-z0-9]")
ARABIC_INDIC = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")

# Mirrored characters: rendered flipped when they land in a right-to-left run.
# Both halves of a pair therefore have to share one run — see segment().
BRACKETS = {"(": ")", "[": "]", "{": "}", "<": ">", "«": "»", "‹": "›",
            "〈": "〉", "⟨": "⟩"}
CLOSERS = {close: open_ for open_, close in BRACKETS.items()}


# --------------------------------------------------------------------------
# Low-level XML helpers
# --------------------------------------------------------------------------

def _sub(parent, tag):
    """Return child `tag`, creating it if absent."""
    el = parent.find(qn(tag))
    if el is None:
        el = parent.makeelement(qn(tag), {})
        parent.append(el)
    return el


def _flag(parent, tag, val=None):
    """Set a boolean-ish OOXML flag element such as <w:bidi/> or <w:rtl/>."""
    el = _sub(parent, tag)
    if val is not None:
        el.set(qn("w:val"), val)
    return el


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
    """Split text into (chunk, is_arabic) pieces.

    Neutral characters — spaces, punctuation, brackets — carry no direction of
    their own, so which run they land in decides where they appear. In an RTL
    paragraph a neutral belongs with the Arabic side unless it sits *between*
    two Latin tokens. Attaching a boundary space to the Latin run instead makes
    it collapse at the direction switch, which is what produces the classic
    `PostgreSQLو` with the space visually swallowed.

    Consequences of the rule:
      "Vue.js 3 مع Nuxt 3."  → ["Vue.js 3"] [" مع "] ["Nuxt 3"] ["."]
      "(20 فأكثر)"           → ["("] ["20"] [" فأكثر)"]

    That rule alone splits bracket pairs, which is a second, worse defect. In
    `DECIMAL(5,2)` the inner `(` sits between two Latin tokens and stays Latin,
    while the trailing `)` has Arabic after it and goes to the Arabic run —
    where Word mirrors it into `(` and moves it to the far edge of the Latin
    island, rendering `(DECIMAL(5,2`. So a matched pair is forced into one run
    whenever either half resolved Latin.

    `(API)` and `(1.25)` need no such repair and must not get it: there both
    brackets resolve Arabic, both are mirrored, and their positions swap, which
    cancels out. Only pairs that are actually split are touched. The span is
    left alone if it contains Arabic — that means an unpaired `<`/`>` used as a
    comparison operator, not a bracket.
    """
    if not text:
        return []

    def classify(ch):
        if ARABIC_RE.match(ch):
            return "A"
        return "L" if LATIN_RE.match(ch) else "N"

    kinds = [classify(ch) for ch in text]
    n = len(kinds)

    # Nearest strong kind on each side, so neutrals can be resolved in one pass.
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
    """Convert Western digits to Arabic-Indic (٠–٩)."""
    return text.translate(ARABIC_INDIC)


# --------------------------------------------------------------------------
# Builder
# --------------------------------------------------------------------------

class ArabicDocx:
    """Build a right-to-left Word document.

    cs_font   font used for the Arabic (complex-script) slot
    latin_font font used for Latin text
    size      base point size, applied to both slots
    """

    def __init__(self, cs_font="Arial", latin_font="Arial", size=12, template=None):
        self.doc = Document(template) if template else Document()
        self.cs_font, self.latin_font, self.size = cs_font, latin_font, size
        self._patch_settings()
        self._patch_styles()

    # -- document-wide switches -------------------------------------------

    def _patch_settings(self):
        """themeFontLang is the switch Word actually checks for RTL layout."""
        settings = self.doc.settings.element
        tfl = _sub(settings, "w:themeFontLang")
        tfl.set(qn("w:val"), "en-US")
        tfl.set(qn("w:bidi"), "ar-SA")

    def _patch_styles(self):
        styles = self.doc.styles.element
        defaults = _sub(styles, "w:docDefaults")

        rpr = _sub(_sub(defaults, "w:rPrDefault"), "w:rPr")
        lang = _sub(rpr, "w:lang")
        lang.set(qn("w:val"), "en-US")
        lang.set(qn("w:bidi"), "ar-SA")
        fonts = _sub(rpr, "w:rFonts")
        fonts.set(qn("w:ascii"), self.latin_font)
        fonts.set(qn("w:hAnsi"), self.latin_font)
        fonts.set(qn("w:cs"), self.cs_font)

        ppr = _sub(_sub(defaults, "w:pPrDefault"), "w:pPr")
        _flag(ppr, "w:bidi")
        _flag(ppr, "w:jc", "start")

    def _patch_sections(self):
        for section in self.doc.sections:
            _flag(section._sectPr, "w:bidi")

    # -- paragraph / run internals ----------------------------------------

    def _rtl_para(self, p, align="start"):
        ppr = p._p.get_or_add_pPr()
        _flag(ppr, "w:bidi")
        _flag(ppr, "w:jc", align)
        return p

    def _add_runs(self, p, text, bold=False, size=None):
        """Split mixed text and mark only the Arabic runs as RTL."""
        pts = size or self.size
        for chunk, is_arabic in segment(text):
            run = p.add_run(chunk)
            run.bold = bold
            run.font.size = Pt(pts)
            rpr = run._r.get_or_add_rPr()

            fonts = _sub(rpr, "w:rFonts")
            fonts.set(qn("w:ascii"), self.latin_font)
            fonts.set(qn("w:hAnsi"), self.latin_font)
            fonts.set(qn("w:cs"), self.cs_font)
            # Complex-script size is separate; without it Arabic ignores the size.
            _flag(rpr, "w:szCs", str(int(pts * 2)))

            if is_arabic:
                _flag(rpr, "w:rtl")
        return p

    # -- public API --------------------------------------------------------

    def title(self, text, size=None):
        p = self.doc.add_paragraph()
        self._rtl_para(p, align="center")
        self._add_runs(p, text, bold=True, size=size or self.size + 6)
        return self

    def subtitle(self, text):
        p = self.doc.add_paragraph()
        self._rtl_para(p, align="center")
        self._add_runs(p, text, size=self.size)
        return self

    def heading(self, text, size=None):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        self._rtl_para(p)
        self._add_runs(p, text, bold=True, size=size or self.size + 2)
        return self

    def para(self, text, bold=False):
        p = self.doc.add_paragraph()
        self._rtl_para(p)
        self._add_runs(p, text, bold=bold)
        return self

    def bullets(self, items, marker="–"):
        for item in items:
            self.para(f"{marker}  {item}")
        return self

    def table(self, header, rows, widths_cm=None, style="Table Grid"):
        """Add an RTL table.

        `header` and each row are given in natural reading order — the first
        element is the rightmost column. `w:bidiVisual` performs the flip, so
        the column order is NOT reversed here (unlike the PDF builder, where
        reportlab has no equivalent and reversal must be done by hand).
        """
        t = self.doc.add_table(rows=0, cols=len(header))
        t.style = style
        _flag(t._tbl.tblPr, "w:bidiVisual")

        def fill(values, bold):
            cells = t.add_row().cells
            for cell, value in zip(cells, values):
                cell.text = ""
                p = cell.paragraphs[0]
                self._rtl_para(p)
                self._add_runs(p, str(value), bold=bold, size=self.size - 0.5)

        fill(header, True)
        for row in rows:
            fill(row, False)

        if widths_cm:
            for row in t.rows:
                for cell, w in zip(row.cells, widths_cm):
                    cell.width = Cm(w)
        return self

    def page_break(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        return self

    def spacer(self, points=10):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(points)
        return self

    def save(self, path):
        self._patch_sections()   # sections may be added late; patch at the end
        self.doc.save(str(path))
        return str(path)


# --------------------------------------------------------------------------
# Verification
# --------------------------------------------------------------------------

def find_word():
    """Locate WINWORD.EXE across Office versions. Windows only — the conversion
    below drives Word through COM, which does not exist on macOS or Linux."""
    if sys.platform != "win32":
        return None
    roots = [os.environ.get("ProgramFiles", r"C:\Program Files"),
             os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")]
    subdirs = ["root/Office16", "Office16", "root/Office15", "Office15", "Office14"]
    for root in roots:
        for sub in subdirs:
            cand = Path(root) / "Microsoft Office" / sub / "WINWORD.EXE"
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
    "Cannot render a preview: neither Microsoft Word nor LibreOffice was found. "
    "Install LibreOffice (https://www.libreoffice.org) — headless conversion is "
    "enough — or open the .docx on a machine that has Word. The document itself "
    "is already written; only the visual check needs a converter."
)

WORD_FAILED = (
    "Microsoft Word was found but failed to convert the document, and there is "
    "no LibreOffice install to fall back to. Word reported:\n{}"
)

# Pure ASCII, and paths arrive through the environment rather than being
# interpolated in. Windows PowerShell 5.1 reads a BOM-less script as the system
# ANSI codepage, so an Arabic path baked into the text comes out mangled, Word
# cannot find the file, and the conversion fails for every Arabic-named
# document — which is most of them. Environment variables are passed as
# Unicode, so the script never has to carry a non-ASCII character.
WORD_SCRIPT = r'''
$ErrorActionPreference = "Stop"
$src = $env:ARABIC_DOCX_SRC
$pdf = $env:ARABIC_DOCX_PDF
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($src, $false, $true)
    $doc.SaveAs([ref]$pdf, [ref]17)    # 17 = wdFormatPDF
    $doc.Close([ref]0)
} finally {
    $word.Quit()
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


def _word_to_pdf(docx_path, pdf_path):
    """Drive Word through COM. Returns None on success, else the failure text."""
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                     encoding="utf-8-sig") as fh:
        fh.write(WORD_SCRIPT)
        ps = fh.name
    env = dict(os.environ, ARABIC_DOCX_SRC=str(docx_path),
               ARABIC_DOCX_PDF=str(pdf_path))
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


def docx_to_pdf(docx_path, pdf_path=None):
    """Convert via Word or LibreOffice so the result can be rendered and viewed.

    Arabic defects do not show up in the document's own text — the characters
    are all present and correct, only their shaping and order are wrong. Only a
    render reveals them, so convert and look before delivering.

    A Word failure is carried, not swallowed. Reporting "no converter found"
    when Word is installed and merely errored sends the caller off installing
    LibreOffice for a problem that has nothing to do with LibreOffice.
    """
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve() if pdf_path else docx_path.with_suffix(".pdf")
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    word_error = None
    if find_word():
        word_error = _word_to_pdf(docx_path, pdf_path)
        if word_error is None:
            return str(pdf_path)

    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(WORD_FAILED.format(word_error) if word_error
                           else NO_CONVERTER)

    try:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                        "--outdir", str(pdf_path.parent), str(docx_path)],
                       check=True, capture_output=True)
    except subprocess.CalledProcessError as exc:
        detail = _console_text(exc.stderr) or _console_text(exc.stdout)
        raise RuntimeError(
            f"LibreOffice failed to convert {docx_path.name}: {detail}"
            + (f"\nWord was tried first and failed: {word_error}" if word_error else "")
        ) from exc
    # LibreOffice names the output after the *source* file, ignoring pdf_path.
    produced = pdf_path.parent / f"{docx_path.stem}.pdf"
    if produced != pdf_path and produced.exists():
        produced.replace(pdf_path)
    return str(pdf_path)


def preview(docx_path, out_dir=None, scale=2):
    """Convert to PDF and render every page to PNG. View these images."""
    import pypdfium2 as pdfium

    docx_path = Path(docx_path)
    out_dir = Path(out_dir) if out_dir else docx_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(docx_to_pdf(docx_path, out_dir / f"{docx_path.stem}.pdf"))
    paths = []
    for i in range(len(pdf)):
        png = out_dir / f"{docx_path.stem}_p{i + 1}.png"
        pdf[i].render(scale=scale).to_pil().save(png)
        paths.append(str(png))
    return paths
