---
name: arabic-docx
description: Generate Microsoft Word (.docx) documents in Arabic or any right-to-left script (Persian, Urdu, Hebrew) with correct paragraph direction, RTL tables, complex-script fonts, and proper mixed Arabic/Latin runs. Use this skill whenever the user asks for a Word document, report, letter, contract, proposal, CV, thesis chapter, or any .docx whose content is Arabic or mixes Arabic with Latin technical terms — even if they never say "RTL". Also use it when an existing Word file opens with columns on the wrong side, text drifting left, Latin words landing in the wrong place, bold Arabic that is not bold, a table of contents in the wrong font, or missing spaces around English terms — it can audit and repair existing files, and build reports with headings, a table of contents, numbered captions and page numbers.
---

# Arabic / RTL Word documents

## Do not reshape the text

This is the opposite of PDF generation, and getting it backwards is the most
damaging mistake available here.

| | reportlab / Pillow | Word |
|---|---|---|
| Text engine | none | full bidirectional engine |
| Expects | final visual-order glyphs | **logical-order, unshaped Unicode** |

`arabic_reshaper` and `python-bidi` are correct for PDF and wrong here. Word
shapes at render time, so pre-shaped presentation forms get shaped a *second*
time — the result looks broken, and the text can no longer be searched, copied,
spell-checked, or edited. If you find reshaping code in a .docx generator,
that is the bug.

What Word needs instead is plain text plus direction flags in the XML.
python-docx exposes almost none of them, so `scripts/arabic_docx.py` writes them
directly. **Use it rather than reinventing this** — each flag below fails
silently and independently.

## Dependencies

```bash
pip install python-docx pypdfium2
```

Writing the .docx needs nothing else. Rendering the preview needs Microsoft Word
(any version, via COM on Windows) or LibreOffice (`soffice`, any platform) —
whichever is present is found automatically.

## Building a document

Import the builder either by copying `scripts/arabic_docx.py` next to your
generation script, or by adding this skill's `scripts/` folder to `sys.path`:

```python
import pathlib, sys
sys.path.insert(0, str(pathlib.Path.home() / ".claude" / "skills" / "arabic-docx" / "scripts"))
from arabic_docx import ArabicDocx, preview, audit_docx, refresh_fields

doc = ArabicDocx(cs_font="Arial", latin_font="Arial", size=12)

doc.title("عنوان المستند")
doc.subtitle("سطر فرعي")
doc.toc()                                   # real TOC field

doc.heading("أولاً — المقدمة", page_break=True)
doc.para("نص عربي مع مصطلحات لاتينية مثل PostgreSQL و FastAPI.")
doc.heading("المتطلبات", level=2)
doc.bullets(["نقطة أولى", "نقطة ثانية"])    # real list, bullet on the right

doc.table(
    ["البند", "التفصيل"],              # first element = RIGHTMOST column
    [["قاعدة البيانات", "PostgreSQL"],
     ["الواجهة الخلفية", "FastAPI"]],
    [5, 11],                           # column widths in cm
    caption="مكونات النظام",            # «الجدول 1: …» above the table
)
doc.figure("diagram.png", caption="بنية النظام")   # «الشكل 1: …» below

doc.page_numbers(skip_first=True, start=0)   # cover unnumbered, next page = 1
doc.save("report.docx")

assert audit_docx("report.docx") == []
preview("report.docx")         # converts via Word (fields updated) → PNG per page
refresh_fields("report.docx")  # final file: TOC page numbers baked in (Word only)
```

Adjust the path if the skill lives elsewhere — nothing else in the builder
depends on where it is installed.

Methods chain. Cell values are passed in natural reading order — unlike the PDF
builder, columns are **not** reversed by hand, because `w:bidiVisual` performs
the flip inside Word.

| Method | Purpose |
|---|---|
| `title` / `subtitle` | Centred document head |
| `heading(text, level=1, page_break=False)` | `Heading N` style — feeds the TOC, navigation pane and PDF bookmarks |
| `para(text, bold=False, italic=False, align="start")` | Body paragraph |
| `bullets(items)` / `numbered(items)` | Real lists (`List Bullet` / `List Number`); `marker="–"` falls back to typed markers |
| `table(header, rows, widths_cm, caption=None, repeat_header=True, header_fill=None)` | RTL table; widths written to grid, cells and table; header row repeats on every page |
| `figure(path, caption=None, width_cm=None, alt=None)` | Centred image kept with its caption, scaled to the text width, alt text set |
| `caption(text, kind="figure"\|"table")` | Numbered caption on a SEQ field |
| `toc(title="المحتويات", levels="1-3")` | Table of contents field |
| `page_numbers(align="center", skip_first=False, start=None)` | PAGE field in every footer |
| `insert_before(anchor_text)` / `insert_at_end()` | Add content in the middle of an existing document |
| `page_break()` / `spacer(pts)` | Layout — prefer `heading(..., page_break=True)` |
| `save(path)` | Writes the file |

Module functions: `audit_docx(path)`, `fix_docx(src, dst)`,
`normalize_styles_rtl(doc)`, `refresh_fields(path)`, `preview(path)`,
`docx_to_pdf(path)`, `segment(text)`, `arabic_digits(text)`.

`arabic_digits("38")` → `"٣٨"` if Arabic-Indic numerals are wanted;
`arabic_digits(text, keep_latin_context=True)` leaves `Nuxt 3` and
`Python 3.12` alone. Keep Latin digits for IBANs, codes, URLs and version
numbers, and use one digit system per document.

## Starting from a template or editing an existing document

```python
doc = ArabicDocx.from_template("university-template.docx",
                               clear_body=True, clear_headers=True)
```

keeps the template's styles, page size and margins, empties the body without
deleting the final `w:sectPr` (which holds the page setup), and clears headers
and footers at the XML level — `Paragraph.clear()` leaves legacy PAGE fields
behind, which then show up next to the new ones.

Once the user has edited a document by hand, **never rebuild it**: their
edits are not in your script. Open it and insert in place:

```python
doc = ArabicDocx(template="report_v3.docx")
doc.insert_before("المراجع")         # exactly one matching paragraph, or it raises
doc.heading("الفصل الخامس — النتائج", page_break=True)
doc.para("…")
doc.save("report_v4.docx")
```

For a file made elsewhere, `fix_docx(src, dst)` retrofits the RTL flags onto
the existing structure: `w:bidi` on Arabic paragraphs, left/right alignment
converted to start/end, mixed runs split, `w:rtl` and the complex-script
bold/size slots filled, styles normalised. `tables=True` also adds
`w:bidiVisual` — off by default because it flips the column order of tables
whose columns were already reversed by hand.

## Audit before delivering

`audit_docx(path)` returns a list of findings, empty when clean. It checks for
what a reader would see but a text dump would not: pre-shaped presentation
forms, Arabic paragraphs without `w:bidi`, `jc=left/right` in RTL paragraphs,
mixed-script runs, bold without `w:bCs`, `w:sz` without `w:szCs`, tables
without `w:bidiVisual`, sections without `w:bidi`, theme fonts and
mismatched sizes in styles, and a TOC field with no cached entries. Run it on
generated files, and on files a user hands over to find out what to fix.

## Verification — convert and look

Call `preview()` and **view the PNGs with the Read tool**. It converts through
Word itself (falling back to LibreOffice), so what you see is what the recipient
sees. Arabic defects never appear in the document's text — the characters are
all correct, only their placement is wrong. Check:

- Is the first table column on the **right**?
- Are there spaces around Latin terms, or has `PostgreSQL و` collapsed to
  `PostgreSQLو`?
- Do paragraphs start at the right margin?
- Is Arabic rendered in the intended font and size, or has it fallen back?
- Do brackets and punctuation sit on the correct side — `(20 فأكثر)` not
  `20) فأكثر(`?
- Is every bracket pair intact — `DECIMAL(5,2)` not `(DECIMAL(5,2`?
- Are `C#`, `C++`, `.NET`, `+963` intact — not `#C`, `NET.`, `963+`?
- Is bold Arabic actually bold, at the same size as the Latin next to it?
- Does the table of contents show entries and page numbers?

`preview()` updates fields in memory before exporting (Word opens the file
read-only and never saves it), so the TOC in the preview has page numbers.
LibreOffice never updates fields: there the TOC shows the heading list the
builder wrote into it. LibreOffice also places an RTL table narrower than the
text block at the left margin; Word places it at the right. Judge table
position in Word.

`preview()` drives Word over COM through a temporary PowerShell script. That
script is kept pure ASCII and the paths are handed over in environment
variables, because Windows PowerShell 5.1 reads a BOM-less file as the system
ANSI codepage — an Arabic path written into the script arrives mangled and the
conversion fails on exactly the documents this skill exists to produce. If Word
is present and still fails, the error it reported is raised as-is; a "no
converter found" message means neither program was found, nothing else.

## What the builder sets, and why each matters

| Where | Element | Symptom if missing |
|---|---|---|
| `settings.xml` | `<w:themeFontLang w:bidi="ar-SA"/>` | Word never engages RTL layout at all — the master switch |
| `styles.xml` docDefaults | `<w:lang w:bidi>`, `<w:bidi/>`, `<w:jc w:val="start"/>` | Paragraphs added later revert to LTR |
| every `<w:p>` | `<w:bidi/>` + `<w:jc w:val="start"/>` | Indentation and alignment mirror wrongly |
| every `<w:r>` | `<w:rtl/>` (Arabic runs only) | Mixed text reorders incorrectly |
| every `<w:tbl>` | `<w:bidiVisual/>` | First column renders on the left |
| every `<w:sectPr>` | `<w:bidi/>` | Margins and page furniture stay LTR |
| every `<w:r>` | `<w:rFonts w:cs>` + `<w:szCs>` | Arabic silently falls back to a default face and size |
| bold / italic runs | `<w:bCs/>` / `<w:iCs/>` next to `<w:b/>` / `<w:i/>` | Latin is bold, the Arabic beside it is not |
| every style | theme font attributes removed, `szCs = sz`, `bCs` with `b` | TOC, captions and headings render Arabic in Calibri Light, smaller, not bold |
| every `<w:tbl>` | `tblW` + `gridCol` + `tcW`, fixed layout | Word redistributes the columns and ignores the widths |

### Alignment: `start` / `end`, never `left` / `right`

Write `w:jc` as `start`, `end`, `center` or `both`. In a bidi paragraph Word
reads `left` and `right` *relative to the paragraph direction* — MS-OE376
§2.3.1.13: "left is the right side of a right-to-left paragraph, and right is
the left side" — while other consumers read them physically. So
python-docx's `WD_ALIGN_PARAGRAPH.RIGHT` on an Arabic paragraph aligns it to
the **left** in Word. `para(..., align="right")` is translated to `start`.

### Schema order

Property elements (`w:rPr`, `w:pPr`, `w:tblPr`, `w:sectPr` …) are XML
sequences. The builder inserts every flag at its schema position — `w:bidi`
before `w:spacing`, `w:rFonts` before `w:sz`, `w:bidiVisual` before `w:tblW` —
instead of appending. Word forgives misordering; validators and other
consumers do not always.

## Mixed Arabic and Latin

A single run holding both scripts is the most common defect in generated Arabic
documents, so `segment()` splits text automatically and sets `w:rtl` only on the
Arabic parts.

Neutral characters — spaces, punctuation, brackets — carry no direction, so
which run they land in decides where they appear. The rule applied: a neutral
belongs with the Arabic side unless it sits *between two Latin tokens*.

```
"Vue.js 3 مع Nuxt 3."  → ["Vue.js 3"] [" مع "] ["Nuxt 3"] ["."]
"(20 فأكثر)"           → ["("] ["20"] [" فأكثر)"]
"C# و .NET"            → ["C#"] [" و "] [".NET"]
```

Characters written flush against a Latin token belong to it — `+963`,
`.NET`, `$5`, `@user` in front, `C#`, `C++` behind — and are pulled into the
Latin run. Sentence punctuation is not: the full stop in `مع Nuxt 3.` ends the
Arabic sentence. `%` is not either: Arabic places it left of the number.

Word and PowerPoint take a neutral's direction from its run, so that is enough
for them. LibreOffice, PDF viewers and browsers ignore run boundaries, so a
Latin token that starts or ends with such a character is also wrapped in an
invisible LEFT-TO-RIGHT EMBEDDING … POP DIRECTIONAL FORMATTING pair. Measured:
a plain run fails in LibreOffice, an LRM fails in Word, isolates (LRI…PDI)
are drawn as visible boxes by Word; LRE…PDF is correct in all three.

Attaching a boundary space to the Latin run instead is what produces
`PostgreSQLو` with the space visually swallowed — a defect that survived the
first build of this skill and was only caught by rendering the page.

### Bracket pairs are never split

That rule alone is not enough, because it can send the two halves of one pair
into runs of opposite direction. Word mirrors a bracket that sits in an RTL
run, so the far half comes back as its partner *and* at the far edge of the
Latin island:

| Text | Split as | Renders |
|---|---|---|
| `DECIMAL(5,2)` | `(` Latin, `)` Arabic | `(DECIMAL(5,2` — two opening parens |
| `Array<String>` | `<` Latin, `>` Arabic | `<Array<String` |
| `(API)`, `(1.25)` | both Arabic | correct — both mirror, positions swap, it cancels |

So `segment()` pulls a matched pair wholly into the Latin run whenever either
half resolved Latin, and leaves alone the pairs that resolved Arabic on both
sides. `()`, `[]`, `{}`, `<>`, `«»` are covered. A span containing Arabic is
skipped, so `a < b … c > d` stays two separate comparisons rather than one
enormous bracket.

The lesson generalises: **any mirrored character has to share a run with its
partner.** A pair split across a direction boundary always renders wrong.

## Diagrams and images with Arabic labels

Text baked into an image never gets Word's bidi engine. Draw it with an
engine that shapes:

- **SVG**: write logical Unicode with `direction="rtl"`
  `unicode-bidi="plaintext"` on Arabic labels and `direction="ltr"` on
  identifiers and code, then rasterise with a browser or `rsvg-convert`.
- **Pillow**: `draw_text()` from `arabic-pdf` — it uses libraqm when Pillow
  has it (logical text, no reshaping) and falls back to reshaping otherwise.

Arrows are not mirrored by bidi: a flow that runs right to left needs `←`.
Size label fonts for the printed result, not the canvas — a label meant to
read at 9 pt in an image placed 15 cm wide on a 2000 px canvas needs
`9 × 2000 / (15 / 2.54 × 72)` ≈ 42 px.

## Fonts

Word keeps three font slots per run. Arabic is drawn from the **complex-script**
slot, so setting only the Latin slots leaves Arabic on a default face:

```xml
<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Traditional Arabic"/>
<w:szCs w:val="24"/>
```

Good complex-script faces: Arial and Tahoma (safe everywhere), Traditional
Arabic and Sakkal Majalla (formal documents), Amiri (academic). Pass via
`ArabicDocx(cs_font=...)`.

## Typography conventions

Use Arabic punctuation in Arabic text — `،` (U+060C), `؛` (U+061B), `؟`
(U+061F), and `«…»` for quotes. Latin `,;?` inside an Arabic run reads as
careless and can sit on the wrong side.

## What makes generated Arabic read as generated

Not the vocabulary — the additions. Each habit below tells the reader something
about how the document was written rather than about its subject, and a reader
editing by hand deletes all of them. Write the sentence, not the account of
writing it.

| Habit | Instead of | Write |
|---|---|---|
| Narrating the act of writing | «الشروط ثلاثة، نذكرها صراحة فيما يلي» | «الشروط ثلاثة:» |
| Pointing at another section | «…ونعود إلى هذه النقطة لاحقاً» | احذف العبارة |
| Version or status labels in a title | «الخطة (النسخة الثانية — معتمدة)» | «الخطة» |
| Arguing against an option nobody raised | «نحفظ المسار فقط. تخزين الملف كاملاً يضخّم الحجم بلا فائدة.» | «نحفظ المسار فقط.» |
| Justifying a choice by what the brief omitted | «لم يرد ذلك في الطلب، غير أن طبيعة العمل تفرضه.» | «طبيعة العمل تفرض ذلك.» |
| First-person singular | «وأستطيع لاحقاً تحليل الحالات» | «ويمكن لاحقاً تحليل الحالات» |
| Restating a fact the document already gave | تكرار المعلومة في قسم آخر | احذف التكرار |

Formal Arabic prefers the impersonal or the plural over «أنا»; the singular
reads as a note to oneself rather than a document.

## Related

Use `arabic-pdf` for PDF output and `arabic-pptx` for slides. The three solve
the same problem with **opposite** techniques — never copy an approach between
them without re-reading why.
