---
name: arabic-docx
description: Generate Microsoft Word (.docx) documents in Arabic or any right-to-left script (Persian, Urdu, Hebrew) with correct paragraph direction, RTL tables, complex-script fonts, and proper mixed Arabic/Latin runs. Use this skill whenever the user asks for a Word document, report, letter, contract, proposal, CV, thesis chapter, or any .docx whose content is Arabic or mixes Arabic with Latin technical terms — even if they never say "RTL". Also use it when an existing Word file opens with columns on the wrong side, text drifting left, Latin words landing in the wrong place, or missing spaces around English terms.
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
from arabic_docx import ArabicDocx, preview

doc = ArabicDocx(cs_font="Arial", latin_font="Arial", size=12)

doc.title("عنوان المستند")
doc.subtitle("سطر فرعي")

doc.heading("أولاً — المقدمة")
doc.para("نص عربي مع مصطلحات لاتينية مثل PostgreSQL و FastAPI.")

doc.table(
    ["البند", "التفصيل"],              # first element = RIGHTMOST column
    [["قاعدة البيانات", "PostgreSQL"],
     ["الواجهة الخلفية", "FastAPI"]],
    [5, 11],                           # column widths in cm
)

doc.bullets(["نقطة أولى", "نقطة ثانية"])
doc.save("report.docx")

preview("report.docx")   # converts via Word and renders PNG per page
```

Adjust the path if the skill lives elsewhere — nothing else in the builder
depends on where it is installed.

Methods chain. Cell values are passed in natural reading order — unlike the PDF
builder, columns are **not** reversed by hand, because `w:bidiVisual` performs
the flip inside Word.

| Method | Purpose |
|---|---|
| `title` / `subtitle` | Centred document head |
| `heading` | Right-aligned bold section head |
| `para(text, bold=False)` | Body paragraph |
| `bullets(items, marker="–")` | One paragraph per item |
| `table(header, rows, widths_cm)` | RTL table |
| `page_break()` / `spacer(pts)` | Layout |
| `save(path)` | Writes the file |

`arabic_digits("38")` → `"٣٨"` if Arabic-Indic numerals are wanted. Keep Latin
digits for IBANs, codes, URLs and version numbers.

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

Prefer `w:jc w:val="start"` over `right`: `start` is direction-relative, so it
stays correct if a paragraph turns out to be LTR. Hardcoded `right` breaks
mixed documents.

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
```

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
