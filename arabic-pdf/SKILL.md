---
name: arabic-pdf
description: Generate PDF documents containing Arabic (or any right-to-left script — Persian, Urdu, Hebrew) using reportlab, with correct glyph shaping, bidirectional ordering, line wrapping, and RTL tables. Use this skill whenever the user asks for a PDF, report, summary, proposal, thesis chapter, invoice, certificate, or any printable document whose content is in Arabic or mixes Arabic with Latin technical terms — even if they don't mention "RTL" or name a library. Also use it when an existing Arabic PDF comes out with reversed lines, shuffled words, disconnected letters, or black boxes instead of glyphs.
---

# Arabic / RTL PDF generation

Naive reportlab code produces Arabic that is subtly and thoroughly wrong: letters
appear disconnected, words within a line run backwards, and lines within a
paragraph swap places. None of this is visible in extracted text — only in a
render. This skill bundles a builder that has these problems already solved.

**Use `scripts/arabic_pdf.py` rather than writing reshaping logic from scratch.**
Every failure mode below was found by hitting it in a real document.

## Dependencies

```bash
pip install reportlab arabic-reshaper python-bidi pypdfium2
```

`pypdfium2` is only for previewing, but previewing is not optional — see
Verification below.

## Building a document

Import the builder either by copying `scripts/arabic_pdf.py` next to your
generation script, or by adding this skill's `scripts/` folder to `sys.path`:

```python
import pathlib, sys
sys.path.insert(0, str(pathlib.Path.home() / ".claude" / "skills" / "arabic-pdf" / "scripts"))
from arabic_pdf import ArabicPDF, preview

pdf = ArabicPDF("report.pdf", doc_title="عنوان المستند", page_numbers=True)

pdf.title("عنوان المستند")
pdf.subtitle("سطر فرعي")

pdf.heading("أولاً — المقدمة")
pdf.para("فقرة عربية تتضمن مصطلحات لاتينية مثل PostgreSQL و FastAPI.")

pdf.heading("ثانياً — جدول")
pdf.table(
    ["البند", "التفصيل"],                     # rightmost column first
    [["الواجهة الخلفية", "Python — FastAPI"],
     ["قاعدة البيانات", "PostgreSQL — SQLAlchemy"]],
    [5.0, 11.0],                              # column widths in cm
)

pdf.bullets(["نقطة أولى", "نقطة ثانية"])
pdf.save()

preview("report.pdf")     # one PNG per page
```

Adjust the path if the skill lives elsewhere — nothing else in the builder
depends on where it is installed.

Methods chain, so `pdf.heading(...).para(...)` also works.

| Method | Purpose |
|---|---|
| `title` / `subtitle` | Centred document head |
| `heading` | Right-aligned bold section head |
| `para` | Body paragraph |
| `bullets(items, marker="–")` | One right-aligned line per item |
| `table(header, rows, widths_cm)` | RTL table; natural reading order in, reversed columns out |
| `spacer(pts)` / `page_break()` | Layout control |
| `save()` | Writes the file, returns the path |

Cell and paragraph text may contain `\n` to force a line break.

## Verification — always render and look

After `save()`, call `preview()` and **view the PNGs with the Read tool**. Text
extraction cannot catch RTL defects, because the characters are all present and
correct; only their *order and shape* are wrong. Check specifically:

- Does each paragraph's first word sit at the **top right**?
- Do Latin terms (`PostgreSQL`, `FastAPI`) sit where they belong mid-sentence,
  rather than jumping to the far left of the line?
- Are the letters joined, or standing apart like `ا ل س ل ا م`?
- Any solid black boxes? The font lacks those glyphs — pick another.
- Do brackets face the right way — `(20 فأكثر)` not `)20 فأكثر(`?
- Are pairs inside Latin terms intact — `Array<String>`, `DECIMAL(5,2)`?
- Did a single table row get orphaned onto the next page?

Fixing a defect you never looked for costs far more later, when the document is
already with its recipient.

## How the correctness is achieved

Useful when adapting the builder or debugging an odd document.

1. **Shaping** — `arabic_reshaper` maps characters to positional presentation
   forms so letters connect.
2. **Bidi with an explicit base direction** — `get_display(..., base_dir="R")`.
   Without `base_dir`, direction is inferred from the first strong character, so
   a line starting with a Latin word is laid out left-to-right and the Arabic
   after it lands wrong. This is the single most common cause of "mostly right
   but some lines are scrambled".
3. **Manual line wrapping** — reshaped text is in *visual* order. If reportlab
   wraps it, the words it pushes to the next line are the logically *first*
   ones, so lines come out shuffled. The builder wraps first, reshapes each line
   separately, and joins with `<br/>` so reportlab never re-wraps. An 8pt safety
   margin covers measurement differences; a line that overflows by one point is
   enough to reverse it.
4. **Mirroring** — `get_display` reorders but never applies bidi rule L4, which
   says a mirrored character laid out right-to-left is drawn as its partner.
   Without it `(20 فأكثر)` comes back with the two parens in exchanged
   positions and their original glyphs, and prints `)20 فأكثر(`. `prepare()`
   substitutes them first, so the reorder moves the correct glyph.
5. **Latin islands** — `<` and `>` are mirrored but are *not* Unicode paired
   brackets, so the rule that keeps `DECIMAL(5,2)` together does not cover
   them; the `>` of `Array<String>` takes the direction of the Arabic after it
   and is thrown to the far side. Each Latin span is wrapped in `LRI…PDI` and
   the isolates are stripped again after reordering, so nothing invisible
   reaches the PDF.
6. **Escaping after shaping** — reportlab needs `&`, `<`, `>` escaped, but
   escaping *first* sends `&lt;` through the bidi pass as four characters and
   its trailing `;` is ejected to the far side of the line. Shape, then escape.
7. **Reversed table columns** — reportlab lays columns left-to-right, so the
   builder reverses them, letting you pass columns in natural reading order.
8. **`KeepTogether` on tables** — prevents a lone row stranded on a new page.

Points 4–6 are the same class of defect as the split bracket runs described in
`arabic-docx`, reached from the opposite direction: there a pair must not be
split across runs, here a pair must not be half-mirrored. **Any mirrored
character has to be resolved together with its partner.**

## Fonts

`register_fonts()` searches for an Arabic-capable font: Arial then Tahoma on
Windows, Arial on macOS, Noto Naskh Arabic then DejaVu Sans on Linux. To use a
specific face (Amiri and Cairo suit formal documents well):

```python
from arabic_pdf import register_fonts
register_fonts("/path/Amiri-Regular.ttf", "/path/Amiri-Bold.ttf")   # before ArabicPDF(...)
```

Reportlab's built-in fonts have no Arabic coverage, so a TTF must be registered —
this is why a missing font shows up as black boxes rather than an error.

## Never carry this approach into Word or PowerPoint

This is the most important boundary in the whole family of RTL skills.

| | reportlab / Pillow | Word / PowerPoint |
|---|---|---|
| Text engine | **none** | full bidirectional engine |
| Expects | final visual-order glyphs | **raw logical-order Unicode** |
| Technique | reshape + bidi + manual wrap | direction flags in the XML |

Pre-shaping is *required* here and *destructive* there: Word and PowerPoint
shape at render time, so handing them presentation forms makes them shape a
second time — broken glyphs, and text that can no longer be searched, copied, or
spell-checked. The two approaches are opposites, not variations.

Use `arabic-docx` for `.docx` and `arabic-pptx` for `.pptx`. Do not port code
between the three without re-reading why they differ.

## Notes

- `arabic_digits("38")` → `"٣٨"` when Arabic-Indic numerals are wanted. Purely
  cosmetic; Western digits are perfectly normal in Arabic technical writing.
- Keep formatting restrained by default — black text, thin rules. Academic and
  official Arabic documents read as more credible that way, and users asking for
  "simple" mean it.
- The same machinery handles Persian, Urdu, and Hebrew; only the font needs to
  cover the script.

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
