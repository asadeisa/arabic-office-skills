# Arabic Office Skills

Three agent skills for producing **PDF, Word, and PowerPoint files in Arabic**
— and any other right-to-left script — that actually render correctly.

Arabic output breaks in ways that are invisible until someone opens the file:
letters stand apart instead of joining, lines within a paragraph swap places,
table columns end up mirrored, and the space around an English term silently
disappears. None of it shows up in extracted text, because the characters are
all present and correct. Only their shape and placement are wrong.

![The same paragraph rendered by plain reportlab and by arabic-pdf](docs/before-after.png)

Same text, same font, same call. On the left the letters stand apart, the title
reads back to front, and `Nuxt 3` has jumped out of the sentence — that is what
plain reportlab produces, not an exaggeration of it.

Each skill fixes this for one format, and each bundles a working Python library
plus a `preview()` that renders the result to PNG so it can be checked before
delivery.

| Skill | Format | Library |
|---|---|---|
| [`arabic-pdf`](arabic-pdf/) | `.pdf` | reportlab |
| [`arabic-docx`](arabic-docx/) | `.docx` | python-docx |
| [`arabic-pptx`](arabic-pptx/) | `.pptx` | python-pptx |

## The one thing worth knowing

These three formats need **opposite** techniques, and the most common way to
break Arabic output is to carry the right fix into the wrong format.

| | reportlab / Pillow | Word / PowerPoint |
|---|---|---|
| Text engine | **none** | full bidirectional engine |
| Expects | final visual-order glyphs | **raw logical-order Unicode** |
| Technique | reshape + bidi + manual line wrapping | direction flags in the OOXML |

`arabic_reshaper` + `python-bidi` are mandatory for PDF and destructive for
Word and PowerPoint — those shape at render time, so pre-shaped text gets shaped
twice, leaving broken glyphs and text that can no longer be searched, copied, or
spell-checked.

Word and PowerPoint instead need direction flags the Python libraries do not
expose, so the bundled scripts write them into the XML directly:

- **Word** — `w:themeFontLang`, `w:bidi` on paragraphs and sections, `w:rtl` on
  runs, `w:bidiVisual` on tables, and the complex-script font slot `w:cs`/`w:szCs`
- **PowerPoint** — `lang` per run, `rtl="1"` and `algn="r"` on paragraphs,
  `rtl="1"` on tables, `<a:cs>` for fonts, `xml:space="preserve"` on text

Each flag fails silently and independently. Missing `lang` on a PowerPoint run
turns `PostgreSQL و` into `PostgreSQLو`; putting it on the whole paragraph
instead turns `Nuxt 3` into `3Nuxt`.

## Install

```bash
git clone https://github.com/asadeisa/arabic-office-skills.git
cp -r arabic-office-skills/arabic-* ~/.claude/skills/
```

On Windows, copy the three folders into `C:\Users\<you>\.claude\skills\`.

Dependencies, by skill:

```bash
pip install reportlab arabic-reshaper python-bidi pypdfium2   # arabic-pdf
pip install python-docx pypdfium2                             # arabic-docx
pip install python-pptx pypdfium2                             # arabic-pptx
```

Rendering `.docx` and `.pptx` previews needs Microsoft Office or LibreOffice.
Generating the files needs neither.

## Using them

With Claude Code the skills trigger on their own — ask for an Arabic report,
document, or deck. With any other agent or on their own, the libraries are plain
Python:

```python
from arabic_pdf import ArabicPDF, preview

pdf = ArabicPDF("تقرير.pdf")
pdf.title("عنوان المستند")
pdf.heading("أولاً — المقدمة")
pdf.para("نص عربي مع مصطلحات مثل PostgreSQL و FastAPI.")
pdf.table(["الطبقة", "التقنية"], [["الواجهة", "Nuxt 3"]], [5, 11])
pdf.save()
preview("تقرير.pdf")
```

`ArabicDocx` and `ArabicPptx` follow the same shape. Each `SKILL.md` documents
its own API, the XML it sets, and what to look for when checking the output.

## Scripts

Persian, Urdu, and Hebrew work the same way. For PDF only the font needs to
cover the script; for PowerPoint pass the language too:

```python
ArabicPptx(rtl_lang="fa-IR")   # also ur-PK, he-IL
```

## License

MIT — see [LICENSE](LICENSE).
