---
name: arabic-pptx
description: Generate PowerPoint (.pptx) presentations in Arabic or any right-to-left script (Persian, Urdu, Hebrew) with correct paragraph direction, RTL tables, complex-script fonts, and properly spaced mixed Arabic/Latin text. Use this skill whenever the user asks for slides, a deck, a presentation, a pitch, a defense talk, or any .pptx whose content is Arabic or mixes Arabic with Latin technical terms — even if they never say "RTL". Also use it when an existing deck shows Arabic drifting to the left, table columns in the wrong order, version numbers like "Nuxt 3" reversed, or spaces missing around English words.
---

# Arabic / RTL PowerPoint decks

## Do not reshape the text

PowerPoint has a full bidirectional text engine, so it must be given
**logical-order, unshaped Unicode**. `arabic_reshaper` and `python-bidi` are
correct for PDF output and wrong here — PowerPoint would shape the presentation
forms a second time, producing broken glyphs and text that cannot be searched or
edited. If you find reshaping code in a .pptx generator, that is the bug.

Direction is expressed in DrawingML, which is a different vocabulary from Word's
WordprocessingML. `python-pptx` exposes none of it, so `scripts/arabic_pptx.py`
writes it to the XML directly. **Use it rather than reinventing this.**

## Dependencies

```bash
pip install python-pptx pypdfium2
```

Writing the .pptx needs nothing else. Rendering the preview needs Microsoft
PowerPoint (any version, via COM on Windows) or LibreOffice (`soffice`, any
platform) — whichever is present is found automatically.

## Building a deck

Import the builder either by copying `scripts/arabic_pptx.py` next to your
generation script, or by adding this skill's `scripts/` folder to `sys.path`:

```python
import pathlib, sys
sys.path.insert(0, str(pathlib.Path.home() / ".claude" / "skills" / "arabic-pptx" / "scripts"))
from arabic_pptx import ArabicPptx, preview

deck = ArabicPptx(cs_font="Arial", base_size=18)

deck.title_slide("عنوان العرض", "سطر فرعي")
deck.bullets_slide("المحاور", [
    "نقطة عربية تتضمن مصطلحات لاتينية مثل PostgreSQL و FastAPI.",
    "أرقام الإصدارات مثل Nuxt 3 تبقى بترتيبها الصحيح.",
])
deck.table_slide("جدول",
    ["البند", "التفصيل"],           # first element = RIGHTMOST column
    [["الواجهة الخلفية", "FastAPI"],
     ["قاعدة البيانات", "PostgreSQL"]],
    [10, 12])                       # column widths in cm

deck.save("deck.pptx")
preview("deck.pptx")   # converts via PowerPoint and renders PNG per slide
```

Adjust the path if the skill lives elsewhere — nothing else in the builder
depends on where it is installed.

| Method | Purpose |
|---|---|
| `title_slide(title, subtitle)` | Centred opening slide |
| `bullets_slide(title, items, marker="•")` | Heading plus bulleted list |
| `text_slide(title, paragraphs)` | Heading plus body paragraphs |
| `table_slide(title, header, rows, widths_cm)` | Heading plus RTL table |
| `save(path)` | Writes the file |

Columns are passed in natural reading order — `rtl="1"` on `<a:tblPr>` performs
the flip inside PowerPoint, so no manual reversal (unlike the PDF builder).

Slides use explicit text boxes rather than layout placeholders, because
placeholder geometry varies between templates and predictable RTL output matters
more here than theme integration.

## `lang` is the switch that decides everything

The single most important attribute, and the one nothing else compensates for.
PowerPoint uses `lang` on `<a:rPr>` to choose shaping and bidi behaviour for a
run. `python-pptx` never sets it. Measured on the same sentence
`"منصة ويب مبنية على PostgreSQL و FastAPI مع Nuxt 3."`:

| Setup | Result |
|---|---|
| No `lang` anywhere | `PostgreSQLو` — every boundary space collapses |
| `lang="ar-SA"` on one run holding everything | `3Nuxt` — Latin numbers reorder |
| **`lang` per run** — `ar-SA` on Arabic, `en-US` on Latin | **correct** |

This is why the text is split into runs at all: not for fonts, but so each
fragment can carry the right language. It is the DrawingML counterpart of
`w:lang w:bidi` in Word, and it fails just as silently — the deck opens, the
text is all present, and it simply reads wrong.

For other scripts pass `ArabicPptx(rtl_lang="fa-IR")` — also `ur-PK`, `he-IL`.

## Everything the builder sets

| Level | Attribute | Symptom if missing |
|---|---|---|
| `<a:rPr>` | `lang` per run | Spaces collapse, numbers reorder — see above |
| `<a:pPr>` | `rtl="1"` | Mixed Arabic/Latin reorders wrongly |
| `<a:pPr>` | `algn="r"` | Text hugs the left edge of its box |
| `<a:rPr>` | child `<a:cs typeface="…"/>` | Arabic ignores the font entirely |
| `<a:tblPr>` | `rtl="1"` | First column renders on the left |
| `<a:t>` | `xml:space="preserve"` | OOXML strips leading/trailing whitespace |

Note `<a:cs>` is a **child element**, not an attribute — setting `font.name`
only fills the Latin slot and leaves Arabic on a fallback face.

## Verification — convert and look

Call `preview()` and **view the PNGs with the Read tool**. It converts through
PowerPoint itself (falling back to LibreOffice), so what you see is what the
audience sees. Zoom into a line with mixed text if anything looks tight —
several of the defects above are a few pixels wide and easy to miss at full-slide
scale. Check:

- Spaces around Latin terms: `PostgreSQL و` not `PostgreSQLو`
- Version numbers intact: `Nuxt 3` not `3Nuxt`
- First table column on the **right**
- Brackets on the correct side: `(20 فأكثر)` not `20) فأكثر(`
- Bracket pairs intact: `DECIMAL(5,2)` not `(DECIMAL(5,2`
- Arabic in the intended font, not a fallback

None of this shows up in the file's text — the characters are all correct, only
their placement is wrong. Only a render reveals it.

`preview()` drives PowerPoint over COM through a temporary PowerShell script.
That script is kept pure ASCII and the paths are handed over in environment
variables, because Windows PowerShell 5.1 reads a BOM-less file as the system
ANSI codepage — an Arabic path written into the script arrives mangled and the
conversion fails on exactly the decks this skill exists to produce. If
PowerPoint is present and still fails, the error it reported is raised as-is; a
"no converter found" message means neither program was found, nothing else.

## Mixed Arabic and Latin

`segment()` splits text and assigns each fragment a language. Neutral characters
— spaces, punctuation, brackets — carry no direction of their own, so they go
with the Arabic side unless they sit *between two Latin tokens*:

```
"Vue.js 3 مع Nuxt 3."  → ["Vue.js 3"] [" مع "] ["Nuxt 3"] ["."]
"(20 فأكثر)"           → ["("] ["20"] [" فأكثر)"]
```

### Bracket pairs are never split

That rule alone can send the two halves of one pair into runs of opposite
direction. PowerPoint mirrors a bracket sitting in an RTL run, so the far half
comes back as its partner *and* at the far edge of the Latin island:

| Text | Split as | Renders |
|---|---|---|
| `DECIMAL(5,2)` | `(` Latin, `)` Arabic | `(DECIMAL(5,2` — two opening parens |
| `Array<String>` | `<` Latin, `>` Arabic | `<Array<String` |
| `(API)`, `(1.25)` | both Arabic | correct — both mirror, positions swap, it cancels |

So `segment()` pulls a matched pair wholly into the Latin run whenever either
half resolved Latin, and leaves alone pairs that resolved Arabic on both sides.
`()`, `[]`, `{}`, `<>`, `«»` are covered; a span containing Arabic is skipped so
`a < b … c > d` stays two comparisons rather than one enormous bracket.

**Any mirrored character has to share a run with its partner.** A pair split
across a direction boundary always renders wrong.

## Typography

Use Arabic punctuation in Arabic text — `،` `؛` `؟` and `«…»`. Keep Latin digits
for version numbers, codes and URLs; `arabic_digits("38")` → `"٣٨"` is available
when Arabic-Indic numerals suit the audience.

## What makes generated Arabic read as generated

Not the vocabulary — the additions. Each habit below tells the audience
something about how the deck was written rather than about its subject, and it
is the first thing a presenter deletes. Slides punish it doubly: there is no
room.

| Habit | Instead of | Write |
|---|---|---|
| Narrating the act of writing | «الشروط ثلاثة، نذكرها صراحة فيما يلي» | «الشروط ثلاثة:» |
| Pointing at another slide | «…ونعود إلى هذه النقطة لاحقاً» | احذف العبارة |
| Version or status labels in a title | «الخطة (النسخة الثانية — معتمدة)» | «الخطة» |
| Arguing against an option nobody raised | «نحفظ المسار فقط. تخزين الملف كاملاً يضخّم الحجم بلا فائدة.» | «نحفظ المسار فقط.» |
| Justifying a choice by what the brief omitted | «لم يرد ذلك في الطلب، غير أن طبيعة العمل تفرضه.» | «طبيعة العمل تفرض ذلك.» |
| First-person singular | «وأستطيع لاحقاً تحليل الحالات» | «ويمكن لاحقاً تحليل الحالات» |
| Restating a fact an earlier slide gave | تكرار المعلومة | احذف التكرار |

Formal Arabic prefers the impersonal or the plural over «أنا»; the singular
reads as a note to oneself rather than a presentation.

## Related

Use `arabic-pdf` for PDF and `arabic-docx` for Word. All three solve the same
problem with **different, non-transferable** techniques — PDF needs pre-shaped
visual-order text, while Word and PowerPoint need raw logical text plus direction
flags. Never copy an approach between them without re-reading why.
