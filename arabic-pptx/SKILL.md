---
name: arabic-pptx
description: Generate PowerPoint (.pptx) presentations in Arabic or any right-to-left script (Persian, Urdu, Hebrew) with correct paragraph direction, RTL tables, complex-script fonts, and properly spaced mixed Arabic/Latin text. Use this skill whenever the user asks for slides, a deck, a presentation, a pitch, a defense talk, or any .pptx whose content is Arabic or mixes Arabic with Latin technical terms — even if they never say "RTL". Also use it when an existing deck shows Arabic drifting to the left, table columns in the wrong order, version numbers like "Nuxt 3" reversed, or spaces missing around English words — it can audit and repair existing decks and lays out numbered steps right to left.
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
from arabic_pptx import ArabicPptx, preview, audit_pptx

deck = ArabicPptx(cs_font="Arial", base_size=24)

deck.title_slide("عنوان العرض", "سطر فرعي")
deck.bullets_slide("المحاور", [
    "نقطة عربية تتضمن مصطلحات لاتينية مثل PostgreSQL و FastAPI.",
    "أرقام الإصدارات مثل Nuxt 3 تبقى بترتيبها الصحيح.",
])
deck.steps_slide("مراحل العمل", [
    ("جمع البيانات", "السجلات والوصفات"),
    ("التحليل", "مطابقة القواعد"),
    ("التنبيه", "إشعار الطبيب"),
])                                  # 01 on the right, joined by ← arrows
deck.table_slide("جدول",
    ["البند", "التفصيل"],           # first element = RIGHTMOST column
    [["الواجهة الخلفية", "FastAPI"],
     ["قاعدة البيانات", "PostgreSQL"]],
    [10, 12])                       # column widths in cm
deck.image_slide("البنية", "diagram.png", caption="الشكل 1: بنية النظام")

deck.save("deck.pptx")
assert audit_pptx("deck.pptx") == []
preview("deck.pptx")   # converts via PowerPoint and renders PNG per slide
```

Adjust the path if the skill lives elsewhere — nothing else in the builder
depends on where it is installed.

| Method | Purpose |
|---|---|
| `title_slide(title, subtitle)` | Centred opening slide |
| `bullets_slide(title, items, marker="•")` | Heading plus a real bulleted list (`a:buChar`, hanging indent on the right) |
| `text_slide(title, paragraphs)` | Heading plus body paragraphs |
| `steps_slide(title, steps)` | 2–5 numbered cards, 01 on the right, joined by ← arrows |
| `table_slide(title, header, rows, widths_cm)` | Heading plus RTL table |
| `image_slide(title, path, caption=None)` | Heading plus a picture fitted to the body area |
| `save(path)` | Writes the file |

Module functions: `audit_pptx(path)`, `fix_pptx(src, dst)`,
`rtl_positions(n, left, width, item_width, gap)`, `reading_order(shapes)`,
`make_rtl_defaults(prs)`, `preview(path)`, `segment(text)`, `arabic_digits(text)`.
Optional animation helpers are described under "Optional: if the deck is
animated".

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
| `<a:endParaRPr>` | `lang` + `a:cs` | Text the presenter types at the end of a line comes in as English |
| `<a:pPr>` | `marL` + `indent` + `a:buChar` | Bullets typed as text: wrapped lines start under the bullet |
| `p:defaultTextStyle`, master `p:txStyles` | `rtl="1"`, `algn="r"` | New text boxes added in PowerPoint start on the left |

Note `<a:cs>` is a **child element**, not an attribute — setting `font.name`
only fills the Latin slot and leaves Arabic on a fallback face.

## Layout runs right to left too

Direction flags fix text inside a box; they do nothing for the *order of the
boxes*. Numbered cards, process steps, timelines and phase diagrams laid out
with `x = left + i * step` read backwards in Arabic — 01 on the left — and no
text setting can repair that. It is the most common defect in otherwise
correct Arabic decks.

- `rtl_positions(n, left, width, item_width, gap)` returns x offsets with item
  0 on the right; `steps_slide()` uses it.
- `reading_order(shapes)` sorts existing shapes top band first, then right to
  left — the order to reveal them in.
- Arrows are characters, and bidi does not mirror them: a right-to-left flow
  needs `←`. A `→` typed in an RTL deck still points right, i.e. backwards.

## Readability

The audit flags text below 18 pt. That is a floor, not a target: in a hall,
titles read at 40 pt and up and list text at 24 pt and up. `base_size=24` is
a better start for a talk than the default 18. Long titles shrink to fit one
line (down to `base_size + 4`) before they wrap. Units: `a:rPr@sz` is in
hundredths of a point (`sz="2400"` = 24 pt); some toolkits take pixels
(1 px = 0.75 pt) — check before trusting a number.

## Fixing and auditing existing decks

`audit_pptx(path)` lists what a viewer would see wrong: Arabic paragraphs that
are not RTL (resolved through the inheritance chain — slide, shape list
style, master text styles, presentation defaults), Latin-only paragraphs
marked RTL whose edge punctuation jumps, mixed-script runs, Arabic runs
tagged `en-*`, tables with Arabic but no `rtl="1"`, pre-shaped presentation
forms, and small text.

`fix_pptx(src, dst)` repairs those in place, keeping positions and formatting:
`rtl="1"` on Arabic paragraphs (alignment changed only where it was left or
unset, so centred titles stay centred), mixed runs split with a `lang` each,
`a:cs` added where a Latin font is named, and the presentation and master
defaults made RTL so text typed later starts on the right. `tables=True`
also flips tables — off by default, since it reverses column order. Layout
order (cards, steps) is not changed: check those slides by eye.

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
- `C#`, `C++`, `.NET`, `+963` intact — not `#C`, `NET.`, `963+`
- Arabic in the intended font, not a fallback
- Numbered items and steps start on the **right**; arrows point left

LibreOffice ignores `rtl="1"` on tables, so its previews show the first column
on the left even when PowerPoint shows it correctly on the right. Judge table
column order in PowerPoint.

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
"C# و .NET"            → ["C#"] [" و "] [".NET"]
```

Characters written flush against a Latin token — `+963`, `.NET`, `$5`,
`C#`, `C++` — belong to it and go into its run; sentence punctuation and `%`
do not. Such tokens are also wrapped in invisible LRE…PDF marks, because
LibreOffice and PDF viewers ignore run boundaries (see arabic-docx for the
measurements).

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

## Optional: if the deck is animated

Nothing is animated by default, and nothing in this section is needed for a
correct Arabic deck. Animation style is a matter of taste; what follows are
the parts that are about direction, for decks that do animate.

- **Directional effects read backwards.** Fly-in, wipe, push and cover enter
  from the left by default. In an RTL deck they have to be mirrored (enter
  from the right) or replaced with a direction-neutral effect such as fade.
- **Builds follow reading order.** Items revealed one by one should appear
  right to left, top row first. `reading_order(shapes)` sorts existing shapes
  that way.

The library ships one ready-made, direction-neutral option, opt-in:

| Call | Effect |
|---|---|
| `ArabicPptx(transition="fade")` / `add_transition(slide, "fade")` | Fade between slides (`fade`, `dissolve` or `cut` only) |
| `bullets_slide(..., reveal=True)` | One bullet per click |
| `steps_slide(..., reveal=True)` | One step per click, in reading order |
| `set_build(slide, steps, dur_ms=350, stagger_ms=60)` | Click-by-click fade builds for any shapes; a step is a list of shapes or `(shape, paragraph_index)` |

`set_build` writes the structure PowerPoint itself saves for "Fade — On Click
/ With Previous", after `p:transition` as the schema requires; PowerPoint
opens it without repair and lists the effects in its Animation Pane. It
replaces any animation already on the slide, so do not call it on a slide
whose animations the user designed.

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
