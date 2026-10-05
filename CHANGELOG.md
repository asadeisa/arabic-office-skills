# Changelog

## Unreleased

Lessons from a full Arabic graduation report (Word, ~60 pages) and its
defense deck (PowerPoint, 24 slides), both produced by AI agents over several
failed and then successful iterations, turned into generic functions.

### arabic-docx
- Bold, italic and size now reach Arabic: `w:bCs`, `w:iCs` and `w:szCs` are
  written next to their Latin counterparts. Previously bold Arabic stayed
  regular weight.
- `normalize_styles_rtl()` runs on every document: theme font references are
  removed from styles (they override explicit fonts), `szCs` matches `sz`,
  `bCs` matches `b`, and `w:bidi w:val="0"` is dropped.
- Alignment is only ever written as `start`/`end`/`center`/`both`;
  `align="right"` maps to `start`. Word reads `left`/`right` relative to the
  paragraph direction in RTL paragraphs.
- XML properties are inserted at their schema position instead of appended.
- New: `heading(level=, page_break=)` on `Heading N` styles, `toc()`,
  `caption()`, `figure()`, `numbered()`, real `bullets()`,
  `page_numbers(skip_first=, start=)`, `ArabicDocx.from_template()`,
  `insert_before()` / `insert_at_end()`.
- `table()` writes widths to `tblW`, `gridCol` and `tcW` with a fixed layout,
  repeats the header row, never splits rows, falls back to explicit borders
  when the template has no "Table Grid", and takes `caption=` and
  `header_fill=`.
- New: `audit_docx()`, `fix_docx()`, `refresh_fields()`. `preview()` updates
  fields before exporting so the TOC shows page numbers.

### arabic-pptx
- Real bullets (`a:buChar` with a hanging indent) instead of typed markers.
- `a:endParaRPr` carries `lang` and `a:cs`; presentation and master text
  styles are made RTL, so text added later in PowerPoint starts on the right.
- The blank layout is found by name or by placeholder count, not by index 6.
- Long titles shrink to fit before they wrap.
- Latin-only table cells are LTR (still right-aligned).
- New: `steps_slide()`, `image_slide()`, `rtl_positions()`,
  `reading_order()`, `audit_pptx()`, `fix_pptx()`, `make_rtl_defaults()`.
- Optional, off by default: `add_transition()`, `set_build()` and the
  `transition=` / `reveal=` options — direction-neutral fade only, built in
  RTL reading order. Documented as optional; animation is a style choice.

### all
- `segment()` keeps flush affixes with their Latin token (`C#`, `C++`,
  `.NET`, `+963`, `$5`), and such tokens are wrapped in LRE…PDF so they render
  correctly in LibreOffice and PDF viewers too — measured in Word,
  PowerPoint and LibreOffice.
- `arabic_digits(text, keep_latin_context=True)` leaves `Nuxt 3` alone.
- arabic-pdf: `draw_text()` for Arabic in Pillow images (uses libraqm when
  available), `digits=` for page numbers, and line wrapping no longer breaks
  no-break spaces.
- Unit tests, a render check script, and CI.
