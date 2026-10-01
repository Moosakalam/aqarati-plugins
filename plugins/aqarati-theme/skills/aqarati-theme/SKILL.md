---
name: aqarati-theme
description: Create or restyle PPTX/PDF decks in the Aqarati brand theme (dark purple cover, lavender content slides, green accent, gold Arabic wordmark, Calibri). Use when the user asks for an Aqarati deck/PDF/slides, wants the Aqarati theme applied to an existing .pptx, or wants a roadmap/pitch/status deck in Aqarati style.
---

# Aqarati theme

Tokens live in `theme.json` (colors, fonts, sizes, status-badge colors, footer text). Engine: `scripts/aqarati.py` (needs `python-pptx`, `reportlab`; `pip3 install python-pptx reportlab`).

## Build new deck from content

1. Write a spec JSON (see `examples/sample.json`). Slide types:
   - `cover`: `arabic`, `title`, `subtitle`, `tagline`, `footnote`
   - `section`: `eyebrow`, `title`, `lede`
   - `cards`: `eyebrow`, `title`, `lede?`, `badge?`, `per_row?`, `cards[]` = `{title, body, tag?, icon?, highlight?}`
   - `bullets`: `eyebrow`, `title`, `lede?`, `badge?`, `bullets[]`
   - any slide: `notes` (speaker notes, PPTX only)
   - top-level `footer` overrides footer text.
2. Run:
   ```bash
   python3 ~/.claude/skills/aqarati-theme/scripts/aqarati.py build spec.json out.pptx --pdf   # pptx + pdf
   python3 ~/.claude/skills/aqarati-theme/scripts/aqarati.py build spec.json out.pdf          # pdf only
   ```
3. Render check: `pdftoppm -r 60 -png out.pdf preview`, view PNGs. Fix overflow (shorten text, lower `per_row`) and rebuild.

Tag/badge labels with theme colors: CONFIRMED, NEW (green), PENDING DECISION (deep purple), FUTURE, PHASE 2 (violet). Other labels default green; edit `status` in `theme.json` to add.

## Restyle existing deck

```bash
python3 ~/.claude/skills/aqarati-theme/scripts/aqarati.py restyle in.pptx out.pptx
```
Sets slide 1 dark, rest lavender; Calibri; title/body text colors by size; adds footer. Does not move shapes or recolor fills — check render after, adjust by hand. For PDF of result: needs LibreOffice (`soffice --headless --convert-to pdf`); if missing, rebuild from spec with the `build` command instead.

## Design rules (keep when hand-editing)

- 16:9, 13.333x7.5 in. Cover bg `1F1640` + two big offset circles (`2E2163`, `3A2A7A`) bleeding off right edge.
- Content bg `F5F3FC`; white rounded cards, 0.75pt `E4DFF5` border, soft shadow. Highlight card fill `F7FCF9`.
- Eyebrow: 13pt bold caps green `138A5E`. Title 36pt bold `1F1640`. Lede 16pt `6B6685`. Body `2B2640`.
- Icon circle `E3F4EC` top-left of card; status pill top-right.
- Footer 10pt `6B6685`: brand left, slide number right.
- Arabic wordmark `عقاراتي` gold `D8B25A` above "AQARATI" on cover.
