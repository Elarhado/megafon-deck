---
name: deck
description: Make a MegaFon presentation in the corporate brand style (template layouts, brandbook palette, Arial, typesetting rules) and deliver .pptx + PDF. Use whenever the user asks for slides or a presentation for MegaFon, or to restyle an existing deck into the MegaFon template.
argument-hint: "[тема, тезисы или файлы] [--style light|green|dark] [--no-pdf]"
---

# MegaFon deck

One command: `/megafon-deck:deck <тема / материалы>`. You plan the content, the scripts do the layout; never draw brand elements by hand.

Shorthands:
- `PY` = `python3` (on Windows `python`)
- `S` = `${CLAUDE_PLUGIN_ROOT}/scripts`, `A` = `${CLAUDE_PLUGIN_ROOT}/assets`

Talk to the user in Russian. Slide text in Russian unless asked otherwise.

## 1. Check the environment (once per session, silently)

Run `PY "S/build_deck.py" --check`. Exit 3 means a dependency is missing: run the printed `pip` command (python-pptx, Pillow), then re-check. Do not ask the user about this unless the install fails.

## 2. Understand the task

- Inputs: the user's text, attached files, an existing deck to restyle, numbers. Read attachments in full where they matter.
- Decide: audience, the 1-sentence message of the deck, number of slides (default 8–12), style.
  Style default: `light` for content slides, cover and section slides in `green` (brand: green dominates title slides). `dark` only when the user asks. Per-slide `"style"` overrides are allowed.
- If the request is thin (only a topic), ask at most 2 short questions with AskUserQuestion (audience/goal, length), otherwise proceed.
- **Facts.** Never invent numbers, dates or facts about real companies. Use the user's data. If research is needed and the Conductor plugin is installed, gather and verify facts with `/conductor:conduct` (or `conductor:verify`) first, then build the deck from verified facts only. If you must show a format without data, mark numbers as illustrative in the speaker notes and say so in your reply.

## 3. Plan the storyline

Write a slide plan before any JSON: for each slide a layout key and the takeaway.
- Titles are takeaways, not topics: «5G-сеть растёт быстрее LTE», not «Динамика сети». ≤ 60 characters.
- One idea per slide. Prefer structure (rows, pills, KPIs, charts) over walls of text. ≤ 5 bullets per block, ≤ 12 words per bullet.
- Choose layouts by content shape. Look at `A/slotmaps/<style>.png` once to see them (every placeholder is labelled with its idx), and read roles in `A/layout_roles.json`:
  - cover, section — opening and dividers
  - kpi_2x2 — 4 metrics (values ≤ 5 characters: «12,5», «38%»)
  - chart_two_items / chart_five_items / chart_wide — a chart with 2 / up to 5 / 3 commentary items; colour dots in the layout match series 1..5
  - rows_arrows (from → to), rows_bars (A | B), table_3col (3×3), plate_list (up to 5 label → text)
  - three_frames (3 pillars), three_steps_conclusion (1-2-3 → conclusion), four_arrows_conclusion (4 inputs → 2 outcomes)
  - three_pills (3 items with a big value each), two_pills (2 statements), two_plates_conclusion (2 options + verdict)
  - timeline (5 steps), icon_bg_list / icon_bg_two_columns / frame_image / three_pills_image (with icon or picture)
  - quote, blank (free canvas: chart, table, text, icons via `elements`)
- Slot capacity is small (8–14 pt boxes). Keep heads to 2–4 words, cells to 1–2 short lines. Long text belongs in speaker notes.
- Icons: pick from `A/icons/index.json` (tags in Russian and English) or look at `A/icons/sheet.png`. Colours: green, purple, white, black, berry — white on green/dark slides.

## 4. Write deck.json and build

Schema with an example: `S/deck.schema.md` (read it the first time). Write the file next to the user's working files as `<name>.deck.json` (keep it: the user can ask for edits later and you rebuild from it).

```
PY "S/build_deck.py" <name>.deck.json -o "<name>.pptx"
```
Exit 2 = validation error: read the message (it lists allowed roles), fix the JSON, rebuild.

## 5. Check — lint and look (always, until clean)

```
PY "S/lint_deck.py" "<name>.pptx"
PY "S/render.py" "<name>.pptx" --pdf "<name>.pdf" --png-dir "<tmp>/png"
```
- Fix every lint ERROR (overflow, off-brand font/colour, leftover prompt text, typesetting) by shortening text, moving detail to notes or choosing another layout; rebuild.
- Lint ERRORs are authoritative even when the picture looks fine: template titles use Arial Black, which renderers without that font (e.g. LibreOffice on Linux) replace with a narrower face, so the PNG can hide an overflow PowerPoint will show. Never deliver with lint errors; shorten instead.
- Open `<tmp>/png/sheet.png` and every slide PNG that looks dense. Check: nothing cut off or overlapping, readable contrast, no empty boxes, charts legible, green dominates and purple is an accent, no more than ~10 % additional colours. Fix and rebuild.
- `render.py` exit 3 means no LibreOffice and no PowerPoint for PDF/PNG export: deliver the .pptx, tell the user PDF needs LibreOffice (`brew install --cask libreoffice`) or «Файл → Экспорт → PDF» in PowerPoint, and do the visual check by reasoning about slot capacity (lint) only.

## 6. Deliver

Give the user the .pptx and the PDF (unless `--no-pdf`), one line on what is inside, and anything illustrative or unverified. Do not paste the slide text back into the chat.

## Brand rules the scripts cannot fully check

Source: `A/brand.json` (brandbook «Константы бренда», январь 2026).
- Green #00B956 is the main colour, purple #731982 the accent (~4:1). Additional colours (lime, grass, night, berry) for charts and small accents only; sky and flamingo only inside gradients.
- Arial only in PowerPoint (GT Walsheim is not for office files).
- «ёлочки», ё everywhere it belongs, em dash with spaces, en dash in ranges (5–6 октября), thousands with a narrow space (12 500), МегаФон with capital М and Ф. The builder applies the mechanical part; you write ё and phrasing correctly.
- Do not add or redraw the logo: it lives in the cover layouts. Legal tagline text is not needed in presentations.
