# deck.json schema

`build_deck.py deck.json -o out.pptx` — exit 0 ok, 2 validation error, 3 missing dependency.

Top level: `title` (deck name), `date` (string), `style` (`light` default | `green` | `dark`), `slides` (list).
`title` + `date` form the header label (name, line break, date) in the date placeholder of every slide that has one.
Slide numbers are automatic.

Slide: `layout` (key from `assets/layout_roles.json`), optional `style` override, `notes` (speaker notes;
`speaker_notes` is an alias), role fields, optional `elements`.

Role fields (names are the keys of `layout_roles.json[layout].roles`):
- scalar role (one placeholder): `"text"`, `["bullet", ...]` or `{"text": "head", "bullets": [...]}`
- list role (several placeholders, top-to-bottom / left-to-right): a list whose items are each of the forms above;
  shorter lists leave the remaining placeholders removed
- matrix role (`table_3col.rows`): list of rows, each a list of cell values
- `chart` (layouts with a chart placeholder, required): `{"type": "bar|column|stacked_column|line|pie|doughnut",
  "categories": [...], "series": [{"name": "", "values": [...]}], "number_format": "0", "labels": true, "legend": false, "colors": ["00B956"] (optional override; default: legend-dot order in chart_* layouts, green for a single series)}`
- `picture`: `{"icon": "<lucide name>", "color": "green|purple|white|black|berry"}` or `{"image": "path"}`
- `elements` (any layout, meant for `blank`), inches:
  `{"kind":"chart","x","y","w","h","chart":{...}}`, `{"kind":"table","x","y","w","h","rows":[[..]],"header":true,"pt":10}`,
  `{"kind":"text","x","y","w","h","text","pt":14,"bold":false,"color":"green"}`,
  `{"kind":"icon","x","y","size","name","color"}`, `{"kind":"image","x","y","w","h","path"}`

All text goes through typesetting (« », em dash, en-dash ranges, U+202F thousands and before ₽, МегаФон spelling).
Unused placeholders are removed; icons that do not fit a non-square placeholder are fitted, not cropped.

```json
{
 "title": "5G и развитие сети", "date": "14.10.2026", "style": "light",
 "slides": [
  {"layout": "cover", "style": "green", "title": "5G и развитие сети 2027", "label": "14.10.2026"},
  {"layout": "kpi_2x2", "title": "Ключевые показатели",
   "values": ["12 500", "+18%", "3,4 млн", "92%"],
   "labels": ["базовых станций", "рост трафика", "абонентов 5G", "покрытие городов"],
   "notes": "Цифры иллюстративные."},
  {"layout": "chart_two_items", "title": "Динамика",
   "chart": {"type": "column", "categories": ["2025", "2026"], "series": [{"name": "БС", "values": [8, 12]}]},
   "heads": ["Рост", "Нагрузка"], "texts": [["тезис 1", "тезис 2"], "тезис"],
   "summary_head": "Вывод", "summary": "Строим там, где растёт трафик"},
  {"layout": "blank", "title": "Свободный слайд", "elements": [
   {"kind": "table", "x": 0.6, "y": 1.4, "w": 8.6, "h": 1.6, "rows": [["Регион", "БС"], ["Север", "120"]]},
   {"kind": "icon", "x": 0.6, "y": 3.6, "size": 0.8, "name": "antenna", "color": "green"},
   {"kind": "text", "x": 1.6, "y": 3.7, "w": 6, "h": 0.6, "text": "Подпись", "pt": 14}]}
 ]
}
```
