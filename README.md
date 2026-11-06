# excel2json

![main](https://github.com/FS-Frost/excel2json/actions/workflows/main.yml/badge.svg)

https://fs-frost.github.io/excel2json/

Convert and edit **Excel (.xlsx), CSV/TSV and JSON** right in your browser. Nothing is uploaded: files are read and
written on your device, and the app works offline once installed (PWA). UI in English and Spanish.

## Features

- **Any direction**: xlsx ⇄ JSON ⇄ CSV/TSV, several files at once, with all their sheets.
- **Multiple sheets**
  - choose which sheets to export;
  - one file per sheet, one file per input file, everything in one file (one JSON `{"Sheet": [rows]}` or one
    workbook with a sheet per table) or **merge** every sheet into a single table (columns joined by name, optional
    origin column).
- **Real types**: numbers, booleans, `null` and ISO dates (`2024-01-31`, `2024-01-31T08:30:00`); formulas export their
  result. Turn on *Everything as text* for the classic all-strings output.
- **Editable preview**: a virtualized grid (100k+ rows) where you can edit cells, rename headers, insert, delete and move
  rows and columns, set a type per column, copy/paste with Excel, undo/redo, find and replace, sort and filter.
- **Options**: header row, skip empty rows, nested objects ⇄ `a.b.c` columns, JSON indentation, CSV separator (`;` uses
  a decimal comma, like Excel in Spanish), BOM, formatted Excel output (bold header, filters, frozen row, widths).
- **Autosave**: the working session is stored in IndexedDB and can be restored after closing the tab.
- **Installable PWA**: offline, "Open with excel2json" for .xlsx/.csv/.tsv/.json (Chromium).

Try it with the built-in example: https://fs-frost.github.io/excel2json/?demo

## Supported input

| Format | Notes |
| --- | --- |
| `.xlsx`, `.xlsm` | Every sheet. Merged cells keep the value in the top-left cell. Excel errors (`#N/A`…) become empty, with a warning. Old `.xls` is not supported. |
| `.csv`, `.tsv`, `.txt` | Delimiter detected (`,` `;` tab `\|`). UTF-8 with or without BOM; otherwise Windows-1252. Only clean numbers are typed (`00123` stays text). |
| `.json` | `[{…}]` (columns = union of keys of every row), `[[…]]` (first row = header), `[1, 2]`, `{"Sheet": [...]}` (several tables) or a single `{…}`. |

## Development

Requires [Bun](https://bun.sh) (CI uses 1.4.4).

```sh
bun install
bun run dev      # http://localhost:5173
bun test         # unit tests (tests/)
bun run check    # svelte-check, must stay at 0 errors / 0 warnings
bun run build    # static site in build/
```

For GitHub Pages the site lives under `/<repo>`: `BASE_PATH=/excel2json bun run build` (the CI does it).
`bun run preview` caches the file list at startup: restart it after every build.

Stack: SvelteKit 3 (static SPA) · Svelte 5 · TypeScript 6 · Vite 8 · Tailwind CSS 4 · exceljs · papaparse · fflate ·
idb · zod. The conversion runs in a Web Worker.
