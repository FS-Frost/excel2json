# CLAUDE.md

excel2json: converter and editor for Excel (.xlsx), CSV/TSV and JSON as an offline-first **static PWA on GitHub
Pages** (https://fs-frost.github.io/excel2json/). No backend: files never leave the browser. UI in Spanish and English.
Modernized on 2026-11-06 following `../mai-docs` (main reference) and `../sensei`; plan in `docs/plan-modernizacion.md`
(local, gitignored).

## Commands

- `bun run dev` — Vite dev server (port 5173). `?demo` (or `?ejemplo`) loads the built-in example workbook.
- `bun test` — pure logic tests in `tests/` (core conversion, edit ops/inverses, view/find, i18n parity + prefs schema,
  IndexedDB storage via `fake-indexeddb/auto`). Fixtures/factories in `tests/factories/`. Rune files (`*.svelte.ts`)
  can't be imported from Bun tests: keep testable logic in plain `.ts` (e.g. `client/prefs-schema.ts`).
- `bun run check` — svelte-check; keep 0 errors / 0 warnings.
- `bun run build` — static site in `build/`. For Pages: `BASE_PATH=/excel2json bun run build`. CI copies `404.html` →
  `index.html`. `bun run preview` caches the file index at startup: restart it after each build.
- Deploy: `.github/workflows/main.yml` on push to `main` (Bun 1.4.4 pinned, `--frozen-lockfile`, check, test, build,
  push `build/` to `gh-pages`). Every push to main deploys: keep main green.

## Stack

- SvelteKit 3 SPA (`adapter-static`, `ssr = false`, fallback `404.html`) + Svelte 5 runes + TypeScript 6 + Vite 8 + Bun.
- Kit 3 specifics (as in mai-docs/sensei): no `svelte.config.js` (Kit options in `vite.config.ts`), alias `#lib/*`
  through package.json `imports` and **imports need file extensions**, `$app/state` (no stores), `tsconfig.json`
  excludes `src/service-worker.ts`.
- Tailwind v4 via `@tailwindcss/vite`. Colors are tokens (`app.css`: `--c-*` vars → `@theme inline` → `bg-surface`,
  `text-muted`, `border-line`, `bg-sel`…); no raw colors in components except white text on fills. `data-theme` is
  always concrete (`light|dark`), set before paint by `app.html`. Shared classes: `.btn`, `.tb`, `.input`, `.select`,
  `.check`, `.menu`, `.menu-item`, `.panel`, `.seg`.
- Libraries: `exceljs` (xlsx read/write, loaded with `import()`; browser UMD build in the app/Worker, Node build in
  tests), `papaparse` (CSV), `fflate` (zip), `idb` (IndexedDB), `zod` (prefs validation).

## Architecture

- **One model**: every input becomes `Table[]` (`core/types.ts`: `headers`, `types: ColType[]`, `rows: Cell[][]`,
  `Cell = string | number | boolean | null`, dates as ISO text) and every output is written from `Table[]`.
- `lib/core/` (pure): `read.ts` (dispatch by extension/content), `xlsx.ts`, `csv.ts`, `json.ts`, `table.ts`
  (`gridToTable`: header row, empty rows, width, header normalization), `cell.ts` (exceljs value → Cell, ISO dates,
  `inferText`, `parseTyped`, `coerce` by column type, `cellText`), `naming.ts`, `nested.ts` (`a.b.c` ⇄ objects),
  `merge.ts`, `outputs.ts` (`buildOutputs` by format × mode, `bundle` → zip), `sample.ts`.
- `lib/worker/convert.worker.ts` + `protocol.ts`: parse/build/sample off the main thread with progress messages.
  `client/converter.svelte.ts` = RPC by id; cancel = terminate the Worker (next call creates a new one).
- `lib/edit/` (pure): `ops.ts` (`apply` mutates a Table in place and returns the **exact inverse** op; ops: cells,
  insert/delete/restore rows and cols, header, name, type, permuteRows/Cols), `history.ts` (steps = ops + inverses,
  limit 200), `view.ts` (filters = view only; `sortOrder` → permutation applied as an op), `find.ts`, `clipboard.ts`
  and `layout.ts` (`Axis`) copied from mai-docs.
- `client/session.svelte.ts`: files (`$state.raw`), `rev` counter, selection for export, active table, views
  (filters, widths by header name), history, dirty-table tracking for autosave, originals (for "reapply import
  options"; restored sessions have none).
- `client/persist.svelte.ts` + `storage/db.ts`: IndexedDB `excel2json` v1, stores `meta` (one key `current`) and
  `tables` (one entry per table) so autosave writes only dirty tables. While an unanswered saved session exists
  (`saved.meta`), nothing is saved; adding files discards it. Autosave timing from mai-docs (`storage/autosave.svelte.ts`).
- `client/prefs.svelte.ts` (+ pure `prefs-schema.ts`): import/export options in `localStorage['excel2json.prefs']`,
  each field falls back to its default. Other keys: `excel2json.theme`, `excel2json.lang`, `excel2json.version`.
- UI: `routes/+page.svelte` (landing ↔ workspace, drag & drop anywhere, launchQueue, global shortcuts),
  `components/grid/DataGrid.svelte` (virtualized editable grid), `grid/FindBar.svelte`, `FilesPanel.svelte`,
  `ExportPanel.svelte`, `SheetTabs.svelte`, `Landing.svelte`, `PasteDialog.svelte`; shared from mai-docs: `Toaster`,
  `Dialogs` (`confirmDialog`/`promptDialog`), `ContextMenu`, `Icon`, `ThemeToggle`, `UpdateNotice`.
- i18n (`lib/i18n/`, from mai-docs): `en.ts` defines `Dict`, `es.ts: Dict`, `t()` reactive, `warningText`/`errorText`
  map core warnings/`ConvertError` codes to text. New UI text goes in both dictionaries (`tests/i18n.test.ts`).
- PWA: `static/manifest.webmanifest` (file_handlers xlsx/csv/tsv/json), `src/service-worker.ts` (precache build +
  **Worker graph crawl**: Kit lists only the Worker entry, not its exceljs chunk), cache `excel2json-<version>`,
  update flow from mai-docs (`client/pwa.ts`, `UpdateNotice.svelte`: never reloads while files are open). Icons:
  `static/icons/*.svg` (+ `favicon.svg`, `Logo.svelte`, same drawing); PNGs rendered from the SVGs with headless
  Chromium — regenerate them if the SVGs change.

## Decisions (2026-11-06)

- Export modes: `split` (one file per table: `file.json`, or `file - Sheet.json` when the file has several),
  `perFile`, `combined` (JSON `{name: rows}` / one workbook with a sheet per table; names `file › sheet` when there are
  several input files), `merged` (union of columns by name, origin column `_origen`/`_source` by UI language).
  CSV only has split/merged. Several files → `json.zip`/`xlsx.zip`/`csv.zip`/`tsv.zip`.
- Types by default; "Everything as text" reproduces the old excel2json (all strings, `""` for empty cells in JSON).
- Empty header → its column letter (`C`), duplicates → `name_2`; no header → `A`, `B`, `C`.
- CSV: `;` ⇒ decimal comma (read and write), `,`/tab ⇒ dot. Only clean numbers are inferred (`00123` stays text).
- xlsx write: ISO strings in `auto`/`date` columns become real Excel dates; `.xls`/`.ods` not supported.
- Filters are view-only; export uses all rows unless "Only filtered rows" is checked. Sort is a real (undoable) edit.
- Undo history is not persisted (a restored session starts with an empty history).
- Supported browsers: Chromium and Firefox (as mai-docs).

## Invariants / gotchas

- Tables are mutated in place. Anything rendering table data must depend on `session.rev` **and get a new reference**
  (shallow copy per revision, as `DataGrid`'s `table`, `+page.svelte`'s `active`, `FilesPanel`'s `files`); deriving the
  same object, or an intermediate `{@const row = table.rows[r]}`, will not refresh.
- Never increment a reactive counter inside an effect without `untrack` (`Session.bump` already does).
- Tables sent to the Worker or IndexedDB must be plain objects (never `$state` proxies): files are `$state.raw`.
- `.gitignore` entries must be anchored (`/docs`): Tailwind v4 skips gitignored paths when scanning classes.
- Known limit: the first autosave of a very large table (100k × 20) clones it on the main thread (~150 ms hitch).
- Measured on the prod build (headless Chromium, 100k × 20 xlsx): parse ~9–11 s in the Worker, scroll ~56–58 fps,
  cell edit ~60 ms, JSON export ~5 s.
