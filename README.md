# ExcelPro

A browser-based spreadsheet with a real formula engine. This repository contains **Phases 1–3** of the ExcelPro spec (spreadsheet foundation, professional spreadsheet features, advanced analysis): a working editor, not a mockup. ExcelPro is an independent product with its own branding and is not affiliated with Microsoft.

## Quick start

```bash
npm install
npm run dev          # http://localhost:5173 (rebuilds on change)
npm test             # 65 unit/integration tests (engine, files, persistence, phase 2–3 features)
npm run build        # dist/index.html + dist/assets
npm run build:single # dist/excelpro.html, one self-contained file
npm run bench        # engine benchmark at 10k / 50k / 100k rows
npm run compat       # regenerates docs/COMPATIBILITY.md from the live engine
```

Browser tests (Playwright): `smoke.mjs` (32 checks, including a 100k-row import) and `features.mjs` (26 checks covering filters, sorting, autocomplete, validation, conditional formatting, tables, charts, pivots, print preview/PDF and XLSX export of all of these):

```bash
npm run build && npm run serve &
PLAYWRIGHT_MODULE=playwright CHROME_PATH=$(which chromium) npm run test:e2e
PLAYWRIGHT_MODULE=playwright CHROME_PATH=$(which chromium) npm run test:e2e:features
```

## Publish on GitHub Pages

The build output is fully static (relative asset paths, `#/` routes), so it runs from any sub-path such as `https://<user>.github.io/<repo>/`.

1. Push this folder to a GitHub repository (branch `main`).
2. In the repository: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. The workflow in `.github/workflows/pages.yml` installs, type-checks, tests, builds and deploys on every push. The site URL appears in the workflow run and on the Pages settings screen.

Workbooks are saved in each visitor's own browser (IndexedDB); nothing is uploaded. HyperFormula is GPLv3, so a public deployment should keep this source public under a GPL-compatible licence (or use a HyperFormula commercial licence).

## What works in Phase 1

| Area | Status |
|---|---|
| Grid | Virtualized rendering (only visible cells exist in the DOM), A1 addressing up to XFD1048576, selection, multi-range selection (Ctrl+click), row/column selection, drag-select with auto-scroll, fill handle with series detection (numbers, dates, "Item 1", months, weekdays), resize/AutoFit/hide/unhide rows and columns, merged cells (horizontal and vertical), text overflow into empty neighbours, `####` / General shrink-to-fit for numbers that don't fit, context menus |
| Editing | Type-to-edit, F2/double-click edit, Enter/Tab/Shift variants, Esc cancel, Alt+Enter newline, Ctrl+Enter fill selection, click-to-insert references while typing a formula, formula bar with Name Box navigation (cells, ranges, `Sheet!A1`, defined names) |
| Formulas | HyperFormula engine with dependency graph and incremental recalculation; relative/absolute/mixed refs, cross-sheet refs (incl. quoted names), named ranges, dynamic-array spill, circular reference detection, plain-language error explanations in the status bar. 139 of 165 spec functions fully supported, 11 partially — see [docs/COMPATIBILITY.md](docs/COMPATIBILITY.md) |
| Input parsing | Excel-like: `1,234.50`, `15%`, `$5`/`₹5`/`€5`, dates (ISO, M/D/Y or D/M/Y setting, `5-Mar-2024`), times, TRUE/FALSE, `'00123` as text; number formats applied automatically |
| Formatting | Font family/size, bold, italic, underline, strikethrough, font & fill colour, borders (all, outside, thick, per side), horizontal/vertical alignment, wrap, indent (import), merge/merge & center/unmerge, number formats (General, Number, Currency, Accounting, dates, time, %, fraction, scientific, text, custom codes) with increase/decrease decimals, format painter |
| Sheets | Add, delete, rename, duplicate, reorder (drag), tab colours, hide/unhide, Ctrl+PageUp/Down; references update on rename |
| Clipboard | Internal copy/cut/paste with reference adjustment, tiling, paste values/formulas/formats; paste TSV from Excel/Sheets; copies TSV + HTML out |
| Undo/redo | One command history for edits, paste, fill, formatting, structure, sheet operations (200 steps); typing in a cell is one step |
| Find & replace | Sheet or workbook, values or formulas, match case, entire cell, find all list, replace one/all (undoable) |
| Files | XLSX import/export, CSV/TSV import/export (delimiter + BOM options), import report listing everything not imported |
| Persistence | IndexedDB: per-sheet records (only changed sheets are rewritten), debounced autosave, honest Saving/Saved/Unsaved status, retry with backoff, cross-tab conflict detection with "keep mine / load theirs", unsaved-changes warning, file list with rename, copy, download, trash/restore |
| UI | Ribbon with all nine tabs, light/dark theme, zoom 50–200%, gridline/heading/formula-bar toggles, full screen, platform-aware shortcuts, responsive down to phone width |

## What Phases 2–3 add

| Area | Status |
|---|---|
| Formatting | Format Cells dialog (number, alignment, font, border, fill; Ctrl+1), indent buttons, conditional formatting: cell value / text / duplicates / unique / blanks / errors / formula rules, top/bottom N or %, above/below average, 2- and 3-color scales, data bars, icon sets; rule manager with priority order and stop-if-true |
| Sort & filter | Quick sort A→Z / Z→A, multi-level custom sort (header detection, keeps row formatting, adjusts formulas), AutoFilter (Ctrl+Shift+L) with per-column value lists + search, condition filters (contains, begins, >, between, top N, above average…), clear / reapply |
| Data tools | Data validation (list with in-cell dropdown, whole/decimal/date/time/text length/custom formula; input messages; stop/warning/information alerts; circle invalid data), text to columns (delimited/fixed width), Flash Fill (Ctrl+E), remove duplicates, text cleaning (trim, clean, case, text→number), Goal Seek |
| Formulas | Autocomplete for functions, names and tables with argument hints (cell editor and formula bar), function category menus and recently used, Name Manager, more functions (see compatibility doc), structured table references (`Table1[Units]`) |
| Tables | Format as Table (Ctrl+T), 6 styles, banded rows, header filters, total row (SUBTOTAL), auto-extend when typing below, rename, convert to range; contextual Table Design tab |
| PivotTables | Create from a range or table on a new or existing sheet; rows, columns, values (sum, count, average, min, max, distinct count), filters, date grouping (month/quarter/year), sort, grand totals, refresh / refresh all / optional refresh on change; field-list side panel |
| Charts | Column, bar, line, area, pie, doughnut, scatter, histogram; drag to move, resize, Delete to remove; side panel for type, data range, series orientation, titles, legend, data labels, gridlines, stacking, series colours/visibility, histogram bins |
| View | Freeze panes / top row / first column, page break preview, notes (Shift+F2) and hyperlinks (Ctrl+K, Ctrl+click to open, `#Sheet!A1` links) |
| Print | Margins, orientation, paper size (A4/Letter/Legal), scale or fit to width, print area, manual page breaks, print gridlines/headings, header/footer codes (&P &N &D &A &F), print preview (Ctrl+P), browser printing, PDF export (sheet or whole workbook) |
| Files | XLSX import/export of freeze panes, AutoFilter, data validation, conditional formatting (with dxf styles), hyperlinks, notes (comments + VML), tables, page setup, margins, breaks and print area |

Ribbon tabs show working controls only. Features planned for later phases appear as a labelled note on their tab, never as buttons that do nothing.

## Architecture

```
src/model/        pure data types, A1 addressing, axis metrics (sparse sizes, O(log n) lookups)
src/engine/       Workbook (HyperFormula wrapper + formatting/layout + command/undo history),
                  input parser, Excel number-format renderer, formula text utilities,
                  ExcelPro function plugin, compatibility registry
src/files/        XLSX reader/writer (JSZip + small XML parser), CSV/TSV (PapaParse)
src/persistence/  WorkbookStore interface, IndexedDB + in-memory stores, autosave controller
src/grid/         virtualized grid component and selection model
src/ui/           app shell: file list, editor, ribbon, formula bar, sheet tabs, status bar, dialogs
supabase/         Phase 4 schema with Row Level Security and an atomic save-with-conflict-check RPC
```

Layer rules: the engine and file layers have no React or DOM imports (they run under Node in tests); the UI talks to the model only through `Workbook` methods; persistence only sees serialized `WorkbookDoc`/`SheetDoc` data.

**Undo design.** Every content-changing command runs as one HyperFormula batch, so it is exactly one engine undo entry; formatting changes are stored as diffs and structural changes as snapshots of the affected sheet's layout. Undo = `hf.undo()` plus restoring the layout state.

### Library choices (verified)

| Need | Choice | Notes |
|---|---|---|
| Formula engine | HyperFormula 3.x | 400+ functions, dependency graph, undo/redo, CRUD with reference adjustment. **License: GPLv3 or commercial.** The code uses the `gpl-v3` key; a closed-source or commercial product must buy a HyperFormula commercial license. |
| XLSX | JSZip + own OOXML mapping | ExcelJS (MIT) was evaluated; its Node-stream dependencies made a browser bundle impractical here. SheetJS Community Edition doesn't write styles. The custom mapping round-trips formulas, styles, merges and sizes and is tested against files from openpyxl and LibreOffice. |
| CSV | PapaParse (MIT) | |
| Grid | Own virtualized grid | Glide Data Grid was evaluated: it needs Linaria's build-time CSS transform and has no row-spanning merges. The grid is isolated in `src/grid/` behind a small props interface. |
| UI | React 19, Tailwind v4, shadcn/ui-style primitives | Components follow shadcn conventions (cva + tailwind-merge) without Radix; swapping in the shadcn CLI components later is mechanical. |

**Build note:** the build uses esbuild and Tailwind's compiler directly (`scripts/build.mjs`) instead of Vite. It produces the same static output and keeps a single-file mode for embedding.

## Known limitations

See [docs/FILE-FORMATS.md](docs/FILE-FORMATS.md) and [docs/COMPATIBILITY.md](docs/COMPATIBILITY.md). In short:

- Not yet built (Phase 4+): cloud accounts, sharing and real-time collaboration (the Supabase schema exists but is not wired), threaded comments, protection, version history, templates, automation, sparklines/images/shapes.
- Charts are kept in ExcelPro's own saved format and in the PDF, but are **not written to .xlsx** (and charts in imported .xlsx files are reported, not imported). Pivot tables are exported as their values.
- PDF export is rasterised (each page is an image at ~160 dpi); text in the PDF is not selectable. Data bars and icon sets are not drawn in print.
- Structured table references are converted to normal ranges when a formula is entered (the formula keeps working but does not grow with the table afterwards).
- `.xls` and `.ods` are not opened; users are told to save as .xlsx.
- Formatting a whole column applies to the used area (+100 rows), not to all 1,048,576 rows.
- Circular references show `#CYCLE!` (no iterative calculation). LET, LAMBDA, TEXTSPLIT, INDIRECT and CELL are not supported.

## Performance (measured)

See [docs/PERFORMANCE.md](docs/PERFORMANCE.md). Headline: a 100,000-row × 5-column workbook with 200,000 formulas loads in ~1.8 s, a change that ripples through every row recalculates in ~0.4 s, a change to the last row in ~1 ms; a 100k-row CSV opens in the browser in ~1.5 s and scrolls at ~45 fps.
