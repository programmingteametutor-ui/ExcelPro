# Performance results

Measured on the build machine (Linux container, Node 22, headless Chromium). These are measurements, not guarantees; rerun with `npm run bench` and `npm run test:e2e`.

## Engine (`npm run bench`)

Workbook with 5 columns per row: 3 values, `=B*C`, and a running total `=E(prev)+D` (so the first row's change ripples through every row).

| Size | Build | Edit row 1 (full ripple) | Edit last row | Undo | Insert row | Serialize sheet | Export .xlsx | Import .xlsx |
|---|---|---|---|---|---|---|---|---|
| 10,000 rows (20k formulas) | 314 ms | 73 ms | 1 ms | 0 ms | 7 ms | 99 ms (0.5 MB) | 420 ms (0.5 MB) | 240 ms |
| 50,000 rows (100k formulas) | 1,025 ms | 206 ms | 1 ms | 0 ms | 18 ms | 210 ms (2.6 MB) | 1,698 ms (2.4 MB) | 832 ms |
| 100,000 rows (200k formulas) | 1,782 ms | 389 ms | 0 ms | 0 ms | 47 ms | 330 ms (5.2 MB) | 3,447 ms (4.8 MB) | 2,022 ms |

Observations: recalculation is incremental (an edit with few dependents costs ~1 ms regardless of size); undo is constant-time.

## Browser (`npm run test:e2e`)

| Scenario | Result |
|---|---|
| Import 100,000-row CSV (4 columns) and render | ~1.5 s |
| Jump to row 100,001 via Name Box | ~55 ms |
| Enter `=SUM(C2:C100001)` and see the result | ~0.3 s |
| Scroll through 100k rows (60 frames of 400 px) | ~17–23 ms average frame |

## Known costs / next optimizations

- XLSX import/export runs on the main thread; Phase 2 should move it to a Web Worker (the files layer has no DOM dependency, so it can move as-is).
- Autosave serializes whole changed sheets; very large sheets (100k rows) take ~0.3 s per save. A cell-level change log is the planned next step for the Supabase backend.
- The status-bar aggregate stops after 200,000 cells and marks the result as partial.
