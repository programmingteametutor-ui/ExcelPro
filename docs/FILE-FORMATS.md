# File formats and round-trip limits

## XLSX (.xlsx, .xlsm)

| Feature | Import | Export | Notes |
|---|---|---|---|
| Multiple sheets, order, names | ✅ | ✅ | Chart sheets and dialog sheets are skipped and reported |
| Hidden / very hidden sheets | ✅ | ✅ (as hidden) | |
| Values: numbers, text, booleans, errors, dates | ✅ | ✅ | Dates are serial numbers with a date format, as in Excel. 1904-date workbooks are shifted to the 1900 system with a warning |
| Formulas | ✅ | ✅ | Exported as formulas with cached results. `_xlfn.`/`_xlws.` prefixes are stripped on import and added on export so newer functions (XLOOKUP, IFS, FILTER…) open in Excel |
| Shared formulas | ✅ (expanded) | written as normal formulas | |
| Unsupported functions | kept as text, cell shows `#NAME?`, listed in the import report | exported unchanged | Never replaced by blanks |
| Dynamic-array results | ✅ | exported as array formulas with the spill range | Legacy CSE arrays are imported as dynamic arrays (reported) |
| Defined names (workbook scope) | ✅ | ✅ | Sheet-scoped names are reported, not imported. Print areas (`_xlnm.Print_Area`) round-trip |
| Fonts (family, size, bold, italic, underline, strike, colour) | ✅ | ✅ | Theme and indexed colours resolve to RGB (tint applied) |
| Fills | ✅ solid | ✅ solid | Pattern/gradient fills import as their main colour |
| Borders | ✅ | ✅ | Styles map to thin/medium/thick/dashed/dotted/double |
| Alignment, wrap, indent | ✅ | ✅ (indent included) | Text rotation is reported and shown horizontally |
| Number formats | ✅ built-in and custom | ✅ | Currency symbols are written quoted (`"$"#,##0.00`) for compatibility |
| Merged cells | ✅ | ✅ | |
| Column widths, row heights, hidden rows/columns | ✅ | ✅ | |
| Tab colours, gridline visibility | ✅ | ✅ | |
| Frozen panes | ✅ | ✅ | Split (unfrozen) panes are reported |
| AutoFilter | ✅ range | ✅ range + hidden rows | Filter criteria are not exchanged; rows hidden by a filter are exported as hidden |
| Data validation | ✅ | ✅ | List, whole, decimal, date, time, text length, custom; input and error messages |
| Conditional formatting | ✅ | ✅ | Cell value, text, duplicates/unique, blanks/errors, formula, top/bottom, average, color scales, data bars, icon sets (3 sets); styles written as dxf |
| Notes (legacy comments), hyperlinks | ✅ | ✅ | Threaded comments import as notes. External links only (internal `#Sheet!A1` links are kept in ExcelPro files) |
| Tables | ✅ | ✅ | Name, range, header row, total row, style family, banding. Structured references are converted to ranges on import |
| Page setup | ✅ | ✅ | Orientation, paper, scale/fit to width, margins, gridlines/headings, header/footer, manual breaks |
| Pivot tables | ❌ (values kept, reported) | values only | ExcelPro pivots are saved in ExcelPro's own format |
| Charts, images, shapes | ❌ reported | ❌ | ExcelPro charts are saved in ExcelPro's own format and appear in PDF export |
| Protection, sparklines, macros, external links | ❌ reported | ❌ | |

Corrupt, encrypted or non-workbook files are rejected with a specific message.

Validated: files written by openpyxl and LibreOffice import correctly; files ExcelPro exports open in openpyxl with formulas intact and recalculate to the same values in LibreOffice (see `tests/files.test.ts`).

## CSV / TSV

- Import: delimiter auto-detected (or `\t` for .tsv), UTF-8 BOM removed, quoted fields and embedded newlines handled. Each value is parsed the way typed input is (numbers, %, currency, dates, booleans); text such as `00501` stays text when quoted with a leading apostrophe in ExcelPro, otherwise it becomes a number like in Excel.
- Export: one sheet, displayed values (as shown), choice of comma, semicolon, pipe or tab, optional BOM.
- CSV holds values only: formatting, formulas and other sheets are not part of the format.

## Not supported yet

- `.xls` (BIFF) and `.ods`: users are asked to save as .xlsx first.
- PDF export and printing: Phase 3.

## PDF

Export from Page Layout → PDF or Print preview → Download PDF (active sheet or all visible sheets). Uses the sheet's page setup (paper, orientation, margins, scale, print area, breaks, gridlines, headings, header/footer). Pages are rendered as images at about 160 dpi, so text isn't selectable; cell fills, borders, number formats, conditional fill colours and charts are included, data bars and icon sets are not.
