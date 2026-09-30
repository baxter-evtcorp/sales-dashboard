# Sales overview

Open `index.html` in current Microsoft Edge or Chrome and choose a local `.xlsx` Deal Log. No build or server is required. Choose the file again after saving updates. Replacing a file clears the old model and filters before parsing; an invalid replacement cannot leave stale totals on screen.

## Combined rollup

The default view adds HW/SW revenue and GP to PS exactly once per row. PS uses the complete actual SOLD AT / BOUGHT AT / PROFIT set when present, including zero. Otherwise it uses the complete estimated SOLD AT E / GROSS PROFIT E pair (with its validated cost input). Partial or invalid actuals are flagged and fall back as a whole pair, never per metric. Complete actuals can be used even if estimates are missing. PS actual PROFIT is explicitly labeled as the GP proxy because the source has no actual gross-profit field. The component table reconciles the filtered totals and labels estimated contributions. All date/customer/vendor filters apply to both logs in the combined view.

## Views

- Hardware/software: `SOLD AT`, `GROSS PROFIT`, aggregate gross margin.
- PS estimates: `SOLD AT E`, `GROSS PROFIT E`, aggregate estimated gross margin.
- PS actuals: actuals `SOLD AT`, `BOUGHT AT`, `PROFIT`; includes only rows with all three fields. Missing actuals are unavailable, not zero.

Filters: year, calendar quarter, customer, vendor. Charts: monthly sales and customer concentration. Vendor summary and searchable, paginated detail provide source row references. Source checks surface exclusions, repeated IDs, missing vendors, negative gross profit and stale formula caches. Counts are rows, not unique deals. No accounting revenue recognition or OEM mapping is inferred.

## Source contract

Reads only `2026 HW+SW Log` and `2026 PS Log`. Finds the header within the first 12 rows and maps by normalized header names; reordered columns are supported. Missing/renamed/duplicate headers stop import with a mapping message. A renamed tab requires updating the source contract. All other tabs are excluded, including summaries and historical templates.

Date Closed accepts Excel date serials (1900/1904 systems), MM/DD/YY, MM/DD/YYYY or ISO YYYY-MM-DD. Two-digit years mean 2000–2099. Coverage is derived from valid included records and does not assert completeness. Template rows containing only IDs/formulas and summary totals are excluded. Invalid business rows are reported, not silently converted to zeros.

Financial inputs must be numeric. A small arithmetic parser recalculates same-row `+ - * /` formulas and parentheses. It never uses `eval` or executes spreadsheet code. Blank Deduction references mean zero, matching the inspected template. Other missing inputs, cross-row formulas, Excel errors, functions and external references are rejected for that row with a source finding. For a new formula pattern, review its business meaning before extending the parser. Numeric GP overrides remain as supplied. Cached values that differ from recalculation are flagged. This is not a general Excel calculation engine.

Amounts display the workbook's dollar convention; the currency is not independently verified. No targets, pipeline probability, invoiced revenue, paid status or annualized forecast is invented.

## Privacy and publishing

All workbook processing runs in memory. No upload, runtime CDN, telemetry, network call, browser storage or data export is used. CSP blocks connections and restricts scripts/styles to this directory. Spreadsheet text is rendered via `textContent`. Clear workbook or close the page to discard session data. A website host receives ordinary page requests, not selected workbook contents.

Publish **only this directory**, never its parent or sibling `source/`. No real records, source workbook, customer names, financial totals, private screenshots, or real-data fixtures belong in this directory. Ignore rules supplement the mandatory source-only audit. Do not alter an existing repository or publish until the exact repository and scope are approved. No repository is initialized by this project.

## Bundled dependency

SheetJS Community Edition 0.20.3 is vendored in `vendor/xlsx.full.min.js`, with its Apache 2.0 license. Official source: https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js

SHA-256: `cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41`

Vendoring guidance: https://docs.sheetjs.com/docs/getting-started/installation/standalone/

Input limits: 25 MB file size, 80 MB declared expanded ZIP content, 5,000 ZIP entries, 20,000 sheet-row index and 201 columns. Encrypted and multipart workbooks are unsupported. Keep the bundled library current through deliberate, tested updates; no automatic update mechanism is used.
