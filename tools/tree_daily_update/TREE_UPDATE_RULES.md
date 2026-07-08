# TREE Update Rules

This file records the mandatory calculation and color rules for every future TREE update.
When updating TREE, the automation must follow these rules rather than re-deciding the logic.

## Scope

- Apply these rules to TREE macro data rows.
- For every routine update, first identify the latest authoritative TREE workbook in the workspace and preserve that workbook's actual structure.
- If the latest authoritative TREE is a simplified delivery workbook, keep it simplified and update only the data sheet(s) present in that simplified layout unless the user explicitly asks otherwise.
- If the latest authoritative TREE is a full/complex workbook, keep it full/complex and update the corresponding full/complex data sheet(s) required by that layout.
- Do not convert between simplified and full/complex structures during a routine update. The output must remain the same structure type as the latest authoritative input workbook.
- Archived full or simplified workbooks may be referenced for diagnostics if needed, but must not replace the current latest structure unless the user explicitly names them as the base.
- Do not apply macro color rules to A-share, Hong Kong stock, US stock, broad capital-market index, or industry rows unless the user explicitly asks.
- Treat sentiment, valuation, and industry sections as outside the macro color scope even when an underlying series is stored in the macro-data source sheet.
- Trend charts must keep the original style and remain anchored in the corresponding trend-chart cell.
- TREE trend charts must be clean mini line charts: keep the single red line and chart cache only; remove or disable data labels, point labels, series-name labels, legends, visible tick labels that crowd the cell, titles, gridlines, fills, and borders. Do not rely on hidden template labels remaining hidden after chart-cache refresh.
- During routine TREE updates, assume the user has manually refreshed and saved the daily workbook unless they explicitly ask Codex to refresh it. Do not run the automatic daily Wind refresh step by default.
- In routine TREE updates, identify the latest refreshed daily workbook in the workspace, verify it is the latest plausible daily data source, run only necessary data-integrity checks, then proceed to TREE updating.
- Do not update the daily workbook's front-sheet trend charts unless the user explicitly asks for daily chart refresh. The daily workbook is used as the refreshed data source; TREE trend charts are updated separately from the underlying daily/source series.

## Routine Update Workflow

Every routine TREE update must follow this sequence. Do not skip the audit steps even when the data update itself succeeds.

1. Identify inputs.
   - Select the latest authoritative TREE workbook from the workspace unless the user explicitly names one.
   - Select the latest manually refreshed daily workbook.
   - Rule: version/date and modification time are only candidates; verify workbook structure, sheet names, key headers, and xlsx integrity before using them.

2. Verify daily data source.
   - Confirm the required bottom sheets exist, at minimum `宏观数据`, `指数走势`, and `辅助数据` when present in the layout.
   - Check that bottom source sheets contain recent nonblank observations and that representative front-sheet values are not obviously stale versus bottom data.
   - Rule: do not run daily Wind refresh or daily chart refresh in a routine TREE update unless the user explicitly asks.

3. Preserve TREE layout.
   - Keep the selected TREE workbook's visible sheets, hidden-sheet state, row order, formulas, definitions, weights, formatting, and non-target sheets.
   - Rule: simplified input stays simplified; full/complex input stays full/complex.

4. Sync TREE data from bottom series.
   - Update current value, data date, MoM, YoY, YoY increment, macro colors, and TREE trend-chart cache from the same underlying series.
   - Rule: match by exact Wind/source code first. If a coded row is not found, verify monthly/supplemental source before carrying forward.

5. Update write time.
   - Compare current value and true Excel date value against the previous authoritative TREE.
   - Rule: update `写入时间` only when current value or data date changes. Normalize Excel serial dates and `datetime` values to the same calendar date before comparison; write only the calendar date in `yyyy-mm-dd` format.

6. Update trend-analysis text.
   - Generate or verify text from the same underlying historical series used for the TREE current value and trend chart.
   - Rule: do not infer trend text only from chart appearance, front-sheet cached value, one current value, or one MoM value.

7. Run final audits before delivery.
   - Verify xlsx integrity, sheet structure, non-target sheet XML, formulas, dates, write-time gaps, trend text blanks, chart anchors, chart cache, chart labels, and visual preview.
   - Rule: the file is not deliverable until all mandatory audits pass or the remaining issue is explicitly reported as a blocker.

## TREE Trend Chart Update Rules

- Update TREE trend charts by changing the chart value cache from the verified underlying series; do not copy daily front-sheet charts.
- Preserve each chart's original anchor cell, size, red single-line style, and worksheet location.
- Do not rewrite the workbook with a library path that is known to drop or corrupt chart XML. Prefer OpenXML-level edits for chart cache/style changes.
- After every chart-cache update, run the chart display cleanup step. This cleanup must remove all chart display elements that can create mini-chart text clutter:
  - `c:dLbls` data-label nodes at every chart level.
  - `c:legend` nodes.
  - Data-label flags such as legend key, value, category name, series name, and percent labels.
- Do not rely on template labels being hidden. If the template contains hidden `dLbls`, remove them after refresh because WPS/Excel may render them after cache updates.
- Final chart XML audit must pass:
  - chart count equals expected target-sheet chart count unless row additions/removals explain the difference.
  - chart anchors remain in the expected trend-chart column.
  - total `c:dLbls` count is 0.
  - total `c:legend` count is 0.
  - chart value caches have nonzero point counts for all rows with valid source series.
- Final visual audit must render or open a representative range containing the trend-chart column, normally the first 50 rows plus any rows touched by chart repair.
  - Rule: if rendering fails because of sandbox or file permissions, rerun the renderer with approved local-file write permission or clearly report that visual rendering was blocked. Do not claim visual verification when only XML audit was run.
- If the visual or XML audit finds labels, legends, crowded text, blank charts, wrong anchors, or broken lines, generate a corrected version and repeat both XML and visual audits before delivery.

## Version Selection

- Do not permanently lock TREE updates to a previously delivered workbook.
- At the start of every update, identify the latest TREE workbook because the user may revise the framework or layout between updates.
- If the user explicitly provides or names a workbook, use that workbook as the authoritative base.
- Otherwise, inspect the TREE workbooks in the workspace and determine the latest candidate using filename version/date and filesystem modification time.
- Before writing, verify the candidate's active framework, visible sheets, key headers, row layout, and indicator-code positions. Modification time alone is not sufficient.
- Never replace a newer framework with the structure of an older TREE workbook.
- Update the selected workbook in place conceptually: preserve its framework, weights, definitions, row order, formatting, hidden-sheet states, charts, formulas, and other user revisions unless the user explicitly requests changes.
- If the selected workbook is a simplified delivery version with only `重点策略跟踪情况(V3.0)` and `指标筛选与结论复核表`, keep the same two-sheet structure in the output and verify the non-updated sheet remains unchanged.
- If the selected workbook is a full/complex version with additional visible or hidden support sheets, preserve those sheets and update the workbook according to that layout instead of reducing it to the simplified structure.
- Output the next filename using `（MMDDVX）TREE宏观分析最新版.xlsx`, incrementing the version number for the same date.

## Data Calculation Rules

- China 10-year government bond yield uses Wind code `S0059749` as the single authoritative series in TREE.
- China-US 10-year government bond spread uses `S0059749-G0000891`. Do not use `M0325687` or `M0325687-G0000891` in TREE.

| Column | Rule |
|---|---|
| Current value | Use the latest value from the refreshed daily workbook. Preserve the display unit/format: percent indicators display as percentages; amount, balance, index, and level indicators display as normal numbers. |
| Write time / 写入时间 | Record the local TREE write/capture date for the current row's latest data. This is the date the update process first writes or changes the row's current value/data date in TREE, not the official data release date. |
| MoM / marginal change | Use latest value minus previous observation. Equity index rows are the exception: use latest trading-day percentage change. |
| YoY | If the Wind indicator itself is a YoY / cumulative YoY / monthly YoY indicator, use the current value directly. If the indicator is a level such as amount, balance, total, sales, GDP, money stock, or other stock/flow level, use `(current period - same period last year) / ABS(same period last year)`. If YoY growth is not economically meaningful, write `-`. |
| YoY increment / YoY increase | Use `current period - same period last year`. If the indicator itself is a YoY percentage indicator, this means the YoY reading changed by that many percentage points. |

## Write Time Column Rules

- TREE data sheets must include a `写入时间` column immediately after `数据日期` when the layout can support adding the column without breaking the user's structure.
- `写入时间` means the local TREE data write/capture date. It can be used as an internal update-event date, but it must not be described as the official statistical release date or Wind's exact backend publication time.
- During routine TREE updates, if a row's current value or data date changes versus the previous authoritative TREE workbook, update that row's `写入时间` to the current refresh/write date.
- If a row is recalculated but the current value and data date do not change, preserve the existing `写入时间`.
- If an older workbook first receives this column and no historical write time exists, initialize populated rows with the baseline workbook write time or the column-addition timestamp and record that this is an initialization baseline.
- Leave `写入时间` blank for rows without reliable current data/data date or for non-indicator summary rows.
- Use date-level precision only by default, formatted as `yyyy-mm-dd`; do not include hours, minutes, or seconds unless the user explicitly asks.

## YoY Growth Not Meaningful

For the following types of indicators, do not calculate a YoY growth rate in the YoY column. Write `-` and keep the YoY increment column:

- Net injection / net withdrawal / net liquidity operation
- Interest rate
- Spread
- Yield
- PMI and other diffusion indexes
- Reserve requirement ratio
- Leverage ratio
- Deficit ratio
- Other ratio indicators where percentage growth of the ratio itself is misleading

## Color Rules

| Column | Color rule |
|---|---|
| Current value | Compare the latest current value with its own trailing-one-year history. High percentile is red; low percentile is green. If the series is basically flat over the trailing year, do not fill color. |
| MoM / marginal change | Compare the latest marginal change with the absolute size of its own trailing-one-year historical marginal changes. Only unusually large changes get color. Positive is red; negative is green. Ordinary changes stay unfilled. |
| YoY | Compare the latest YoY value with its own trailing-one-year YoY history. High percentile is red; low percentile is green. Do not compare YoY with historical marginal-change sizes. |
| YoY increment / YoY increase | Compare the latest YoY increment with the absolute size of its own trailing-one-year historical YoY increments. Only unusually large increments get color. Positive is red; negative is green. Ordinary changes stay unfilled. |

## Formatting Rules

- Display precision does not change the stored value or calculation precision; rounding applies only to the visible cell format.
- Percentage, rate, ratio, spread, and growth indicators display with a percent sign and two decimal places, for example `7.50%`.
- Non-percentage values display the unscaled source number only. Do not append text units such as `万亿美元`, `十亿美元`, `亿元`, `点`, or similar suffixes.
- Do not use Excel number-format commas to scale a value into thousands, millions, billions, or trillions. For example, a source value of `23052.3` displays as `23052.30`, not `23.05万亿美元`.
- Non-percentage decimals display two decimal places. Whole numbers may remain whole numbers when no decimal precision exists in the source.
- YoY is always displayed as a percentage when it is numeric.
- Current value, MoM, and YoY increment follow the indicator type:
  - Percent indicators display as percentages with two decimal places.
  - Level indicators display as plain unscaled numbers without a unit suffix.
- Do not let inherited cell formats override the indicator unit. For example, PMI MoM is points, not percent; US M1 MoM is a level change, not percent.
- US commercial-bank credit, C&I loans, consumer loans, real-estate loans, M1, and M2 display their Wind/source values directly without conversion to USD trillions or USD billions.

## Trend Analysis Text Rules

- Trend-analysis text must be based on the underlying historical series in the refreshed daily workbook or the verified monthly/supplemental source, not only on TREE front-sheet values or daily front-sheet cached values.
- Before writing trend-analysis text, inspect the source series for the same Wind code used to update the TREE row. The current value, current date, previous observations, and recent historical path must come from the same underlying series used for the TREE current-value cell and trend chart.
- Use the underlying series to judge at least the latest direction, whether the latest move is meaningful relative to recent history, and whether the current value is high, low, or neutral versus its own recent range when that comparison is meaningful.
- If the front sheet, cached current value, or chart cache conflicts with the underlying source series, treat the underlying source series as authoritative and fix or flag the cached/front-sheet mismatch before writing the final trend-analysis text.
- Trend-analysis text should briefly describe the recent trend and current signal. Do not mechanically infer the conclusion from a single front-sheet current value, a single MoM value, or chart appearance.
- The conclusion wording must reflect the economic meaning of the indicator. A rising value is not automatically positive, and a falling value is not automatically negative.
- If the underlying data are stale, missing, too sparse, or conflict with supplemental sources, write a cautious neutral/stale-data explanation or leave the prior text unchanged with an audit note; do not invent a trend from incomplete front-sheet data.

## Guardrails

- Do not color unchanged policy-rate anchors merely because their absolute level is high or low if the trailing-one-year series is flat.
- Do not use the TREE tracking-focus column to decide whether the underlying indicator is a YoY indicator. Use the Wind/source indicator name and the official data definition.
- Protect known missing-source rows from being overwritten by false zero values, such as margin-balance rows when Wind returns 0 because of missing data.
- Skip the automatic daily Wind refresh step in normal TREE runs. Only run `refresh_daily.ps1` or other WPS/Wind refresh automation when the user explicitly asks Codex to update the daily workbook.
- When a manually refreshed daily workbook is selected, verify at minimum: file date/version is the latest daily candidate, xlsx integrity is valid, expected source sheets exist, bottom source sheets contain recent nonblank observations, and key front-sheet current values/dates are not obviously stale relative to bottom source sheets.
- After refreshing the daily workbook manually or automatically, expand front-sheet formulas to the current last date column of the base data sheets only when those formulas are needed for TREE data matching or validation. The daily front sheets must not remain capped at an old column such as `TN` when the base data already extends further.
- Skip daily front-sheet trend-chart/cache updating in normal TREE runs. Only run daily chart-cache/trend-chart postprocessing when the user explicitly requests an updated daily workbook with refreshed trend charts or when a non-chart data integrity fix cannot be safely separated from that step.
- Verify at least several front-sheet daily macro rows against the base `宏观数据` sheet. The displayed current value/date should match the latest nonblank base observation for the same code.
- If a coded TREE macro row is not found in the daily base workbook, verify it against the refreshed monthly/supplemental macro source before treating it as unmatched. Do not silently carry forward old values without confirming that the supplemental source has no newer observation.
- A row with one explicit Wind code must match that code exactly. When the daily source does not contain the code, do not fall back to a similarly named series; preserve the verified monthly/supplemental overlay.
- Index PE valuation rows must be verified against the bottom `指数走势` PE block, not only the front-sheet cached formula values. PE current value/date must use the latest PE observation whose date is less than or equal to the corresponding index price date.
- If an index price date is current but the front PE cached date is clearly stale, such as a different year or more than three trading days behind the price date, treat the front PE cache as invalid and resync it from the `指数走势` PE block before updating TREE.
- PE valuation conclusion text and per-indicator judgement text must be regenerated from the same corrected PE values, dates, and three-year percentile data used in TREE current-value cells.
- Trend-analysis text must be regenerated or verified from the same underlying source series used for TREE data and trend charts; do not use front-sheet cached data alone as the basis for the text conclusion.
- After every update, audit at least representative rows for current value, MoM, YoY, YoY increment, color fills, date formats, formula errors, and trend-chart anchors.
- Run the standalone macro-color normalization as the final TREE step so inherited styles cannot change the intended historical-percentile color intensity.
- Weight styling is strict: weights greater than or equal to 2% have both the indicator name and weight shown bold red; weights below 2% are regular black.
