# Practice Workbooks

Hands-on `.xlsx` files matched to each module of the course. Open the file, read the **Tasks** sheet, then work through the exercises in real Excel.

## Files

| File | Module | Skills practiced |
| --- | --- | --- |
| [`module-1-workbook-security.xlsx`](./module-1-workbook-security.xlsx) | Module 1 | Document Inspector, Encrypt with Password, Protect Sheet (two-step), Hide & Protect Workbook Structure, Manual calculation mode, Check Accessibility |
| [`module-2-sales-data.xlsx`](./module-2-sales-data.xlsx) | Module 2 | Flash Fill, custom number formats, data validation, Excel Tables, Remove Duplicates, conditional formatting (Data Bars), Advanced Filter, Subtotals |
| [`module-3-finance-and-formulas.xlsx`](./module-3-finance-and-formulas.xlsx) | Module 3 | XLOOKUP, IFS, FILTER, Goal Seek, one-variable Data Tables, formula auditing, recording macros, SUMIFS |
| [`module-4-pivots-and-charts.xlsx`](./module-4-pivots-and-charts.xlsx) | Module 4 | PivotTables, Slicers, Timelines, Calculated Fields, Pareto chart, Sparklines, Trendlines, Waterfall |

## How to use

1. Download or clone the repo. The `.xlsx` files live in `practice/`.
2. Open one in Microsoft Excel (Microsoft 365 / Excel 2021 or newer recommended).
3. Read the **Tasks** sheet — it lists the exercises with the skill being tested and a hint.
4. Try each task without looking at the hint first. Use the lessons in the matching module folder (`01-…/`, `02-…/`, …) to refresh anything you forget.
5. When you're stuck, check the hint. When you're done, compare against a solution by re-running the task on a copy.

> [!TIP]
> If a task asks you to record a macro or run something macro-related, save the workbook as `.xlsm` first — `.xlsx` discards macros silently.

## Want to regenerate these?

The files are produced by `scripts/build_practice.py` (kept locally during course generation; the data is seeded for reproducibility). If you want different sample data, edit the seeds and regenerate.
