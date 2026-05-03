# Practice Workbooks

Hands-on `.xlsx` files matched to each module of the course. Open the file, read the **Tasks** sheet, then work through the exercises in real Excel.

> Build script (kept locally, not committed): `/tmp/build_practice_v2.py`. It reads each existing workbook, appends new tasks for Module 2, adds a hidden `Solutions` sheet to every file, and writes them back here.

## Files

| File | Module | Skills practiced |
| --- | --- | --- |
| [`module-1-workbook-security.xlsx`](./module-1-workbook-security.xlsx) | Module 1 | Document Inspector, Encrypt with Password, Protect Sheet (two-step), Hide & Protect Workbook Structure, Manual calculation mode, Check Accessibility |
| [`module-2-sales-data.xlsx`](./module-2-sales-data.xlsx) | Module 2 | Flash Fill, custom number formats, data validation, Excel Tables, Remove Duplicates, conditional formatting (Data Bars), Advanced Filter, Subtotals, **Named Ranges**, **Power Query (filter/add column/group)**, **Linked Data Types (Geography)** |
| [`module-3-finance-and-formulas.xlsx`](./module-3-finance-and-formulas.xlsx) | Module 3 | XLOOKUP, IFS, FILTER, Goal Seek, one-variable Data Tables, formula auditing, recording macros, SUMIFS |
| [`module-4-pivots-and-charts.xlsx`](./module-4-pivots-and-charts.xlsx) | Module 4 | PivotTables, Slicers, Timelines, Calculated Fields, Pareto chart, Sparklines, Trendlines, Waterfall |

## How to use

1. Download or clone the repo. The `.xlsx` files live in `practice/`.
2. Open one in Microsoft Excel (Microsoft 365 / Excel 2021 or newer recommended).
3. Read the **Tasks** sheet — it lists the exercises with the skill being tested and a hint.
4. Try each task without looking at the hint first. Use the lessons in the matching module folder (`01-…/`, `02-…/`, …) to refresh anything you forget.
5. When you're stuck, check the hint, then run the self-check below.

> [!TIP]
> If a task asks you to record a macro or run something macro-related, save the workbook as `.xlsm` first — `.xlsx` discards macros silently.

## Self-grading: the hidden Solutions sheet

Every practice workbook ships with a **`Solutions`** sheet that is **hidden by default** so you aren't tempted to peek. Each Solutions sheet lists, per task:

- the task number,
- the expected formula or step-by-step procedure,
- the cell range where the result should land,
- a one-line verification note ("Cell X should show Y after applying Z").

To reveal it: **right-click any sheet tab → Unhide → Solutions → OK**. Re-hide it the same way (right-click the `Solutions` tab → Hide) when you want a clean re-attempt.

### Self-grading checklist pattern

For each task, work through this loop before you peek:

1. **Read** the Tasks-sheet row (Task + Skill tested only — cover the Hint with your hand if you must).
2. **Attempt** the task in the data sheet.
3. **Predict** the verification: what cell, what value, what should happen?
4. **Unhide Solutions**, find the matching `#`, and compare your result against the Verification column.
5. **Re-hide Solutions** and move on. If you didn't match, redo the task before continuing — muscle memory beats reading.

## Want to regenerate these?

The files are produced by `/tmp/build_practice_v2.py` (kept locally during course generation; the data is seeded for reproducibility). If you want different sample data or extra tasks, edit that script and rerun it — it reads the existing workbooks in place, so existing data sheets are preserved.
