# Mock Exam — MO-211 Excel Expert

A capstone, exam-style run that simulates the live MO-211 experience: **35 timed tasks across all four modules**, weighted to match Microsoft's Skills Measured blueprint.

## Setup

1. Open all four practice workbooks in `practice/` so you can switch between them quickly.
2. Set a stopwatch for **50 minutes** — the same length as the live exam.
3. Close every other tab and silence notifications. The real exam runs in a sandbox; eliminate distractions to mirror that.
4. Do not open the lessons or the answer key during the run. Mark anything you can't immediately solve and come back to it.

> [!IMPORTANT]
> The live MO-211 exam is performance-based: you build the solution inside a sandboxed copy of Excel, not by picking from multiple-choice options. The tasks below mimic that style — each asks you to perform a specific change in a real workbook.

---

## Section 1 — Manage Workbook Options & Settings (10–15%)

> Open `practice/module-1-workbook-security.xlsx`.

1. Encrypt the workbook with the password `mo211mock` (File → Info → Protect Workbook → Encrypt with Password).
2. On **Customer Quote**, unlock cells `B3:B7` (Format Cells → Protection → uncheck Locked), then protect the sheet with no password and the default user permissions.
3. Hide the formulas in cells `B9:B11` (Format Cells → Protection → Hidden) so they don't appear in the formula bar after protection. Verify by clicking `B9` after Protect Sheet.
4. Set the workbook to **Manual** calculation (Formulas → Calculation Options → Manual). Document why you'd do this for a 50,000-row model.
5. Run **File → Info → Check for Issues → Inspect Document** and remove **Comments and Annotations** + **Hidden Worksheets** (if any). Keep document properties.

---

## Section 2 — Manage and Format Data (30–35%)

> Open `practice/module-2-sales-data.xlsx`.

6. Define a **named range** called `NorthTarget` for the cell on the Targets sheet that holds the North region's quarterly target. Use it in a formula on `Raw Sales` row 2 to flag rows over target.
7. On **Raw Sales**, use **Flash Fill** (`Ctrl+E`) to extract the customer first name from `Full Name` into a new column.
8. Apply a **custom number format** to the `Amount` column that shows positive values in **green** with the `$` symbol and 2 decimals; negative in **red** with parentheses; zero as a dash.
9. Convert `Raw Sales` to an **Excel Table** (`Ctrl+T`), name it `tblSales`, and add a **Total Row** that sums the Amount column.
10. Add **Data Validation** on the `Region` column to restrict input to a list pulled from a named range. Use `INDIRECT` to make the city dropdown depend on the chosen region.
11. Use **Advanced Filter** to copy unique customer names from `Raw Sales` to a new location, with a criteria range that requires Amount > 1000.
12. In Power Query, load the Targets sheet **From Table/Range**, filter Status ≠ "Cancelled", add a custom column `Revenue = Quantity * Price`, then group by Region.
13. Convert cells `M1:M5` to **Geography** linked data type and pull the Population field into column N using the dot operator (`=M1.Population`).

---

## Section 3 — Advanced Formulas and Macros (25–30%)

> Open `practice/module-3-finance-and-formulas.xlsx`.

14. On **Sales**, add an **`XLOOKUP`** that returns the product name for each Product ID by looking up against the Products sheet. Handle missing matches with the `if_not_found` argument.
15. Create a 2D lookup using `INDEX(MATCH, MATCH)` that returns the price for a chosen Product × Region pair.
16. On **Loan Model**, use **`PMT`** to find the monthly payment for the rate/term/principal already entered. Apply the correct sign convention.
17. Use **`NPER`** to determine how many months it takes to pay off `$15,000` at 9% APR with a $400 monthly payment.
18. Use **`IPMT`** and **`PPMT`** to split month 1 of the Loan Model into interest and principal.
19. Use **Goal Seek** (Data → What-If → Goal Seek) to find the loan term that produces a $1,500 monthly payment for the existing principal and rate.
20. Build a **one-variable Data Table** on Loan Model showing the monthly payment for terms of 10, 15, 20, 25, 30 years.
21. On **Broken Formulas**, use **Trace Precedents** and **Evaluate Formula** to identify why the formula in `B5` returns `#N/A`. Wrap it in `IFNA` with a sensible fallback.
22. Record a simple macro that selects the Sales sheet, sorts the data by Amount descending, and applies a Currency format to the Amount column. Name it `cleanSales` and store it in **This Workbook**.

---

## Section 4 — Manage Advanced Charts and Tables (25–30%)

> Open `practice/module-4-pivots-and-charts.xlsx`.

23. From **Quarterly Sales**, build a **PivotTable** on a new sheet with Region in Rows, Product in Columns, and Sum of Amount in Values.
24. Add a **Slicer** for Region and a **Timeline** for Date connected to the PivotTable. Set the timeline level to **Months**.
25. Add a **Calculated Field** to the PivotTable named `Margin %` defined as `Profit / Amount`. Format as percentage with 1 decimal.
26. Group dates in the PivotTable by **Quarters and Years** instead of individual dates.
27. Build a **PivotChart** (clustered column) from the PivotTable with field buttons hidden.
28. From **Defects**, create a **Pareto** chart (Insert → Statistical Charts → Histogram → Pareto) showing defect counts sorted descending with the cumulative line.
29. Build a **Waterfall** chart from a P&L breakdown (Revenue, COGS, Gross Profit, OpEx, Net Income). Right-click the Net Income bar → **Set as Total**.
30. Insert **Line sparklines** in column F of the Trend sheet, one per row of monthly data. Show the High Point in red.
31. Add a **Linear trendline** to the column chart on the Trend sheet with **Forward 3 periods** and **Display R-squared** ticked.
32. From the Quarterly Sales PivotTable, double-click any value cell to **Show Details** (drill-through). Confirm a new sheet opens with the underlying source rows for that cell.

---

## Section 5 — Cross-cutting (any module)

33. Use **`AGGREGATE`** option `7` to sum a range while ignoring both hidden rows and `#N/A` values.
34. Use **`TEXTSPLIT`** to split a comma-and-space-delimited list of cities in cell `A1` into separate cells across one row.
35. Use **`NETWORKDAYS.INTL`** to count working days between two dates with a Friday/Saturday weekend (relevant for Middle East calendars). Pass `7` as the weekend code.

---

## Self-grading

After your 50 minutes are up:

1. Open each practice workbook's **Solutions** sheet (right-click any sheet tab → Unhide → Solutions).
2. Compare your answer for each task against the verification note.
3. Score yourself: 1 point for fully correct, 0.5 for partial, 0 for missed. Pass ≈ 70% (Microsoft's actual exam pass mark is 700/1000, scaled).
4. Re-do every task you missed — muscle memory beats reading.

> [!TIP]
> If you finish in under 35 minutes you are in great shape for the live exam. If you went over the 50-minute mark, find which section ate the most time and drill the corresponding lessons before booking the exam.

## Where to go next

- **Reset progress**: in the website sidebar, double-click the progress bar to clear the localStorage checkmarks and start a fresh run-through.
- **Schedule the real exam**: see [`EXAM-DAY-TIPS.md`](./EXAM-DAY-TIPS.md) for the test-environment walk-through and last-minute strategy.
- **Drill weak spots**: revisit the relevant module's `*.test-yourself.md` quiz for focused review.
