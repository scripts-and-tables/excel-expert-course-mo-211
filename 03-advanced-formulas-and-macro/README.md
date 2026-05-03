![MO-211 Chapter 3](../src/hero-c3.png)

# Module 03: Create Advanced Formulas and Macros (25–30%)

This module is the technical engine of the MO-211 exam. You will move beyond basic calculations to build "intelligent" spreadsheets that handle complex business logic, perform predictive financial modeling, and automate repetitive tasks. We cover the transition from legacy functions to the modern **Dynamic Array engine** (`FILTER`, `SORTBY`, `UNIQUE`) and the professional use of the **`LET()`** function to write clean, high-performance code.

Throughout this chapter, you will learn to "reverse-engineer" solutions using **What-If Analysis** (Goal Seek and Scenario Manager) and master the **Financial Suite**—calculating everything from monthly loan payments (`PMT`) to irregular investment returns (`XIRR`). Because complex formulas are prone to errors, we also dedicate a section to **Formula Auditing**, teaching you to use Tracer Arrows, the Watch Window, and the Evaluate Formula tool to debug your work like a developer.

Finally, we introduce **VBA Automation** through simple Macros. You will learn how to record your manual workflows, assign them to buttons, and perform "surgical" edits to the code in the VBA Editor. By the end of this module, you won't just be entering data; you will be building automated analytical systems.

---

## 📂 Module Contents

### [3.1 Logical Operations](./3.1-logical-operation.md)
* Nested `IF`, `IFS`, `SWITCH`, and the logic group (`SUMIFS`, `MAXIFS`, etc.).
* Using `LET()` to define variables and improve formula speed.

### [3.2 Lookup Data](./3.2-lookup-data.md)
* Mastering `XLOOKUP` (the modern standard) vs. legacy `VLOOKUP`/`HLOOKUP`.
* Advanced `INDEX` & `MATCH` combinations for 2D lookups.

### [3.3 Date and Time Functions](./3.3-date-and-time-funcs.md)
* `WEEKDAY`, `WORKDAY`, and the "Ghost Change" effect of volatile functions (`NOW`, `TODAY`).
* The **EOMONTH** shift cheatsheet for reporting boundaries.

### [3.4 What-If Analysis](./3.4-what-if-analysis.md)
* Using **Goal Seek** to find specific targets.
* Comparing Best/Worst cases with **Scenario Manager** and Summary Reports.
* **Data Tables** (one- and two-variable) for sensitivity analysis.

### [3.5 Dynamic Arrays](./3.5-arrays.md)
* The "Spill" engine: `FILTER`, `SORTBY`, `UNIQUE`, `SORT`, `SEQUENCE`, `VSTACK`, `HSTACK`.
* Text-splitting family: `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTJOIN`.
* Referencing ranges with the Spilled Range Operator (`#`).

### [3.6 Financial Analysis](./3.6-financial-analysis.md)
* Loan modeling: `PMT`, `NPER`, `RATE`, `IPMT`, and `PPMT`.
* Business valuation: **XNPV** and **XIRR** for irregular cash flows.

### [3.7 Troubleshooting Formulas](./3.7-troubleshoot-formulas.md)
* Visualizing webs with **Trace Precedents** and **Dependents**.
* Monitoring distant cells with the **Watch Window** and stepping through logic with **Evaluate Formula**.
* Full error-type reference and the `IFERROR` / `IFNA` patterns.

### [3.8 Macro Automation](./3.8-macro.md)
* Recording, naming, and running simple macros.
* Saving as `.xlsm`, editing recorded VBA, and the **Personal Macro Workbook**.
* Form Controls (Button, Check Box, Spin Button) and assigning macros to them.

### [3.9 Consolidate](./3.9-consolidate.md)
* Combining data from multiple ranges or workbooks via the Consolidate dialog.

### [3.10 Forecast Sheet](./3.10-forecast-sheet.md)
* Excel's built-in time-series forecasting (`FORECAST.ETS`, `FORECAST.LINEAR`).
* Confidence intervals and seasonality detection.

### [Test Yourself](./3.test-yourself.md)
* Comprehensive practice quiz covering Module 3.

---

## ✅ What You Must Be Able To Do

By the end of this module, you must be able to:

- **Build nested logical formulas** using `IF`, `IFS`, `SWITCH`, `AND`/`OR`, and the `*IFS` aggregation family (`SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`).
- **Use `LET()`** to assign intermediate names inside a formula for clarity and recalculation efficiency.
- **Look up data** with `XLOOKUP` (full six-argument signature, including `match_mode` and `search_mode`), legacy `VLOOKUP`/`HLOOKUP` with wildcards, and `INDEX`/`MATCH` for 2-D and left-lookup scenarios.
- **Calculate dates and business days** with `WEEKDAY`, `WORKDAY`, `WORKDAY.INTL`, `NETWORKDAYS`, and `EOMONTH`, and reason about volatile functions (`NOW`, `TODAY`).
- **Run What-If Analysis**: solve for inputs with **Goal Seek**, compare cases with **Scenario Manager** (including Summary reports), and build **one- and two-variable Data Tables** for sensitivity analysis.
- **Author dynamic-array formulas** — `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `VSTACK`/`HSTACK`, `TAKE`/`DROP`, `CHOOSEROWS`/`CHOOSECOLS`, and the text-split family — and reference results with the spilled-range operator (`#`).
- **Model loans and investments** using `PMT`, `NPER`, `RATE`, `IPMT`, `PPMT`, `XNPV`, and `XIRR` with correct sign convention and rate-period consistency.
- **Audit and troubleshoot formulas** using **Trace Precedents/Dependents**, the **Watch Window**, **Evaluate Formula** (step-through), **Trace Error**, and structured handlers (`IFERROR`, `IFNA`).
- **Record, edit, and run macros**: name them, save the workbook as `.xlsm`, edit the recorded VBA in the editor, deploy reusable macros to the **Personal Macro Workbook**, and assign them to **Form Controls** (Button, Check Box, Spin Button).
- **Consolidate** data from multiple ranges or workbooks via the Consolidate dialog, and **build a Forecast Sheet** (`FORECAST.ETS`/`FORECAST.LINEAR`) with confidence intervals and seasonality.
