# Module 03: Create Advanced Formulas and Macros (25–30%)

This module is the technical engine of the MO-211 exam. You will move beyond basic calculations to build "intelligent" spreadsheets that handle complex business logic, perform predictive financial modeling, and automate repetitive tasks. We cover the transition from legacy functions to the modern **Dynamic Array engine** (`FILTER`, `SORTBY`, `VSTACK`) and the professional use of the **`LET()`** function to write clean, high-performance code.

Throughout this chapter, you will learn to "reverse-engineer" solutions using **What-If Analysis** (Goal Seek and Scenario Manager) and master the **Financial Suite**—calculating everything from monthly loan payments (`PMT`) to irregular investment returns (`XIRR`). Because complex formulas are prone to errors, we also dedicate a section to **Formula Auditing**, teaching you to use Tracer Arrows, the Watch Window, and the Evaluate Formula tool to debug your work like a developer.

Finally, we introduce **VBA Automation** through simple Macros. You will learn how to record your manual workflows, assign them to buttons, and perform "surgical" edits to the code in the VBA Editor. By the end of this module, you won't just be entering data; you will be building automated analytical systems.

---

## 📂 Module Contents

### [3.1 Logical Operations & LET](./3.1-logical-logic.md)
* Nested `IF`, `IFS`, `SWITCH`, and the logic group (`SUMIFS`, `MAXIFS`, etc.).
* Using `LET()` to define variables and improve formula speed.

### [3.2 The Lookup Suite](./3.2-lookups.md)
* Mastering `XLOOKUP` (the modern standard) vs. legacy `VLOOKUP`/`HLOOKUP`.
* Advanced `INDEX` & `MATCH` combinations for 2D lookups.

### [3.3 Advanced Dates & Volatility](./3.3-dates-times.md)
* `WEEKDAY`, `WORKDAY`, and the "Ghost Change" effect of volatile functions (`NOW`, `TODAY`).
* The **EOMONTH** shift cheatsheet for reporting boundaries.

### [3.4 What-If Analysis](./3.4-what-if-analysis.md)
* Using **Goal Seek** to find specific targets.
* Comparing Best/Worst cases with **Scenario Manager** and Summary Reports.

### [3.5 Modern Dynamic Arrays](./3.5-array-functions.md)
* The "Spill" engine: `FILTER`, `SORTBY`, `UNIQUE`, and `VSTACK`.
* Referencing ranges with the Spilled Range Operator (`#`).

### [3.6 Financial Analysis](./3.6-financial-analysis.md)
* Loan modeling: `PMT`, `NPER`, `RATE`, `IPMT`, and `PPMT`.
* Business valuation: **XNPV** and **XIRR** for irregular cash flows.

### [3.7 Troubleshooting & Auditing](./3.7-troubleshoot-formulas.md)
* Visualizing webs with **Trace Precedents** and **Dependents**.
* Monitoring distant cells with the **Watch Window** and stepping through logic with **Evaluate Formula**.

### [3.8 Intro to Macros](./3.8-macros.md)
* Recording, naming, and running simple macros.
* Saving as `.xlsm` and performing minor edits in the VBA Editor.
