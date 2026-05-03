![MO-211 Chapter 2](../src/hero-c2.png)
# Module 02: Manage and Format Data (30–35%)

This module is the heaviest on the MO-211 exam — and for good reason. In any real-world data role, **raw data is never clean, never consistently formatted, and never immediately ready for analysis**. This module gives you the tools to transform messy, inconsistent datasets into reliable, presentation-ready information.

We begin with **smart data entry techniques**: using Flash Fill to instantly reformat data by example, the Fill Series engine for generating sequences and projections, and the `RANDARRAY()` function to produce dynamic numeric datasets. We then move into **formatting and validation**, where you will enforce data quality through custom number formats, data validation rules, and structured grouping — all while using subtotals to extract meaningful summaries without restructuring your data. The module closes with **advanced conditional formatting**, where formula-driven rules allow you to build dynamic, self-updating visual dashboards directly inside the grid.

-----

## 📂 Module Contents

### [2.0 Named Ranges](./2.0-named-ranges.md)

  * Three creation paths: Name Box, **Define Name** dialog (with workbook vs. worksheet scope), and **Create from Selection**.
  * **Name Manager** (Ctrl+F3) operations — edit, delete, filter — and the **F3 Paste Name** dialog.
  * Naming rules and using names in formulas, validation lists, and conditional formatting.

### [2.1 Flash Fill and Linked Data Types](./2.1-flash-fill.md)

  * Using pattern recognition to split, combine, and reformat text data instantly.
  * Triggering Flash Fill with `Ctrl+E` and guiding it with additional examples.
  * Converting plain text to **Linked Data Types** (Stocks, Geography, Currencies) and pulling live fields with the insert-data icon and dot-operator formulas.

### [2.2 Advanced Fill Series](./2.2-advanced-fill-series.md)

  * Using the Fill Series dialog for Linear, Growth, and Date sequence generation.
  * Configuring Step values, Stop values, and Date units (Day, Weekday, Month, Year).

### [2.3 RANDARRAY()](./2.3-randarray.md)

  * Generating dynamic arrays of random numbers using `RANDARRAY([rows],[columns],[min],[max],[integer])`.
  * Understanding volatile function behavior and when to freeze output with Paste Special.

### [2.4 Custom Number Formats](./2.4-custom-number-formats.md)

  * Building format codes using the `Positive;Negative;Zero;Text` four-section structure.
  * Applying color codes, suffix labels, date/time patterns, and zero-suppression techniques.

### [2.5 Data Validation](./2.5-data-validation.md)

  * Configuring validation rules for whole numbers, decimals, lists, dates, text length, and custom formulas.
  * Setting Input Messages and Error Alerts (Stop, Warning, Information) to guide users.

### [2.6 Group and Ungroup Data](./2.6-group-and-ungroup-data.md)

  * Grouping rows and columns to create collapsible outline levels in a worksheet.
  * Using Auto Outline and understanding the difference between grouping and hiding.

### [2.7 Subtotals and Totals](./2.7-subtotals-and-totals.md)

  * Using `=SUBTOTAL(function_num, ref)` with function numbers 1–11 and 101–111.
  * Inserting automatic subtotals via Data > Outline > Subtotal and removing them cleanly.

### [2.8 Remove Duplicate Records](./2.8-remove-duplicate-records.md)

  * Using the Remove Duplicates tool and selecting which columns define uniqueness.
  * Non-destructive deduplication using `UNIQUE()` and pre-removal identification with Conditional Formatting.

### [2.9 Custom Conditional Formatting Rules and Rule Management](./2.9-custom-conditional-formatting-rules.md)

  * Creating new rules using Highlight Cell Rules, Top/Bottom Rules, Data Bars, Color Scales, and Icon Sets.
  * Setting precise conditions (cell value, specific text, dates, blanks, errors) with custom formats.
  * Using the **Manage Rules** dialog to view, reorder, edit, and delete rules — including **Stop If True** for mutually exclusive formatting conditions.

### [2.10 Formula-Based Conditional Formatting](./2.10-formula-based-conditional-formatting.md)

  * Writing formula rules that evaluate TRUE/FALSE to apply formatting dynamically.
  * Using absolute and relative references correctly to format entire rows based on a single column value.

### [2.12 Excel Tables and Structured References](./2.12-excel-tables-and-structured-references.md)

  * Converting a range to a Table (`Ctrl+T`), naming it, adding Total Rows.
  * Structured-reference syntax: `Sales[@Amount]`, `Sales[#Totals]`, `Sales[#Headers]`.
  * Why Tables matter for slicers, PivotTables, and auto-extending formulas.

### [2.13 Advanced Filter](./2.13-advanced-filter.md)

  * Building criteria ranges with AND/OR logic across rows and columns.
  * In-place filtering vs. copy-to-location, and the unique-records-only flag.

### [2.14 Power Query Basics](./2.14-power-query-basics.md)

  * Get & Transform entry points: **Data → Get Data**, **From Table/Range**.
  * Common transforms: Promote Headers, Change Type, Filter, Split Column, Custom Column, Group By, Pivot/Unpivot.
  * **Merge Queries** (left/right/inner/outer/anti joins) and **Append Queries** workflows.
  * Close & Load options and refresh behavior.

### [Test Yourself](./2.test-yourself.md)

  * Comprehensive practice quiz covering all of Module 2.

-----

## ✅ What You Must Be Able To Do

By the end of this module, you must be able to:

- Define, scope, edit, and delete **named ranges** via the Name Manager and use them inside formulas, lists, and conditional formatting.
- Fill cells using **Flash Fill** and convert text to **Linked Data Types** (Stocks, Geography, Currencies) with the dot-operator formula syntax.
- Generate sequences and projections using **advanced Fill Series** options (Linear, Growth, Date).
- Produce arrays of random numeric data using **`RANDARRAY()`** and manage its volatile nature.
- Create **custom number formats** with the `Positive;Negative;Zero;Text` structure, conditional brackets, and color codes.
- **Configure data validation** rules — including dependent dropdowns via `INDIRECT` — with input messages and error alerts.
- **Group and ungroup** rows and columns to build collapsible worksheet outlines.
- Use the **`SUBTOTAL()`** and **`AGGREGATE()`** functions, and the Subtotal menu feature, to calculate grouped totals that respect filtered/hidden rows.
- **Remove duplicate records** using both the Remove Duplicates tool and the `UNIQUE()` function.
- Create and **manage** custom and formula-based **conditional formatting rules**, including rule priority and Stop If True.
- Convert ranges to **Excel Tables** and use **structured references** (`Sales[@Amount]`, `Sales[#Totals]`, etc.).
- Build **Advanced Filter** criteria ranges (AND/OR logic) and choose in-place vs. copy-to-location.
- Pull and shape data with **Power Query**: Promote Headers, Filter, Split Column, Group By, Pivot/Unpivot, Merge, and Append.
