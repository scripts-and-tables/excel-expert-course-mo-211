![MO-211 Chapter 1](../src/hero-c1.png)
# Module 01: Manage Workbook Options and Settings (10–15%)

This module focuses on the infrastructure, security, and connectivity of Microsoft Excel workbooks. As an Excel Expert, you aren't just filling cells; you are managing the environment in which data lives. We begin by mastering the movement of automation through the **Visual Basic Editor (VBE)** and establishing robust data pipelines by **referencing external workbooks** to create a "Single Source of Truth."

You will also learn to navigate the high-stakes world of **Workbook Security and Governance**. This involves mastering the **Trust Center** to safely enable macros, implementing multi-layered protection—ranging from **cell-level locking** to **workbook structure constraints**—and ensuring data integrity through **version management**. Finally, we optimize performance for large-scale datasets by fine-tuning **calculation options**, including manual triggers and iterative logic, ensuring your models remain responsive even under heavy data loads.

-----

## 📂 Module Contents

### [1.1 Copy Macros](./1.1-copy-macro.md) 

  * Navigating the Visual Basic Editor (VBE) and Project Explorer.
  * Moving and copying modules between workbooks via drag-and-drop or `.bas` export/import.

### [1.2 Reference Data](./1.2-reference-data.md)

  * Creating external references to other workbooks and understanding path syntax.
  * Managing, updating, and breaking links via the Edit Links dialog.

### [1.3 Enable Macros](./1.3-enable-macro.md)

  * Configuring Trust Center settings and Macro Security levels.
  * Setting up **Trusted Locations** to bypass security warnings for known folders.

### [1.4 Workbook Versions](./1.4-workbook-versions.md)

  * Recovering unsaved work using the AutoRecover engine.
  * Managing, viewing, and restoring previous file states through **Version History**.

### [1.5–1.6 Protect Workbooks: Soft Restrictions, Worksheet Protection, and Cell Ranges](./1.6-protect-worksheet-and-ranges.md)

  * "Soft" restrictions: **Mark as Final**, **Always Open Read-Only**, **Encrypt with Password**, and **Information Rights Management (IRM)**.
  * The "Two-Step" worksheet protection workflow: unlocking input cells and protecting the sheet.
  * Hiding proprietary formulas and using **Allow Edit Ranges** for per-user/per-range permissions.

### [1.7 Protect Workbook Structure](./1.7-protect-workbook-structure.md)

  * Locking the workbook "container" to prevent deleting, renaming, or moving tabs.
  * Strategic use of hidden sheets to protect background mapping tables.

### [1.8 Formula Calculation Options](./1.8-formula-calculation-options.md)

  * Toggling between **Manual and Automatic** calculation for performance optimization.
  * Enabling **Iterative Calculations** and setting **Precision as Displayed** for financial accuracy.

### [1.9 Inspect Workbook](./1.9-inspect-and-check.md)

  * **Document Inspector** — find and remove hidden data (comments, hidden sheets, custom XML, embedded files).
  * **Check Accessibility** — alt text, header rows, contrast, sheet names.
  * **Check Compatibility** — flag features that won't survive a save to older formats.

### [Test Yourself](./1.test-yourself.md)

  * Comprehensive practice quiz covering all of Module 1.

-----

## ✅ What You Must Be Able To Do

By the end of this module, you must be able to:

- **Copy macros and modules between workbooks** using the Visual Basic Editor's Project Explorer (drag-and-drop) or `.bas` export/import.
- **Build external workbook references** with correct path syntax and **manage links** through the Edit Links dialog (update values, change source, break link).
- **Configure Trust Center settings** — macro security levels, **Trusted Locations**, and Trusted Publishers — to enable macros without disabling protection workbook-wide.
- **Recover unsaved work** with AutoRecover and **restore previous file states** through Version History on OneDrive/SharePoint.
- **Apply layered workbook protection**: Mark as Final, Always Open Read-Only, password encryption, and **Information Rights Management (IRM)**.
- **Protect worksheets and cell ranges** using the unlock-then-protect workflow, including **Allow Edit Ranges** with per-user permissions and hidden formulas.
- **Lock workbook structure** to prevent users from inserting, deleting, renaming, moving, or unhiding sheets.
- **Toggle calculation modes** (Automatic, Automatic Except Tables, Manual) and enable **Iterative Calculation** and **Precision as Displayed** when models require them.
- Run **Document Inspector** to find and remove hidden data (comments, hidden sheets, custom XML, embedded files) before sharing.
- Run **Check Accessibility** and **Check Compatibility** to flag alt-text, contrast, and version-incompatible features prior to distribution.
