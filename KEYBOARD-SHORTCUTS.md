# MO-211 Keyboard Shortcuts Cheatsheet ⌨️

The MO-211 Excel Expert exam gives you **50 minutes for ~35 tasks** — roughly 85 seconds each. Reaching for the mouse for things you could do in two keystrokes burns time you do not have. This cheatsheet collects the shortcuts that actually pay off on the exam.

All shortcuts below are for **Excel on Windows (Microsoft 365 Apps)**, the platform the exam runs on. Mac equivalents are not included — the testing environment is Windows.

---

## 🧭 Navigation & Selection

| Shortcut | What it does |
| --- | --- |
| `Ctrl+Home` | Jump to A1 (or top-left of the data region after a freeze) |
| `Ctrl+End` | Jump to the last used cell on the sheet |
| `Ctrl+Arrow` | Jump to the edge of the current data region in that direction |
| `Ctrl+Shift+Arrow` | Extend selection to the edge of the current data region |
| `Ctrl+A` | Select the current region; press again to select the whole sheet |
| `Ctrl+Space` | Select the entire current column |
| `Shift+Space` | Select the entire current row |
| `Ctrl+G` (or `F5`) | Open **Go To** (then `Special…` for blanks, formulas, constants, etc.) |
| `Ctrl+F` | Open **Find** |
| `Ctrl+H` | Open **Replace** |
| `Ctrl+Shift+End` | Select from the active cell to the last used cell |

---

## ✏️ Editing

| Shortcut | What it does |
| --- | --- |
| `F2` | Edit the active cell (cursor at end of contents) |
| `F4` | Toggle absolute reference while editing a formula; in dialogs, redo last action |
| `Alt+Enter` | Insert a line break inside a cell |
| `Ctrl+;` | Insert today's date (static value) |
| `Ctrl+Shift+;` | Insert the current time (static value) |
| `Ctrl+D` | Fill **down** from the cell above |
| `Ctrl+R` | Fill **right** from the cell to the left |
| `Ctrl+Z` | Undo |
| `Ctrl+Y` | Redo |
| `Ctrl+-` | Open the **Delete** cells dialog |
| `Ctrl+Shift++` | Open the **Insert** cells dialog |
| `Ctrl+C` / `Ctrl+X` / `Ctrl+V` | Copy / Cut / Paste |
| `Ctrl+Alt+V` | Paste **Special** (values, formulas, formats, transpose…) |

---

## 🎨 Formatting

| Shortcut | What it does |
| --- | --- |
| `Ctrl+1` | Open the **Format Cells** dialog (the most useful shortcut on the exam) |
| `Ctrl+B` / `Ctrl+I` / `Ctrl+U` | Bold / Italic / Underline |
| `Ctrl+Shift+5` | Apply **Percent** format (no decimals) |
| `Ctrl+Shift+1` | Apply **Number** format with thousands separator and 2 decimals |
| `Ctrl+Shift+4` | Apply **Currency** format ($#,##0.00) |
| `Ctrl+Shift+7` | Apply outline border to the selection |
| `Ctrl+Shift+&` | Apply outline border (same as above on US layout) |
| `Ctrl+Shift+_` | Remove all borders from the selection |
| `Ctrl+5` | Toggle strikethrough |

> Tip: When a task says "format as percent with one decimal," `Ctrl+Shift+5` then `Ctrl+1` to fine-tune is faster than the ribbon every time.

---

## 🧮 Formulas

| Shortcut | What it does |
| --- | --- |
| `F9` | Recalculate all open workbooks; in the formula bar, **evaluate the selected portion** of a formula |
| `Shift+F9` | Recalculate the active sheet only |
| `Ctrl+Alt+F9` | Force a full recalculation of all open workbooks |
| `Ctrl+\`` | Toggle **Show Formulas** view |
| `F4` | While editing a reference, cycle through `A1` → `$A$1` → `A$1` → `$A1` |
| `Ctrl+Shift+Enter` | Enter a legacy CSE array formula (rarely needed in M365 with dynamic arrays) |
| `Alt+=` | Insert **AutoSum** for the selected range |
| `Ctrl+[` | Select all direct **precedents** of the active cell |
| `Ctrl+]` | Select all direct **dependents** of the active cell |
| `F3` | Paste a defined name into a formula |

---

## 📋 Tables, Filters, PivotTables & Charts

| Shortcut | What it does |
| --- | --- |
| `Ctrl+T` | Convert a range into a **Table** (confirms the "My table has headers" dialog) |
| `Ctrl+Shift+L` | Toggle **AutoFilter** on/off |
| `Alt+N`, `V` | Insert **PivotTable** (sequential, not a chord) |
| `F11` | Insert a chart from the selected data on a **new chart sheet** |
| `Alt+F1` | Insert a chart from the selected data on the **same sheet** |
| `Alt+Down` | Open the AutoFilter dropdown on the active header cell |

---

## 🤖 Macros & Developer

| Shortcut | What it does |
| --- | --- |
| `Alt+F11` | Open the **Visual Basic Editor (VBE)** |
| `Alt+F8` | Open the **Macro** dialog (run, edit, delete) |
| `F5` *(in VBE)* | Run the current macro |
| `F8` *(in VBE)* | **Step Into** — execute one line at a time |
| `Ctrl+Break` *(in VBE)* | Halt a running macro |

---

## 🪟 Window & View

| Shortcut | What it does |
| --- | --- |
| `Ctrl+N` | New workbook |
| `Ctrl+O` | Open workbook |
| `Ctrl+S` | Save (use this constantly during the exam) |
| `Ctrl+W` | Close the active workbook |
| `Ctrl+PgDn` | Move to the **next** sheet |
| `Ctrl+PgUp` | Move to the **previous** sheet |
| `Alt+W`, `F` | Open the **Freeze Panes** menu |
| `Ctrl+F1` | Toggle the ribbon (collapse / expand) |
| `Ctrl+Shift+F1` | Toggle full-screen mode (hides ribbon, tabs, status bar) |

---

## 🏋️ Use these in practice

You will not retain any of this from a list. The only way these shortcuts become muscle memory is to **use them while you build the practice workbooks** in this course — every time you would reach for the mouse, stop and use the shortcut instead. After two or three modules, `Ctrl+1`, `Ctrl+T`, `Alt+=`, and `Ctrl+Shift+L` will feel automatic, and you will buy yourself the minutes that turn a borderline score into a clear pass.
