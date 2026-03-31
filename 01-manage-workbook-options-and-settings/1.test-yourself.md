# Module 1.1 Test: Managing Workbooks & Macros
**Course:** MO-211: Microsoft Excel Expert  
**Topic:** Copying Macros between Workbooks

---

## Part 1: Multiple Choice
*Choose the best answer for each scenario.*

### Q1. Practical Application
You are working in the Visual Basic Editor (VBE) and need to move a macro from `Report_A.xlsm` to `Report_B.xlsm`. You decide to use the Drag-and-Drop method. Which specific item do you click and drag in the Project Explorer?
- A) The `VBAProject (Report_A.xlsm)` folder icon.
- B) The specific `Sub MacroName()` line in the code window.
- C) The specific Module name (e.g., `Module1`) within the Modules folder.
- D) The `Microsoft Excel Objects` folder.

### Q2. File Extensions
You have successfully copied a complex automation script into a new, blank workbook. You save the file as `Final_Analysis.xlsx`. When you reopen the file the next morning, the macro is gone. Why?
- A) The VBE was not closed before saving the file.
- B) Standard `.xlsx` files do not support VBA storage and strip macros upon saving.
- C) The macro was not "compiled" before the file was closed.
- D) The file was saved in a "Read-Only" directory.

### Q3. Individual Script Management
A colleague sends you a workbook with a module containing 20 different macros. You only need **one** specific macro for your project. What is the most efficient way to get only that code into your workbook?
- A) Drag-and-Drop the entire module and delete the 19 macros you don't need.
- B) Export the module as a `.bas` file and import it.
- C) Open the code window, highlight the specific `Sub` to `End Sub` block, copy it, and paste it into a module in your workbook.
- D) Use the "Share Macro" feature in the Developer Tab.

### Q4. Naming Conflicts
You drag `Module1` from a source file into a destination file that already contains a module named `Module1`. How does Excel resolve this?
- A) It overwrites the existing code in the destination file.
- B) It automatically renames the new module to `Module11`.
- C) It merges the code from both modules into one.
- D) It displays an error message and cancels the transfer.

---

## Part 2: True or False

**Q5.** To see the "Developer" tab on the Ribbon, you must first enable it in the Excel Options or by right-clicking the Ribbon and selecting "Customize the Ribbon."  
*(True / False)*

**Q6.** The `.bas` file format is used when exporting a module to be saved as a standalone file on your computer.  
*(True / False)*

**Q7.** You can drag an individual `Sub-routine` (macro) directly from the code window of one workbook into the Project Explorer of another.  
*(True / False)*

---

## Part 3: Answer Key

| Question | Correct Answer | Explanation |
| :--- | :--- | :--- |
| **Q1** | **C** | Drag-and-drop works at the **Module** level, not the individual script or project level. |
| **Q2** | **B** | This is a common exam "gotcha." Always use `.xlsm` or `.xlsb` to keep your code. |
| **Q3** | **C** | While drag-and-drop is fast for modules, manual Copy/Paste is the only way to isolate a single script. |
| **Q4** | **B** | Excel appends a number to the end of the module name to avoid overwriting existing data. |
| **Q5** | **True** | The Developer tab is hidden by default in standard Excel installations. |
| **Q6** | **True** | `.bas` stands for Basic (as in Visual Basic) and is the standard export format for modules. |
| **Q7** | **False** | You can only drag **Modules** in the Project Explorer. To move a Sub-routine, you must copy/paste text. |

---
**Tip for the Exam:** If the question asks for the "fastest" way to move code between two *open* workbooks, always look for the **Drag-and-Drop** option in the VBE.








# Module 1.2 Test: External References
**Course:** MO-211: Microsoft Excel Expert  
**Topic:** Referencing Data in Other Workbooks

---

## Part 1: Multiple Choice
*Choose the best answer for each scenario.*

### Q1. Creating Large-Scale Links
You need to link a summary table (Range A1:G50) from a source workbook into a destination report. Which method is most efficient for transferring this entire block while ensuring Excel creates all 350 individual links automatically?
- A) Typing the manual syntax for each cell.
- B) Using the Point-and-Click method for the first cell and dragging the fill handle.
- C) Copying the range in the Source and using **Paste Link** in the Destination.
- D) Using the "Consolidate Data" tool in the Data tab.

### Q2. Managing Broken Links
A source workbook has been moved to a different folder on the company server, causing your report formulas to return `#REF!` errors. Which button in the **Edit Links** dialog box should you use to point Excel to the file's new location?
- A) Update Values
- B) Check Status
- C) Change Source
- D) Open Source

### Q3. Syntax Identification
Which of the following formulas correctly represents a link to a workbook named **Expenses 2026** that is currently **closed**?
- A) `=[Expenses 2026.xlsx]Sheet1!$A$1`
- B) `='C:\Users\Admin\Documents\[Expenses 2026.xlsx]Sheet1'!$A$1`
- C) `=Expenses 2026.xlsx!Sheet1!$A$1`
- D) `='[Expenses 2026.xlsx]Sheet1'!$A$1`

### Q4. Security and Startup
You want your automated dashboard to pull the latest data from linked files silently every time it opens, without showing a yellow warning bar to the end-user. Which **Startup Prompt** setting should you choose?
- A) Let users choose
- B) Don't display the alert and don't update
- C) Don't display the alert and update links
- D) Disable all external references in Trust Center

---

## Part 2: True or False

**Q5.** When using the "Point-and-Click" method, you must type the `=` sign in the destination cell before switching to the source workbook.  
*(True / False)*

**Q6.** Breaking a link via the "Edit Links" dialog can be reversed by pressing `Ctrl + Z` (Undo) if you realize you made a mistake.  
*(True / False)*

**Q7.** In an external reference formula, square brackets `[ ]` are used specifically to enclose the Sheet Name.  
*(True / False)*

---

## Part 3: Answer Key

| Question | Correct Answer | Explanation |
| :--- | :--- | :--- |
| **Q1** | **C** | **Paste Link** is the fastest official method for linking large ranges or tables in one action. |
| **Q2** | **C** | **Change Source** is the standard tool to "re-path" a link to a moved or renamed file. |
| **Q3** | **B** | If a file is **closed**, Excel requires the **full file path** and **single quotes** for names with spaces. |
| **Q4** | **C** | This setting allows for seamless, "silent" updates in professional dashboards. |
| **Q5** | **True** | The `=` tells Excel you are starting a formula; without it, clicking the other book won't create a link. |
| **Q6** | **False** | Breaking a link is permanent. It "flattens" the formula into a static value and cannot be undone. |
| **Q7** | **False** | Square brackets `[ ]` enclose the **Workbook Name**. The Sheet Name follows the brackets. |

---
**Exam Tip:** On the MO-211, if you are asked to "Break all links in the workbook," remember that you must go to **Data > Edit Links**, select all sources in the list, and then click **Break Link**.



# Module 1.3 Test: Macro Security & Trust Center
**Course:** MO-211: Microsoft Excel Expert  
**Topic:** Enabling Macros and Managing Trust Settings

---

## Part 1: Multiple Choice
*Choose the best answer for each scenario.*

### Q1. Identifying the Default State
When you open a macro-enabled workbook (`.xlsm`) for the first time on a standard Excel installation, what is the default behavior?
- A) The macros run immediately without any prompts.
- B) Excel displays a yellow Message Bar stating "SECURITY WARNING: Macros have been disabled."
- C) Excel permanently deletes the VBA code for security.
- D) A pop-up window forces you to password-protect the workbook.

### Q2. Global Security Levels
You are working in a highly restricted corporate environment where IT policy forbids any unverified code from running. Which Macro Setting in the Trust Center allows only verified, internal company tools to function?
- A) Disable all macros without notification.
- B) Disable all macros with notification.
- C) Disable all macros except digitally signed macros.
- D) Enable all macros.

### Q3. Managing Workflow Efficiency
You have a specific folder (`C:\CompanyReports\`) where you store 50 macro-enabled files. You want to stop the "Enable Content" prompt for every file in this directory. What is the most "Expert" solution?
- A) Change the global setting to "Enable all macros."
- B) Add the `CompanyReports` folder as a **Trusted Location** in the Trust Center.
- C) Move all files to the "Downloads" folder.
- D) Re-save all files as standard `.xlsx` workbooks.

### Q4. Advanced Developer Settings
You are using a specialized Add-in that programmatically modifies VBA code within your workbooks. The Add-in fails to function. Which specific setting in the Macro Settings menu must be enabled?
- A) Disable Excel 4.0 macros when VBA macros are developed.
- B) Trust access to the VBA project object model.
- C) Require Trusted Publishers for all Add-ins.
- D) Enable all macros.

---

## Part 2: True or False

**Q5.** Clicking "Enable Content" on the yellow Message Bar makes that specific workbook a "Trusted Document" on your computer.  
*(True / False)*

**Q6.** For maximum security, it is considered a "Best Practice" to set your computer's "Downloads" folder as a Trusted Location.  
*(True / False)*

**Q7.** If you rename or move a "Trusted Document" to a new folder, Excel will likely ask you to enable the content again.  
*(True / False)*

---

## Part 3: Answer Key

| Question | Correct Answer | Explanation |
| :--- | :--- | :--- |
| **Q1** | **B** | "Disable all macros with notification" is the Excel default to protect users while allowing them to opt-in. |
| **Q2** | **C** | Digitally signed macros use a certificate to prove the code comes from a safe, verified source. |
| **Q3** | **B** | **Trusted Locations** allow all files in a specific directory to bypass security prompts automatically. |
| **Q4** | **B** | This setting is required for any tool or script that needs to "read" or "write" to the VBA Project itself. |
| **Q5** | **True** | Excel remembers your choice for that specific file path so you aren't prompted every time you open it. |
| **Q6** | **False** | Never trust the Downloads folder; it is the primary entry point for malicious files from the internet. |
| **Q7** | **True** | Trust is tied to the specific file path. If the path or name changes, the "Trusted Document" status is reset. |

---
**Exam Tip:** On the MO-211, you may be asked to "Configure Excel to trust all subfolders of the 'Project' directory." Remember to check the **"Subfolders of this location are also trusted"** box when adding the location in the Trust Center.
