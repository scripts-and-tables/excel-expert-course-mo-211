# 3.1 Perform Logical Operations Using Advanced Functions

## The Role of Logical Functions
In data analytics, logical functions act as the "control flow" of your spreadsheet. They allow your workbook to evaluate data, ask questions (e.g., "Is this sale over $10,000?"), and execute different calculations based on the answer. 

By mastering nested functions and conditional aggregations, you can transform static data dumps into dynamic, automated reporting engines.

---

## Part 1: Core Conditionals & Boolean Logic
These functions evaluate whether a condition is TRUE or FALSE and allow you to chain multiple criteria together.

### `IF()`
* **What it does:** Returns one value if a condition is true and another value if it's false.
* **When to use it:** For simple, binary decision-making (e.g., assigning a "Pass" or "Fail" grade based on a test score).
* **Official Docs:** [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2)

### `AND()`
* **What it does:** Returns TRUE only if **all** provided arguments are true.
* **When to use it:** To strictly enforce multiple criteria. Often nested inside an `IF` statement (e.g., checking if an employee met their sales target `AND` completed their training).
* **Official Docs:** [AND function](https://support.microsoft.com/en-us/office/and-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9)

### `OR()`
* **What it does:** Returns TRUE if **any** of the provided arguments are true.
* **When to use it:** To broaden criteria. Useful when an item falls into multiple acceptable categories (e.g., flagging an invoice if the region is "North" `OR` "East").
* **Official Docs:** [OR function](https://support.microsoft.com/en-us/office/or-function-7d17ad14-8700-4281-b308-00b131e22af0)

### `NOT()`
* **What it does:** Reverses the logic of its argument (changes TRUE to FALSE, and vice versa).
* **When to use it:** To exclude specific scenarios without writing complex logic (e.g., `IF(NOT(Status="Cancelled"), ...)`).
* **Official Docs:** [NOT function](https://support.microsoft.com/en-us/office/not-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77)

---

## Part 2: Advanced Branching (Replacing Nested IFs)
Deeply nested `IF` statements are notoriously difficult to read and troubleshoot. These modern functions streamline complex branching logic, acting much like a `CASE WHEN` statement in SQL.



### `IFS()`
* **What it does:** Evaluates multiple conditions and returns a value corresponding to the **first** TRUE condition.
* **When to use it:** When you have a sequential list of conditions (e.g., categorizing test scores: >90 is A, >80 is B, >70 is C). It eliminates the need to nest multiple `IF` functions inside each other.
* **Official Docs:** [IFS function](https://support.microsoft.com/en-us/office/ifs-function-36329a26-37b2-467c-972b-4a39bd951d45)

### `SWITCH()`
* **What it does:** Evaluates an expression against a list of exact values and returns the result corresponding to the first matching value.
* **When to use it:** When mapping specific codes to text (e.g., if Column A is "1", return "Standard"; if "2", return "Express"; if "3", return "Overnight"). It is cleaner than `IFS` for exact matches.
* **Official Docs:** [SWITCH function](https://support.microsoft.com/en-us/office/switch-function-47ab33c0-28ce-4530-8a45-d532ec4aa25e)

---

## Part 3: Conditional Aggregation (Single Criteria)
These functions perform math only on rows that meet a specific, single condition.

### `SUMIF()`
* **What it does:** Adds the values in a range that meet the criteria you specify.
* **When to use it:** To find the total revenue for a single specific product line.
* **Official Docs:** [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b)

### `AVERAGEIF()`
* **What it does:** Calculates the average (arithmetic mean) of cells that meet a given criteria.
* **When to use it:** To find the average transaction size for a specific salesperson.
* **Official Docs:** [AVERAGEIF function](https://support.microsoft.com/en-us/office/averageif-function-faec8e2e-0dec-4308-af69-f5576d8ac642)

### `COUNTIF()`
* **What it does:** Counts the number of cells within a range that meet a single condition.
* **When to use it:** To count how many transactions occurred on a specific date or how many times a customer's name appears.
* **Official Docs:** [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34)

---

## Part 4: Conditional Aggregation (Multiple Criteria)
In the real world, you rarely filter by just one thing. The "IFS" family of aggregation functions is robust and requires the calculation range to be listed *first*, followed by pairs of criteria ranges and criteria.

### `SUMIFS()`
* **What it does:** Adds all of its arguments that meet multiple criteria.
* **When to use it:** To calculate total revenue for a specific product line **AND** within a specific region.
* **Official Docs:** [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b)

### `AVERAGEIFS()`
* **What it does:** Calculates the average of cells that meet multiple criteria.
* **When to use it:** To find the average delivery time for "Express" shipments **AND** during the month of "December".
* **Official Docs:** [AVERAGEIFS function](https://support.microsoft.com/en-us/office/averageifs-function-48910c45-1fc0-4389-a028-f7c5c3001690)

### `COUNTIFS()`
* **What it does:** Counts the number of cells that meet multiple criteria.
* **When to use it:** To count how many invoices are "Overdue" **AND** belong to "Client X".
* **Official Docs:** [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842)

### `MAXIFS()`
* **What it does:** Returns the maximum value among cells specified by a given set of conditions or criteria.
* **When to use it:** To find the highest single sale amount recorded for a specific quarter.
* **Official Docs:** [MAXIFS function](https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-32c7-4336-ac8c-f5ce315c8bf8)

### `MINIFS()`
* **What it does:** Returns the minimum value among cells specified by a given set of conditions or criteria.
* **When to use it:** To find the lowest valid bid price from a list of approved vendors.
* **Official Docs:** [MINIFS function](https://support.microsoft.com/en-us/office/minifs-function-6ca1ddaa-079b-4e74-80cc-72eef32e6599)

---

## Part 5: Advanced Variable Management
This is one of the most powerful modern additions to Excel, fundamentally changing how complex formulas are written.

### `LET()`
* **What it does:** Assigns names to calculation results. This allows storing intermediate calculations, values, or defining names inside a formula.
* **When to use it:** When a formula requires the same calculation multiple times (e.g., `IF(LongCalculation > 10, LongCalculation, 0)`). `LET` computes the value once, assigns it to a variable, and reuses it, vastly improving calculation speed and formula readability.
* **Official Docs:** [LET function](https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999)
