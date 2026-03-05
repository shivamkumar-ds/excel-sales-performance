#Navigation, Views & Efficient Data Entry


# 1. Navigating in Excel

## 1.1 Using Scroll Bars

- Horizontal scroll bar → move left/right
- Vertical scroll bar → move up/down
- Arrow buttons on scroll bars → move small increments

Useful but slow for large sheets.

---

## 1.2 Keyboard Navigation (Faster Method)

### Arrow Keys
Move one cell at a time.

### Page Up / Page Down
Moves one full screen at a time.

### Ctrl + Home
Instantly jumps to cell A1.
Extremely useful in large spreadsheets.

---

## 1.3 Moving Between Sheets

Click worksheet tabs at bottom.

When multiple workbooks are open:
- View → Switch Windows
- Shortcut: Ctrl + F6

---

# 2. Selecting Data Correctly

## Non-Contiguous Selection

To select separate areas:
1. Select first range
2. Hold Ctrl
3. Select second range

---

# 3. View Options & Screen Control

## 3.1 Split Screen

View → Split

Allows viewing multiple sections simultaneously.

Can split:
- Horizontally
- Vertically
- Both

Double-click split line to remove.

---

## 3.2 Freeze Panes (Very Important)

Used to keep headings visible while scrolling.

### Freeze Top Row
Freezes only first row.

### Custom Freeze

Click cell:
- Below rows to freeze
- Right of columns to freeze

Then:
View → Freeze Panes → Freeze Panes

Example:
If we want to freeze top 3 rows and first 3 columns:
Click cell D4, then freeze panes.

---

# 4. Creating and Editing Workbooks
## 4.1 Entering Dates

Use correct regional format:
- India → DD/MM/YYYY
- US → MM/DD/YYYY

Excel auto-adjusts column width for dates sometimes.

---

## 4.2 AutoSum

Home → AutoSum

Automatically selects numeric range above.

Used for quick totals.

---

# 5. Fill Handle & AutoFill Logic
## 5.1 Series Generation

Excel detects patterns:

- Numbers → increments automatically
- Dates → increases by 1 day
- Days of week → cycles automatically
- Months → cycles automatically

---
## 5.2 Creating Custom Patterns

To create pattern:
Select first two values
Then drag

Example:
8  
16  

Drag → 24, 32, 40…

Excel detects interval.

---
## 5.3 Flash Fill (Excel 2013+)

Automatically detects complex patterns.

Example:
If first email = abc.xyz@company.com  
Flash Fill creates remaining based on name pattern.

Access via:
Auto Fill Options → Flash Fill


## 6. Core Calculations
Excel formulas must begin with `=`. Use cell references (e.g., `B3`) instead of hardcoded numbers to ensure values update automatically.

### Basic Operators
* **Addition:** `+`
* **Subtraction:** `-`
* **Multiplication:** `*` (Asterisk)
* **Division:** `/` (Forward Slash)
* **Percentages:** `=Cell * 5%` (Excel treats 5% as 0.05).

---

## 7. Essential Functions
Functions handle complex tasks using the syntax: `=NAME(Range)`.

| Function | Purpose | Syntax Example |
| :--- | :--- | :--- |
| **SUM** | Adds a range of cells | `=SUM(B4:E4)` |
| **AVERAGE** | Finds the mean value | `=AVERAGE(B4:E4)` |
| **MAX** | Returns the highest value | `=MAX(B4:E4)` |
| **MIN** | Returns the lowest value | `=MIN(B4:E4)` |

We can use **AutoSum (Alt + =)** to quickly grab adjacent data. Excel will looks up first, then left.

---

## 8. Cell Referencing Styles
* **Relative Reference (Default):** Changes when copied (e.g., `A1` becomes `A2` when dragged down).
* **Absolute Reference ($):** Locks a specific cell so it doesn't move when copied. 
    * **Syntax:** `$J$1`
    * **Shortcut:** Press **F4** while selecting a cell to toggle the `$` signs.
    * **Usage:** Best for fixed tax rates, commissions, or constants.

---

# Excel Advanced: Comprehensive Error & Troubleshooting Guide

## 1. Common Excel Errors & Solutions
When a formula fails, Excel provides an error code. Understanding these is essential for debugging.

| Error Code | Cause | Solution |
| :--- | :--- | :--- |
| **#REF!** | A cell reference is no longer valid (usually because a row, column, or sheet was deleted). | Undo the deletion or re-link the formula to an existing cell. |
| **#VALUE!** | The formula includes the wrong type of data (e.g., trying to add a number to a word). | Check that all referenced cells contain numbers, not text. |
| **#DIV/0!** | You are trying to divide a number by zero or an empty cell. | Ensure the divisor cell is not blank or zero. |
| **#NAME?** | Excel doesn't recognize text in the formula (usually a typo in the function name). | Check for typos (e.g., `=SUMM()` instead of `=SUM()`). |
| **#NUM!** | The calculation results in a number too large or small for Excel to handle. | Check your math logic or constraints. |



---

## 2. Calculation & Logic Troubleshooting

### Circular References
* **The Issue:** A formula refers to its own cell (e.g., putting `=SUM(A1:A10)` inside cell `A10`).
* **The Result:** Excel cannot calculate because it creates an infinite loop. 
* **Fix:** Look at the Status Bar (bottom left) to see which cell is causing the "Circular Reference" and move the formula.

### Relative vs. Absolute Pitfalls
* **The Issue:** Dragging a formula results in `#REF!` or incorrect zeros.
* **The Cause:** Missing `$` signs. If you reference a single tax rate cell (like `J1`) and drag down, Excel tries to look at `J2`, `J3`, etc.
* **Fix:** Highlight the cell address in your formula and press **F4** to lock it (e.g., `$J$1`).

### Broken Cross-Sheet Links
* **The Issue:** Formulas showing old data or `#REF!`.
* **The Cause:** Renaming a worksheet tab or moving the file without updating links.
* **Fix:** Always press **Enter** while still on the source sheet when creating the link to ensure the syntax `'SheetName'!A1` is captured correctly.

--
* **Trace Precedents:** Use the "Formulas" tab -> "Trace Precedents" to see exactly which cells are feeding into your error.
* **Fill Without Formatting:** If copying a formula ruins your table borders, click the "Auto Fill Options" icon after dragging and select **Fill Without Formatting**.
* **Clean Data:** Use `Ctrl + Arrow Keys` to find gaps in your data. Functions like `SUM` or `AVERAGE` stop or skip at blank cells, which can skew results.

---
