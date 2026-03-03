## 1. Text Manipulation Functions
Text functions allow you to standardize data entries and extract specific strings from larger datasets.

### Case Standardization
* **UPPER:** Converts all text in a cell to capital letters.
    * *Syntax:* `=UPPER(text)`
* **LOWER:** Converts all text to lowercase letters.
    * *Syntax:* `=LOWER(text)`
* **PROPER:** Capitalizes the first letter of every word (ideal for Names or Addresses).
    * *Syntax:* `=PROPER(text)`

### Text Joining and Cleaning
* **CONCAT:** Joins multiple text strings or ranges together.
    * *Syntax:* `=CONCAT(text1, [text2], ...)`
    * *Note:* Unlike the older CONCATENATE, this can handle entire ranges of cells.
* **TRIM:** Removes all leading, trailing, and extra spaces between words.
    * *When to use:* Use this on data exported from external databases to fix "invisible" space errors that break VLOOKUPs.
    * *Syntax:* `=TRIM(text)`
* **LEN:** Returns the total number of characters in a string, including spaces.
    * *Syntax:* `=LEN(text)`

### Data Extraction
| Function | Purpose | Syntax |
| :--- | :--- | :--- |
| **LEFT** | Extracts characters from the start of a string. | `=LEFT(text, num_chars)` |
| **RIGHT** | Extracts characters from the end of a string. | `=RIGHT(text, num_chars)` |
| **MID** | Extracts characters from the middle, based on a starting point. | `=MID(text, start_num, num_chars)` |
| **FIND** | Locates the position of a specific character (Case-Sensitive). | `=FIND(find_text, within_text)` |



---

## 2. Date and Time Functions
Excel treats dates as sequential serial numbers, allowing for mathematical calculations on time periods.

### Current Time and Date
* **TODAY:** Returns the current date. It updates every time the sheet recalculates.
    * *Syntax:* `=TODAY()`
* **NOW:** Returns the current date and the specific time.
    * *Syntax:* `=NOW()`
    * *Note:* Both functions are "Volatile," meaning they can slow down very large workbooks.

### Duration and Age Calculations
* **YEARFRAC:** Calculates the fraction of a year between two dates.
    * *Use Case:* Perfect for calculating Age or Employee Tenure.
    * *Syntax:* `=YEARFRAC(start_date, end_date, [basis])`
    * *Parameters:* The `basis` (0-4) determines the day-count convention (e.g., US 30/360). Usually, 1 (Actual/Actual) is used for precision.
* **EDATE:** Returns a date that is a specific number of months before or after a start date.
    * *Use Case:* Calculating expiration dates or 90-day reviews.
    * *Syntax:* `=EDATE(start_date, months)`



---

## 3. Advanced Logical Combinations
Professional users often nest these functions together to solve complex data issues.

* **Dynamic Extraction:** Using `FIND` inside `LEFT` to extract everything before a space.
    * *Example:* `=LEFT(A2, FIND(" ", A2)-1)`
* **Age as Integer:** Using `INT` with `YEARFRAC` to get a clean age without decimals.
    * *Example:* `=INT(YEARFRAC(B2, TODAY()))`

---

## Professional Best Practices
* **Avoid Manual Entry:** Never type dates as "Jan 1st". Always use the format `MM/DD/YYYY` or `DD/MM/YYYY` according to your system settings so Excel recognizes it as a number.
* **Flash Fill (Ctrl + E):** For simple text patterns (like splitting First and Last names), try Flash Fill before writing a complex `MID` and `FIND` formula.
* **Date Formatting:** If a date formula returns a random number (e.g., 45123), change the Cell Format from **General** to **Short Date**.

---

## Key Takeaways
* **TRIM** is the most important cleaning function for imported data.
* **CONCAT** replaces manual "adding" of cells with `&`.
* **YEARFRAC** is the standard for financial and HR calculations involving time periods.
* **FIND** is case-sensitive; use **SEARCH** if you need a case-insensitive lookup.

---

## Common Errors
* **#VALUE!:** Usually occurs when a `FIND` function cannot locate the character you requested.
* **Incorrect MID Parameters:** If your `start_num` is greater than the total length of the text, Excel returns an empty string.
* **Date Logic:** If `start_date` is later than `end_date` in certain functions, it may return a negative value or an error.
* **Text vs. Number:** If a date is formatted as text, `YEARFRAC` will fail. Use `DATEVALUE` to convert it first.