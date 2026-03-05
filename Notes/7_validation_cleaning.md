# Data Cleaning, Validation, Transformation & Lookup Architecture

# 1. Filling Blank Gaps in Exported Data

When reports are exported, category values (e.g., Country, Department) often appear once with blanks below.  
Blanks prevent sorting, grouping, and pivoting.

---

## 1.1 Manual Repair – Go To Special Method

Use for one-time fixes.

### Steps

1. In first blank cell:
```excel
=CellAbove
```

2. Copy the formula.
3. Select entire column.
4. Go to:
   Home → Find & Select → Go To Special → Blanks
5. Paste.
6. Copy entire column → Paste Values.

Now blanks are filled with static values.

---

## 1.2 Formula-Based Automated Method (Recommended)

Used when reports are refreshed regularly.

### Logic

If cell is blank → use value above  
Else → use original value

### Formula

```excel
=IF(ISBLANK(A2), A1, A2)
```

Place in separate “Clean” sheet to preserve raw data.

---

# 2. Standardizing Dates & Text

Excel often treats imported dates as text.

Text dates cannot:
- Be sorted properly
- Be grouped
- Be used in time calculations

---

## 2.1 Converting Text to Date

Use DATE function.

### Syntax

```excel
=DATE(year, month, day)
```

---

### Example

If A2 contains:

```
20231015
```

Convert using:

```excel
=DATE(LEFT(A2,4), MID(A2,5,2), RIGHT(A2,2))
```

Now Excel recognizes it as a true date.

---

## 2.2 Converting Date to Custom Text Format

Use TEXT function.

### Syntax

```excel
=TEXT(date, format_code)
```

### Examples

```excel
=TEXT(TODAY(), "MMMM")
```

Returns month name.

```excel
=TEXT(TODAY(), "DDDD")
```

Returns weekday name.

Common format codes:

- "DD"
- "MMM"
- "MMMM"
- "YYYY"

---

# 3. Removing Invisible Characters

Data exported from web, PDFs, ERP systems contains hidden characters.

These create duplicate values in PivotTables.

---

## 3.1 TRIM

Removes:
- Leading spaces
- Trailing spaces
- Extra internal spaces

```excel
=TRIM(A2)
```

---

## 3.2 CLEAN

Removes non-printing characters (ASCII 0–31)

```excel
=CLEAN(A2)
```

---

## 3.3 SUBSTITUTE

Replaces specific characters.

Used for non-breaking space (ASCII 160).

```excel
=SUBSTITUTE(A2, CHAR(160), "")
```

---

## 3.4 Combined Cleaning Formula

Professional cleaning pattern:

```excel
=VALUE(TRIM(SUBSTITUTE(A2, CHAR(160), "")))
```

Steps:
1. Remove non-breaking spaces
2. Trim extra spaces
3. Convert text to number

---

# 4. Diagnostic Tools for Dirty Data

When numbers don’t sum, they are likely stored as text.

Use diagnostic functions:

| Function | Purpose |
|----------|----------|
| ISNUMBER | Checks if value is numeric |
| LEN | Counts characters |
| CODE | Returns ASCII value |
| VALUE | Converts text to number |

---

## Example Diagnostics

```excel
=ISNUMBER(A2)
```

```excel
=LEN(A2)
```

```excel
=CODE(LEFT(A2,1))
```

---

# 5. Data Validation Architecture

Validation prevents bad data entry.

Clean input → clean reporting.

---

## 5.1 Static Drop-Down

Data → Data Validation → List

Source:

```excel
Yes,No,Pending
```

---

## 5.2 Dynamic Drop-Down 

1. Remove duplicates.
2. Convert to Table:

```excel
Ctrl + T
```

3. Name the column.
4. Use in validation:

```excel
=TableColumnName
```

Auto-expands.

---

## 5.3 Prevent Future Dates

```excel
=A2<=TODAY()
```

---

## 5.4 Prevent Duplicate Entries

```excel
=COUNTIFS($A:$A,A2)<=1
```

Ensures unique IDs.

---

## 5.5 Numeric Range Restriction

Example: Only allow positive values

```excel
=A2>0
```

---

# 6. Advanced Lookup Architecture

Lookups connect datasets.

---

# 6.1 VLOOKUP (Legacy)

## Syntax

```excel
=VLOOKUP(lookup_value, table_array, col_index, range_lookup)
```

Limitations:
- Can only look right
- Breaks when columns inserted
- Not dynamic

Example:

```excel
=VLOOKUP(A2, $A$2:$D$100, 3, FALSE)
```

---

# 6.2 INDEX + MATCH 

More flexible and stable.

## MATCH

```excel
=MATCH(lookup_value, lookup_range, 0)
```

## INDEX

```excel
=INDEX(return_range, row_number)
```

## Combined

```excel
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
```

Advantages:
- Lookup left or right
- Resistant to column insertions
- Modular structure

---

# 6.3 XLOOKUP 

Replaces VLOOKUP.

## Syntax

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
```

Example:

```excel
=XLOOKUP(A2, A:A, C:C, "Not Found")
```

Advantages:
- No column index number
- Can search left
- Built-in error handling
- Supports approximate match

---

# 7. Recommended Cleaning Workflow

1. Diagnose using ISNUMBER & LEN
2. Remove CHAR(160) using SUBSTITUTE
3. Apply TRIM
4. Convert using VALUE
5. Standardize dates
6. Apply validation
7. Build PivotTables

---

# 8. Best Practices

- Never overwrite raw data.
- Create separate Clean sheet.
- Convert lists to Tables.
- Always validate input cells.
- Use INDEX-MATCH instead of VLOOKUP in financial models.
- Always check data types before analysis.
- Document cleaning logic in model.

---

# Quick Revision

- TRIM removes spaces.
- CLEAN removes non-printing characters.
- SUBSTITUTE removes specific characters.
- VALUE converts text to numbers.
- COUNTIFS prevents duplicates.
- INDEX-MATCH is flexible lookup.
- XLOOKUP is modern replacement.
- Clean data first, analyse later.

---