# Advanced Excel: Automating Lookup Systems

Lookup functions are the backbone of reporting, dashboards, reconciliation models, and financial systems.  
This document explains lookup architecture from basic to advanced dynamic referencing.

---

# 1. Understanding Lookup Logic

A lookup formula answers:

> “Find this value in one place and return related information from another place.”

Core components of lookup:

1. Lookup Value
2. Lookup Range
3. Return Range
4. Match Type (Exact or Approximate)

---

# 2. VLOOKUP

## 2.1 Syntax

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

### Parameters

| Parameter | Meaning |
|-----------|---------|
| lookup_value | Value to search |
| table_array | Full table including lookup column |
| col_index_num | Column number from left (1-based) |
| range_lookup | TRUE (approximate) / FALSE (exact) |

---

## 2.2 Exact Match VLOOKUP

```excel
=VLOOKUP(A2, $A$2:$D$100, 3, FALSE)
```

### FALSE or 0
- Requires exact match
- Most commonly used
- Used for IDs, product codes, names

---

## 2.3 Approximate Match (Range Lookup)

```excel
=VLOOKUP(A2, $A$2:$D$100, 3, TRUE)
```

### TRUE or 1
- Requires sorted lookup column (ascending)
- Returns closest lower value
- Used for:
  - Tax brackets
  - Commission slabs
  - Grading systems

If omitted, default = TRUE.

⚠ Dangerous if data not sorted.

---

## 2.4 Limitations of VLOOKUP

- Can only look right
- Breaks if columns inserted
- Requires column index number
- Slower in large models

---

# 3. MATCH Function

Returns position number.

## Syntax

```excel
=MATCH(lookup_value, lookup_array, [match_type])
```

### match_type Options

| Value | Meaning |
|-------|---------|
| 0 | Exact match |
| 1 | Less than (default, sorted ascending required) |
| -1 | Greater than (sorted descending required) |

---

### Example

```excel
=MATCH(A2, A:A, 0)
```

Returns row position.

---

# 4. INDEX Function

Returns value at specified position.

## Syntax

```excel
=INDEX(array, row_num, [column_num])
```

### Parameters

| Parameter | Meaning |
|-----------|---------|
| array | Data range |
| row_num | Row number inside range |
| column_num | Optional column number |

---

### Example

```excel
=INDEX(A2:D100, 5, 2)
```

Returns value from 5th row, 2nd column inside range.

---

# 5. INDEX + MATCH (Professional Standard)

Replaces VLOOKUP.

## Basic Structure

```excel
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
```

### Example

```excel
=INDEX(C:C, MATCH(A2, A:A, 0))
```

---

## Why INDEX-MATCH is Better

- Lookup left or right
- Resistant to column insertion
- Cleaner logic
- More flexible
- Used in financial modelling

---

# 6. Range Lookup Logic (Advanced Approximate Matching)

Used for slabs / thresholds.

Example:

```excel
=VLOOKUP(A2, Table, 2, TRUE)
```

Lookup column must be sorted ascending.

Equivalent MATCH approach:

```excel
=INDEX(B:B, MATCH(A2, A:A, 1))
```

---

# 7. XLOOKUP (Modern Replacement)

## Syntax

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

### Parameters Explained

| Parameter | Meaning |
|-----------|---------|
| lookup_value | Value to find |
| lookup_array | Search column |
| return_array | Column to return |
| if_not_found | Optional error message |
| match_mode | 0 exact (default), -1 next smaller, 1 next larger, 2 wildcard |
| search_mode | 1 first to last (default), -1 reverse |

---

### Example

```excel
=XLOOKUP(A2, A:A, C:C, "Not Found")
```

Advantages:
- No column index
- Can search left
- Built-in error handling
- Handles approximate without sorting issues

---

# 8. ADDRESS Function

Creates cell reference as text.

## Syntax

```excel
=ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])
```

### Parameters

| Parameter | Meaning |
|-----------|---------|
| row_num | Row number |
| column_num | Column number |
| abs_num | 1=$A$1 (default), 2=A$1, 3=$A1, 4=A1 |
| a1 | TRUE (A1 style), FALSE (R1C1 style) |
| sheet_text | Optional sheet name |

---

### Example

```excel
=ADDRESS(5,3)
```

Returns:

```
$C$5
```

---

# 9. INDIRECT Function

Converts text into actual reference.

## Syntax

```excel
=INDIRECT(ref_text, [a1])
```

### Parameters

| Parameter | Meaning |
|-----------|---------|
| ref_text | Text reference |
| a1 | TRUE (default A1), FALSE (R1C1) |

---

### Example

```excel
=INDIRECT("A1")
```

Returns value in A1.

---

### Dynamic Sheet Example

```excel
=INDIRECT("Sheet1!A1")
```

Often combined with:

```excel
=INDIRECT(SheetName & "!A1")
```

---

# 10. OFFSET Function

Returns reference shifted from starting point.

## Syntax

```excel
=OFFSET(reference, rows, cols, [height], [width])
```

### Parameters

| Parameter | Meaning |
|-----------|---------|
| reference | Starting cell |
| rows | Move down (+) / up (-) |
| cols | Move right (+) / left (-) |
| height | Optional row count |
| width | Optional column count |

---

### Example

```excel
=OFFSET(A1,2,1)
```

Moves 2 rows down, 1 column right.

---

## OFFSET + MATCH Combination

Dynamic lookup example:

```excel
=OFFSET(A1, MATCH(A2,A:A,0), 2)
```

---

# 11. When to Use What

| Situation | Best Function |
|------------|--------------|
| Simple right lookup | VLOOKUP (exact) |
| Slab / threshold | VLOOKUP TRUE or MATCH 1 |
| Professional modelling | INDEX + MATCH |
| Dynamic sheet reference | INDIRECT |
| Build cell reference | ADDRESS |
| Shift reference | OFFSET |
| Modern Excel | XLOOKUP |

---

# 12. Common Parameter Confusion

| Parameter | Meaning |
|-----------|---------|
| 0 | Exact match |
| FALSE | Exact match |
| 1 | Approximate match |
| TRUE | Approximate match |
| -1 | Next larger match (MATCH only) |

---

# 13. Professional Advice

- Always specify match type explicitly.
- Avoid default TRUE in VLOOKUP.
- Prefer INDEX-MATCH in financial models.
- Avoid volatile functions (OFFSET, INDIRECT) in very large models.
- Document approximate lookups clearly.
- Use XLOOKUP in modern Excel.

---

# Final Summary

Lookup hierarchy:

Basic → VLOOKUP  
Flexible → INDEX + MATCH  
Modern → XLOOKUP  
Dynamic reference → INDIRECT / OFFSET  
Reference builder → ADDRESS  

Clean data first.  
Validate input.  
Then build lookup architecture.

---