# Advanced Excel: Formula Auditing, Protection & Arrays

Accurate models require two things:
1. Logical integrity (auditability)
2. Structural protection (controlled editing)

This document covers:
- Formula auditing tools
- Trace Precedents & Dependents
- Error checking
- Sheet & workbook protection
- Array formulas
- Dynamic arrays
- Array vs Tables
- Whether arrays work inside tables

---

# 1. Formula Auditing

Formula auditing ensures transparency in calculations.

Located in:
Formulas → Formula Auditing

---

# 1.1 Trace Precedents

Shows cells that feed into the selected formula.

### Purpose
- Understand formula logic
- Identify source inputs
- Detect incorrect references

### How to Use
1. Select formula cell.
2. Click **Trace Precedents**.

Blue arrows appear from input cells.

If reference is in another sheet:
- A dotted line with worksheet icon appears.
- Double-click dotted arrow to open reference list.

---

# 1.2 Trace Dependents

Shows which cells depend on the selected cell.

### Purpose
- Identify downstream impact
- Understand where a number flows
- Prevent accidental changes

### How to Use
1. Select input cell.
2. Click **Trace Dependents**.

Blue arrows point outward to dependent cells.

---

# 1.3 Remove Arrows

Clear all tracing lines.

Button:
Remove Arrows

---

# 1.4 Show Formulas

Shortcut:

```excel
Ctrl + `
```

Displays formulas instead of results.

Useful for:
- Reviewing large models
- Checking hard-coded numbers
- Auditing logic quickly

---

# 1.5 Evaluate Formula

Step-by-step breakdown of formula execution.

Select formula → Click **Evaluate Formula**

Helps understand:
- Nested IF logic
- MATCH positions
- Array calculations

---

# 1.6 Error Checking

Common Excel Errors:

| Error | Meaning |
|--------|----------|
| #N/A | Value not found |
| #REF! | Invalid reference |
| #VALUE! | Wrong data type |
| #DIV/0! | Division by zero |
| #NAME? | Function name incorrect |

Use:
Formulas → Error Checking

---

# 2. Protection in Excel

Protection prevents accidental modification.

---

# 2.1 Cell Locking

All cells are locked by default.

Lock only works after sheet protection is enabled.

To unlock:
1. Select cells.
2. Format Cells → Protection → Uncheck Locked.

---

# 2.2 Protect Sheet

Review → Protect Sheet

Options include:
- Select locked cells
- Select unlocked cells
- Format cells
- Insert rows
- Delete rows
- Sort
- Use AutoFilter

You can set password.

---

# 2.3 Protect Workbook

Protects:
- Sheet structure
- Prevents adding/deleting sheets

Review → Protect Workbook

---

# 2.4 Best Practice for Protection

- Unlock input cells only.
- Lock calculation cells.
- Protect sheet after testing.
- Avoid password if collaborative model needed.

---

# 3. Array Formulas

An array formula performs calculations on multiple values at once.

Older Excel required:
Ctrl + Shift + Enter

New Excel (Dynamic Arrays) does not.

---

# 3.1 Traditional Array Formula

Example:

```excel
=SUM(A1:A5*B1:B5)
```

Older versions:
Press Ctrl + Shift + Enter.

Excel adds `{}` around formula.

---

# 3.2 Array Functions

Common array-based functions:

- SUMPRODUCT
- TRANSPOSE
- FREQUENCY
- MMULT

Example:

```excel
=SUMPRODUCT(A1:A5, B1:B5)
```

No Ctrl+Shift+Enter needed.

---

# 3.3 Dynamic Arrays (Modern Excel)

Spill results automatically.

Example:

```excel
=FILTER(A2:B100, B2:B100>1000)
```

Spills results downward.

Common Dynamic Array Functions:

- FILTER
- SORT
- UNIQUE
- SEQUENCE
- RANDARRAY

---

# 4. Arrays vs Tables

Excel Tables and Arrays are different structures.

---

## 4.1 Excel Table

Created using:

```excel
Ctrl + T
```

Features:
- Structured references
- Automatic expansion
- Built-in filters
- Total row
- Column naming

Example structured reference:

```excel
=SUM(Table1[Sales])
```

---

## 4.2 Array Formula

- Logical calculation structure
- Operates on ranges
- Can return single or multiple results
- Not a structural container like table

---

# 5. Can Array Formulas Be Used Inside Tables?

Traditional array formulas:
❌ Not fully compatible inside Tables.

Dynamic arrays:
❌ Spill formulas cannot expand inside Tables.

Reason:
Tables control row expansion.
Dynamic arrays need spill space.

Workaround:
- Place array formula outside table.
- Reference table columns inside formula.

Example:

```excel
=FILTER(Table1, Table1[Sales]>1000)
```

Place outside the table.

---

# 6. Key Differences: Array vs Table

| Feature | Array | Table |
|----------|--------|--------|
| Structural container | No | Yes |
| Automatic expansion | Spill only | Yes |
| Structured references | No | Yes |
| Works inside itself | Yes | Yes |
| Spill allowed | Yes | No (inside table) |

---

# 7. Practical Auditing Workflow

1. Trace precedents.
2. Trace dependents.
3. Use Evaluate Formula.
4. Turn on Show Formulas.
5. Check for hard-coded numbers.
6. Confirm no circular references.

---

# 8. Circular References

Occurs when formula refers to itself.

Excel warning appears.

Check:
Formulas → Error Checking → Circular References

---

# 9. Best Practices for Professional Models

- Avoid hidden sheets without documentation.
- Use consistent naming.
- Lock calculation areas.
- Avoid excessive volatile functions.
- Audit before protecting.
- Keep input, calculation, output separated.

---

# Quick Revision

- Trace Precedents → See inputs.
- Trace Dependents → See outputs.
- Evaluate Formula → Step logic.
- Protect Sheet → Prevent edits.
- Arrays perform multi-value calculations.
- Tables structure data.
- Dynamic arrays spill.
- Spill formulas don’t work inside tables.

---