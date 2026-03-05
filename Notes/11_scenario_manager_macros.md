# Advanced Excel: Scenario Manager & Macros

Complex business models require:
1. Controlled scenario testing  
2. Automation of repetitive processes  

This document covers:
- Scenario Manager
- What-If Analysis
- Scenario Summary Reports
- Introduction to Macros
- Recording Macros
- Relative vs Absolute Macros
- Macro security
- VBA basics
- Best practices for automation

---

# PART 1: Scenario Manager

Scenario Manager is a What-If Analysis tool used to test different input combinations without changing the core model.

Location:
Data → What-If Analysis → Scenario Manager

---

# 1. Purpose of Scenario Manager

Used when:
- Multiple assumptions exist
- Financial projections depend on variables
- Need to compare best / worst / base case
- Budget forecasting

It changes selected input cells while keeping formulas constant.

---

# 2. Key Components

| Term | Meaning |
|------|----------|
| Changing Cells | Input cells that vary |
| Scenario | Saved set of input values |
| Scenario Summary | Report comparing scenarios |

---

# 3. Creating a Scenario

## Step 1: Define Input Cells

Example:
- Revenue growth rate
- Cost percentage
- Discount rate

These must already exist in the model.

---

## Step 2: Add Scenario

1. Go to Scenario Manager.
2. Click **Add**.
3. Name scenario (e.g., Base Case).
4. Select Changing Cells.
5. Enter values.
6. Click OK.

Repeat for:
- Best Case
- Worst Case
- Conservative Case

---

# 4. Switching Between Scenarios

In Scenario Manager:
- Select scenario.
- Click **Show**.

Excel replaces values in changing cells automatically.

---

# 5. Scenario Summary Report

Click:
Scenario Manager → Summary

Choose:
- Scenario Summary
- PivotTable Report

Select result cells (e.g., Profit, NPV, ROI).

Excel generates comparison sheet.

---

# 6. Limitations of Scenario Manager

- Manual updates required
- Not dynamic like Data Tables
- Not suitable for hundreds of variables
- Cannot automate complex dependency chains

---

# 7. Scenario Manager vs Data Table

| Feature | Scenario Manager | Data Table |
|----------|------------------|------------|
| Multiple input sets | Yes | Limited |
| Dynamic recalculation | Manual switch | Automatic |
| Large scale sensitivity | Limited | Better |
| Summary generation | Yes | No |

---

# PART 2: Macros

A Macro is a recorded sequence of actions automated using VBA (Visual Basic for Applications).

Macros remove repetitive manual work.

---

# 8. Enabling Developer Tab

File → Options → Customize Ribbon → Check Developer

---

# 9. Recording a Macro

Developer → Record Macro

Options:
- Macro Name
- Shortcut key
- Store macro in:
  - This Workbook
  - New Workbook
  - Personal Macro Workbook

Perform actions.
Click Stop Recording.

Excel converts actions into VBA code.

---

# 10. Running a Macro

Developer → Macros → Select → Run

Or use shortcut key assigned.

---

# 11. Relative vs Absolute Macros

When recording:

- Default = Absolute references
  - Repeats action on same exact cells

- Use Relative Reference Mode
  - Actions repeat relative to active cell

Enable:
Developer → Use Relative References

---

# 12. Viewing VBA Code

Press:

```excel
Alt + F11
```

Opens VBA Editor.

Macro example:

```vba
Sub FormatReport()
    Range("A1:D1").Font.Bold = True
End Sub
```

---

# 13. Basic VBA Structure

```vba
Sub MacroName()

    ' Code goes here

End Sub
```

Key elements:

- Sub = procedure start
- End Sub = procedure end
- Range() = reference
- Cells(row,column) = cell reference

---

# 14. Common Macro Uses

- Formatting reports
- Cleaning data
- Generating summary sheets
- Exporting PDF
- Batch processing
- Repeating filters
- Applying standard templates

---

# 15. Macro Security

Macros can contain harmful code.

Security Levels:
File → Options → Trust Center → Trust Center Settings → Macro Settings

Options:
- Disable all macros
- Disable with notification
- Enable all macros (not recommended)

---

# 16. Saving Macro-Enabled Files

Use file type:

```
.xlsm
```

Standard `.xlsx` cannot store macros.

---

# 17. Editing Macros

In VBA Editor:
- Modify code manually.
- Add loops.
- Add conditions.
- Add user input boxes.

Example:

```vba
Sub ClearColumn()
    Columns("A:A").ClearContents
End Sub
```

---

# 18. Loop Example

```vba
Sub LoopExample()

    Dim i As Integer

    For i = 1 To 10
        Cells(i,1).Value = i
    Next i

End Sub
```

---

# 19. When to Use Macros

Use when:
- Task is repetitive
- Rule-based formatting required
- Large dataset processing needed
- Standardized reporting required

Do not use macros:
- For simple formulas
- If Power Query can solve problem
- If automation unnecessary

---

# 20. Scenario Manager vs Macros

| Feature | Scenario Manager | Macro |
|----------|------------------|--------|
| Purpose | Assumption testing | Automation |
| Requires coding | No | Yes (or recording) |
| Dynamic switching | Manual | Can be automated |
| Reporting | Built-in summary | Customizable |

---

# 21. Best Practices

- Separate inputs from calculations.
- Name input cells clearly.
- Document each scenario.
- Store personal macros in Personal Workbook.
- Avoid unnecessary macro complexity.
- Test macros on sample data.
- Backup before running large automation.

---

# Quick Revision

- Scenario Manager tests assumption sets.
- Changing cells define scenario.
- Summary compares outcomes.
- Macros automate repetitive tasks.
- Use Relative Reference for flexible macros.
- Save macro files as .xlsm.
- VBA allows deeper automation.
- Use macros only when justified.

---