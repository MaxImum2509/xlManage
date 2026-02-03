---
name: excel-development-rules
description: VBA Excel development standards and patterns. Covers ListObject manipulation, UserForm MVVM architecture, event-driven patterns, RAII performance optimization (ExcelOptimizer), error handling framework, naming conventions, and file encoding (Windows-1252/CRLF). Use when writing VBA code, creating Excel macros, designing UserForms, working with ListObjects/Tables, or troubleshooting VBA errors.
disable-model-invocation: true
argument-hint: [topic or VBA question]
allowed-tools: Read, Grep, Glob, Write, Edit, Bash
---

# VBA Excel Development

Expert-level guidance for building robust, maintainable VBA Excel automation.

## Quick Reference

| Topic | See |
|-------|-----|
| Advanced patterns (MVVM, FindRowByField) | [references/ADVANCED_PATTERNS.md](references/ADVANCED_PATTERNS.md) |
| Troubleshooting | [references/TROUBLESHOOTING.md](references/TROUBLESHOOTING.md) |
| Examples | [examples/](examples/) |
| Templates | [assets/templates/](assets/templates/) |

## Architecture: Three Layers

1. **Data Access**: Excel objects (Workbooks, Worksheets, Ranges)
2. **Business Logic**: Domain rules and calculations
3. **Presentation**: UserForms and user interactions

## File Encoding

**CRITICAL**: All VBA files (.bas, .cls, .frm) must use **Windows-1252** encoding with CRLF line endings. **Never use emojis** in VBA code strings.

```python
with open("modUtils.bas", "w", encoding="windows-1252", newline="\r\n") as f:
    f.write(vba_code)
```

## VBA File Structure

Standard order in every VBA file:
1. File header (description, author, date)
2. `Option Explicit` (mandatory)
3. Public constants
4. Private constants / variables
5. Events (Initialize, Terminate)
6. Public procedures
7. Private procedures

## Naming Conventions

| Element | Pattern | Example |
|---------|---------|---------|
| Files | `mod<Name>.bas`, `cls<Name>.cls`, `frm<Name>.frm` | `modUtils.bas` |
| Procedures | PascalCase | `ProcessData` |
| Variables | camelCase | `dataRange` |
| Constants | SCREAMING_SNAKE | `MAX_ROWS` |
| Private members | `m` prefix | `mWorksheet` |

## Procedure Headers

```vba
'-------------------------------------------------------------------------------
' ProcedureName
' Description : Brief description
' Parameters  : paramName - Type - Description
' Return      : Type - Description
'-------------------------------------------------------------------------------
```

## Data Access Strategy

### ListObject first, Range for single cells

| Need | Use | Example |
|------|-----|---------|
| Structured data (tables) | `ListObject` | `ws.ListObjects("SalesData")` |
| Single cell read/write | `Range` | `ws.Range("A1").Value` |
| Named range lookup | `Range` | `ws.Range("ParamTaxRate").Value` |

**ListObject** is the preferred way to access structured data. It provides named columns, auto-expanding ranges, built-in filtering/sorting, and structured references. **Range** is appropriate for accessing a specific cell, a named range parameter, or non-tabular data.

### Array bulk read/write (critical performance pattern)

Read an entire column or range into a VBA array, process in memory, then write back. This is **100x+ faster** than cell-by-cell access.

```vba
' --- Read a full table column into an array ---
Dim colData As Variant
colData = tbl.ListColumns("EmployeeID").DataBodyRange.Value
' colData is now a 2D array (n, 1) - access via colData(i, 1)

' --- Read an entire table into an array ---
Dim tableData As Variant
tableData = tbl.DataBodyRange.Value

' --- Read a named range into an array ---
Dim rangeData As Variant
rangeData = ws.Range("MyNamedRange").Value

' --- Write an array back to a range (same dimensions) ---
tbl.DataBodyRange.Value = tableData
ws.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
```

**Key points**:
- `.Value` on a multi-cell range always returns a **2D Variant array** (base 1)
- Single column: `colData(i, 1)` — single row: `rowData(1, j)`
- Write back by assigning an array to a range of matching dimensions
- Use `Resize` when writing to a range that may differ in size

### ListObject operations

```vba
Dim tbl As ListObject
Set tbl = ws.ListObjects("SalesData")

' Add row
Dim newRow As ListRow
Set newRow = tbl.ListRows.Add
newRow.Range.Cells(1, 1).Value = "New Data"

' Find row (array-based, 100x faster)
Set foundRow = FindRowByField(tbl, "EmployeeID", 12345)
```

See [references/ADVANCED_PATTERNS.md](references/ADVANCED_PATTERNS.md) for FindRowByField and [examples/ListObject_Patterns.bas](examples/ListObject_Patterns.bas).

## Core Patterns

### Performance - RAII (ExcelOptimizer)

```vba
Public Sub OptimizedOperation()
    Dim optimizer As New ExcelOptimizer
    optimizer.Initialize
    ' Your code here - settings auto-restored on exit
End Sub
```

See [examples/ExcelOptimizer.cls](examples/ExcelOptimizer.cls).

### UserForm - MVVM Pattern

```vba
Private mViewModel As UserFormViewModel

Private Sub UserForm_Initialize()
    Set mViewModel = New UserFormViewModel
    mViewModel.Initialize Me
End Sub
```

See [references/ADVANCED_PATTERNS.md](references/ADVANCED_PATTERNS.md).

### Error Handling

```vba
Public Sub RobustOperation()
    On Error GoTo ErrorHandler
    ' Your code here
CleanUp:
    Exit Sub
ErrorHandler:
    LogError Err.Number, Err.Description, "RobustOperation"
    Resume CleanUp
End Sub
```

See [examples/Error_Handling.bas](examples/Error_Handling.bas).

## Best Practices

1. **Always** `Option Explicit` at top of every module
2. **Never** `Select`/`Activate` - reference objects directly
3. **Cache** worksheet references
4. **Use arrays** for bulk operations on large datasets
5. **Implement** proper error handling with cleanup
6. **Separate concerns** using class modules
7. **Test edge cases** (empty ranges, #N/A values)
