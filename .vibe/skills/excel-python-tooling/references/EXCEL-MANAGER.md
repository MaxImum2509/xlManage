# ExcelManager API

RAII context manager for all Excel operations via pywin32.

## Constructor

```python
ExcelManager(visible=False)
```

## Workbook Management

```python
with ExcelManager(visible=False) as mgr:
    wb = mgr.create_workbook("data/output.xlsx")
    wb = mgr.open_workbook("data/existing.xlsx")
    mgr.save()
    mgr.save_as("output/data.xlsm")  # Format auto-detected
```

## Sheet and ListObject

```python
with ExcelManager() as mgr:
    wb = mgr.create_workbook()

    # Create sheet + ListObject in one call
    sheet, lo = mgr.add_sheet_with_listobject("ADV", "tbADV", ["UserName", "LastName", "FirstName"])

    # Or separately
    sheet = mgr.add_sheet("Configuration")
    lo = mgr.add_listobject(sheet, "tbParameters", ["Parameter", "Value", "Description"])

    # Add data rows
    mgr.add_listobject_row(sheet, "tbADV", ["patrick", "Hostein", "Patrick"])

    # Delete default sheets
    mgr.delete_default_sheets(["ADV", "Configuration"])

    mgr.save_as("data/data.xlsx")
```

**CRITICAL**: Do NOT write headers manually before `add_listobject`.

## VBA Module Management

```python
with ExcelManager() as mgr:
    wb = mgr.open_workbook("tbAffaires.xlsm")
    vba_code = """
Option Explicit

Private Sub Class_Initialize()
    ' Code here
End Sub
"""
    mgr.add_class_module("clsApplicationState", vba_code)
    mgr.save()
```

## Save Formats

| Extension | Code | Constant |
|-----------|------|----------|
| `.xlsx` | 51 | xlOpenXMLWorkbook |
| `.xlsm` | 52 | xlOpenXMLWorkbookMacroEnabled |

## ExcelTestCase

Base class for Python tests requiring Excel via win32com. Uses session-scope fixture.

```
conftest.py
  excel_app (fixture session-scope) -> ONE Excel.Application instance

test_base.py
  ExcelTestCase
    _inject_excel_fixture (autouse) -> injects self.excel
    _open_workbook_safely(filepath)
    _safe_close_workbook()
```

### Anti-Patterns

```python
# WRONG: Creating Excel instance in setUpClass
cls.excel = win32.Dispatch("Excel.Application")

# WRONG: Calling Quit()
cls.excel.Quit()  # Causes RPC error 0x800706be

# WRONG: Not closing workbooks
wb = self.excel.Workbooks.Open(path)  # Leaked!

# CORRECT
wb = self._open_workbook_safely(path)
```

## COM Lifecycle

COM uses reference counting. Excel is an out-of-process COM server.

**Golden Rule**: NEVER call `excel.Quit()` explicitly. Let Python's GC release COM references naturally.

```python
@classmethod
def tearDownClass(cls):
    pass  # GC releases references -> Excel shuts down automatically
```
