---
name: excel-python-tooling
description: Excel automation via pywin32 COM. Covers ExcelManager RAII context manager, ExcelTestCase base class, VBAImporter for module import/export, workbook/worksheet/ListObject operations, VBA module handling (standard, class, UserForm), Windows-1252 encoding, and expert-level COM tips for Excel (DispatchEx, EnsureDispatch, zombie process prevention, batch operations). Use when writing Python code that interacts with Excel via COM, creating pywin32 automation, importing/exporting VBA modules, or testing Excel functionality.
---

# Excel Python Tooling (pywin32)

Guide for controlling Excel via pywin32 COM automation in the xlManage project.

## Quick Reference

| Topic | See |
|-------|-----|
| ExcelManager API | [references/EXCEL-MANAGER.md](references/EXCEL-MANAGER.md) |
| VBAImporter | [references/VBA-IMPORTER.md](references/VBA-IMPORTER.md) |

## Critical Rules

1. **NEVER call `excel.Quit()`** - causes RPC error (0x800706be). Let COM GC handle cleanup.
2. **NEVER use openpyxl** - project uses pywin32 exclusively.
3. **Always use `with` statement** for ExcelManager/VBAImporter.
4. **Windows-1252 encoding** for all VBA files (.bas, .cls, .frm) with CRLF line endings.

### Why No Quit()

1. `Quit()` terminates the Excel process immediately
2. Python still holds COM references
3. GC later calls `Release()` on dead process -> RPC fatal exception

## ExcelManager (RAII Context Manager)

**Always use this instead of raw win32com.**

```python
from excel_manager import ExcelManager

with ExcelManager(visible=False) as mgr:
    wb = mgr.create_workbook("data/output.xlsx")
    sheet, lo = mgr.add_sheet_with_listobject("ADV", "tbADV", ["UserName", "LastName"])
    mgr.add_listobject_row(sheet, "tbADV", ["patrick", "Hostein", "Patrick"])
    mgr.save_as("data/output.xlsx")
# Automatic cleanup guaranteed
```

**Do NOT write headers before `add_listobject`** - the method creates headers automatically.

See [references/EXCEL-MANAGER.md](references/EXCEL-MANAGER.md) for complete API.

## VBAImporter (Module Import/Export)

```python
from VBAImporter import VBAImporter

with VBAImporter(r"C:\path\to\workbook.xlsm") as importer:
    importer.import_module(r"C:\vba\modUtils.bas")
    importer.import_module(r"C:\vba\clsOptimizer.cls")
    importer.import_directory(r"C:\vba\modules", pattern="*.bas *.cls *.frm",
                              overwrite=True, auto_dependencies=True)
```

### Module Type Handling

| Type | Extension | Import Method |
|------|-----------|--------------|
| Standard | `.bas` | Direct `VBComponents.Import` |
| Class | `.cls` | Parse header (VB_Name, VB_PredeclaredId), create component, inject code |
| UserForm | `.frm`+`.frx` | `VBComponents.Import` (`.frx` must be in same dir) |

**Class module import - critical steps:**
1. Read with `encoding='windows-1252'`
2. Extract `VB_Name` and `VB_PredeclaredId` from `Attribute` lines
3. Strip header (everything before `Option Explicit`)
4. Create: `VBComponents.Add(2)`
5. Set name and `Properties("PredeclaredId")` **before** adding code
6. Clear auto-generated content, then `AddFromString(clean_code)`

See [references/VBA-IMPORTER.md](references/VBA-IMPORTER.md) for complete documentation.

## ExcelTestCase (Testing)

Session-scope fixture sharing ONE Excel instance across all tests:

```python
# conftest.py
@pytest.fixture(scope="session")
def excel_app():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    yield excel
    # No Quit() - COM lifecycle handles cleanup

# test_*.py
class TestMyFeature(ExcelTestCase):
    def test_module_exists(self):
        workbook = self._open_workbook_safely(XLSM_FILE)
        component = workbook.VBProject.VBComponents("modUtils")
        assert component is not None
```

## Performance Optimization

```python
def optimize_performance(excel, optimize=True):
    if optimize:
        excel.ScreenUpdating = False
        excel.Calculation = -4135  # xlCalculationManual
        excel.EnableEvents = False
    else:
        excel.ScreenUpdating = True
        excel.Calculation = -4105  # xlCalculationAutomatic
        excel.EnableEvents = True
```

## Expert-Level Tips

### DispatchEx vs Dispatch

```python
# Dispatch - connects to existing instance or creates one (shared ROT)
excel = win32.Dispatch("Excel.Application")

# DispatchEx - always creates a new isolated instance (separate process)
# Use for parallel operations to avoid cross-instance interference
excel = win32.DispatchEx("Excel.Application")
```

Use `DispatchEx` when running multiple Excel automations in parallel to get isolated instances that don't share the Running Object Table (ROT).

### EnsureDispatch for Early Binding

```python
from win32com.client import gencache

# Generates type library cache -> enables autocomplete and type safety
excel = gencache.EnsureDispatch("Excel.Application")
```

Early-binding via `EnsureDispatch` generates Python wrappers from the Excel type library, providing IDE autocomplete, type safety, and slightly better performance than late-binding `Dispatch`.

### Thread Safety

```python
import pythoncom

# Initialize COM for multithreaded use (call once per thread)
pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
```

Excel is an STA (Single-Threaded Apartment) server. If you must access it from multiple threads, either marshal the interface or use `CoInitializeEx` per thread with proper synchronization.

### COM Object Release Order

Always release inner objects before outer objects:

```python
# CORRECT order
del worksheet   # inner first
del workbook    # then container
del excel       # app last
import gc; gc.collect()

# WRONG - releasing app while inner refs exist causes errors
del excel
del workbook  # dangling reference!
```

### Avoiding Zombie Excel Processes

1. Use RAII context managers (ExcelManager) for automatic cleanup
2. Release COM references in correct order (inner to outer)
3. Call `gc.collect()` after deleting references for deterministic cleanup
4. As **last resort** only: `subprocess.run(["taskkill", "/f", "/im", "EXCEL.EXE"])`

```python
# Deterministic cleanup pattern
import win32com.client as win32
import gc

excel = win32.Dispatch("Excel.Application")
try:
    # ... work ...
finally:
    excel.DisplayAlerts = False
    for wb in excel.Workbooks:
        wb.Close(SaveChanges=False)
    # Do NOT call excel.Quit()
    del excel
    gc.collect()
```

### UDF Caching with Application.Volatile

```python
# In VBA UDFs called from Python-generated formulas:
# Mark a UDF as volatile to recalculate on every sheet change
# Application.Volatile True   (recalculates every time)
# Application.Volatile False  (default - only recalculates when inputs change)
```

Use `Application.Volatile` sparingly - volatile UDFs recalculate on every change, impacting performance.

### Batch Operations with Range.Value2

```python
# SLOW - cell by cell
for i, val in enumerate(data):
    sheet.Cells(i + 1, 1).Value = val

# FAST - batch via Value2 (avoids Date/Currency coercion overhead)
sheet.Range(f"A1:A{len(data)}").Value2 = [(v,) for v in data]
```

`Value2` is faster than `Value` for numeric data because it skips Date and Currency type coercion. Use it for bulk reads/writes when you don't need date formatting.

## Encoding for VBA Files

```python
# Write
with open(path, "w", encoding="windows-1252", newline="\r\n") as f:
    f.write(vba_code)

# Read
with open(path, "r", encoding="windows-1252") as f:
    content = f.read()
```

## Anti-Patterns

```python
# NEVER
excel.Quit()                    # Use RAII context manager
import openpyxl                 # Use pywin32 only
mgr = ExcelManager(); ...      # Always use `with` statement
sheet.Cells(1,1).Value = "Col"  # Don't write headers before add_listobject
```
