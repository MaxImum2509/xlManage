---
name: python-com-automation
description: Generic Windows COM automation via pywin32. Covers COM fundamentals (in-process vs out-of-process servers), Dispatch/DispatchEx/EnsureDispatch, RAII context manager pattern, threading models (STA/MTA), COM lifecycle and reference counting, error handling (HRESULT, pywintypes.com_error), performance optimization, and testing patterns. Use when writing Python code that automates any Windows COM application (Excel, Word, Outlook, PowerPoint, AutoCAD, SAP GUI, etc.) or when debugging COM-related issues.
---

# Python COM Automation (pywin32)

Generic guide for controlling Windows applications via pywin32 COM automation.

## Quick Reference

| Topic | See |
|-------|-----|
| COM Lifecycle Deep Dive | [references/COM-LIFECYCLE.md](references/COM-LIFECYCLE.md) |
| Application Catalog | [references/APPLICATION-CATALOG.md](references/APPLICATION-CATALOG.md) |

## COM Fundamentals

COM (Component Object Model) allows Python to control Windows applications through well-defined interfaces.

### Server Types

| Type | Description | Examples |
|------|-------------|----------|
| **In-process** | DLL loaded into your Python process | ADO, Shell objects, WMI |
| **Out-of-process** | Separate .exe process | Excel, Word, Outlook, AutoCAD |

Out-of-process servers communicate via RPC (Remote Procedure Call), which adds latency but provides process isolation.

## pywin32 Essentials

### Dispatch Variants

```python
import win32com.client as win32

# Late-binding - connects to existing instance or creates one
app = win32.Dispatch("Excel.Application")

# Late-binding - always creates a new isolated instance (separate process)
# Use for parallel automation to avoid ROT (Running Object Table) conflicts
app = win32.DispatchEx("Excel.Application")

# Early-binding - generates type library wrappers for autocomplete & type safety
app = win32.gencache.EnsureDispatch("Excel.Application")
```

| Method | Binding | New Instance | Type Info | Use Case |
|--------|---------|-------------|-----------|----------|
| `Dispatch` | Late | No (reuses) | None | Simple scripts, single instance |
| `DispatchEx` | Late | Yes (isolated) | None | Parallel automation, isolation |
| `EnsureDispatch` | Early | No (reuses) | Full | IDE autocomplete, type safety |

### Generated Type Cache

`EnsureDispatch` creates Python wrappers in `win32com/gen_py/`. To clear stale cache:

```python
import win32com
import shutil
shutil.rmtree(win32com.__gen_path__, ignore_errors=True)
```

## RAII Context Manager Pattern

Generic template for any COM application:

```python
import win32com.client as win32
import gc


class COMAppManager:
    """RAII context manager for COM application lifecycle."""

    def __init__(self, prog_id: str, visible: bool = False):
        self._prog_id = prog_id
        self._visible = visible
        self._app = None

    def __enter__(self):
        self._app = win32.Dispatch(self._prog_id)
        self._app.Visible = self._visible
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._app is not None:
            try:
                # Close documents/workbooks - app-specific logic here
                pass
            finally:
                # Do NOT call app.Quit() for out-of-process servers
                del self._app
                self._app = None
                gc.collect()
        return False

    @property
    def app(self):
        return self._app
```

**Key rule**: For out-of-process COM servers, **never call `app.Quit()`** explicitly. It terminates the process while Python still holds COM references, causing RPC errors (0x800706be) when GC later calls `Release()` on the dead process.

## Threading Models

### STA vs MTA

| Model | Description | Use |
|-------|-------------|-----|
| **STA** (Single-Threaded Apartment) | One thread per apartment, COM marshals cross-thread calls | Default for most Office apps |
| **MTA** (Multi-Threaded Apartment) | Multiple threads share one apartment | Background services, WMI |

```python
import pythoncom

# Initialize for STA (default, required for Office apps)
pythoncom.CoInitialize()

# Initialize for MTA (background processing)
pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)

# Always uninitialize when done (per-thread)
pythoncom.CoUninitialize()
```

### Cross-Thread Marshaling

COM objects cannot be passed directly between threads. Use marshaling:

```python
import pythoncom

# In source thread: marshal the interface
stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
    pythoncom.IID_IDispatch, com_object
)

# In target thread: unmarshal
pythoncom.CoInitialize()
com_object = pythoncom.CoGetInterfaceAndReleaseStream(
    stream, pythoncom.IID_IDispatch
)
```

## COM Lifecycle

See [references/COM-LIFECYCLE.md](references/COM-LIFECYCLE.md) for deep dive.

### Reference Counting Rules

1. COM uses reference counting (`AddRef` / `Release`)
2. Python's `win32com` wraps COM refs - Python GC calls `Release` when wrapper is collected
3. For out-of-process servers: process stays alive while references exist
4. **Release inner objects before outer objects** (worksheets before workbook before app)

### Deterministic Cleanup

```python
# Release in correct order: inner -> outer
del document
del app
gc.collect()  # Force Release calls now, don't wait for GC cycle
```

## Error Handling

### COM Error Structure

```python
import pywintypes

try:
    app.SomeMethod()
except pywintypes.com_error as e:
    hr = e.hresult       # HRESULT code (e.g., -2147023174)
    msg = e.strerror      # Human-readable message
    exc = e.excepinfo     # (wcode, source, text, helpFile, helpContext, scode)
    argerr = e.argerr     # Argument error index
```

### Common HRESULT Codes

| HRESULT | Hex | Meaning |
|---------|-----|---------|
| -2147023174 | 0x800706BE | RPC server unavailable (app crashed/quit) |
| -2147417848 | 0x80010108 | RPC object disconnected |
| -2146827284 | 0x800A01A8 | Object required (bad reference) |
| -2147352567 | 0x80020009 | Exception from COM object |
| -2146959355 | 0x80080005 | Server execution failed |

### Retry Pattern for RPC Errors

```python
import time
import pywintypes

RPC_ERRORS = {-2147023174, -2147417848}  # RPC unavailable, object disconnected


def com_retry(func, *args, retries=3, delay=1.0):
    """Retry COM operations that fail due to transient RPC errors."""
    for attempt in range(retries):
        try:
            return func(*args)
        except pywintypes.com_error as e:
            if e.hresult not in RPC_ERRORS or attempt == retries - 1:
                raise
            time.sleep(delay * (attempt + 1))
```

## Performance Optimization

### Disable UI Updates (Generic Pattern)

```python
def optimize_com_app(app, optimize=True):
    """Disable UI updates for faster batch operations. App-specific properties."""
    for attr in ("ScreenUpdating", "EnableEvents"):
        if hasattr(app, attr):
            setattr(app, attr, not optimize)
    if hasattr(app, "Calculation"):
        app.Calculation = -4135 if optimize else -4105  # Manual / Automatic
```

### Batch Operations

Minimize cross-process COM calls - each call has RPC overhead for out-of-process servers:

```python
# SLOW: N individual cross-process calls
for i, value in enumerate(data):
    sheet.Cells(i + 1, 1).Value = value

# FAST: 1 cross-process call with batch data
sheet.Range(f"A1:A{len(data)}").Value2 = [(v,) for v in data]
```

### Early vs Late Binding Performance

Early binding (`EnsureDispatch`) is slightly faster because it resolves method DISPIDs at generation time rather than per-call via `GetIDsOfNames`. The difference is negligible for most scripts but measurable in tight loops.

## Testing COM Applications

### Session-Scope Fixture Pattern

```python
import pytest
import win32com.client as win32


@pytest.fixture(scope="session")
def com_app():
    """Share ONE COM app instance across all tests."""
    app = win32.Dispatch("SomeApp.Application")
    app.Visible = False
    yield app
    # Do NOT call app.Quit() - let COM GC handle shutdown


@pytest.fixture
def clean_document(com_app):
    """Provide a fresh document for each test."""
    doc = com_app.Documents.Add()
    yield doc
    doc.Close(SaveChanges=False)
```

### Test Isolation Tips

1. Use `session` scope for the COM app fixture (creating/destroying apps is expensive)
2. Use `function` scope for documents/workbooks (cheap to create/close)
3. Never call `Quit()` in teardown
4. Use `DisplayAlerts = False` to prevent modal dialogs blocking tests

## Common Applications

See [references/APPLICATION-CATALOG.md](references/APPLICATION-CATALOG.md) for a comprehensive catalog with ProgIDs, tips, and gotchas for each application.

## Anti-Patterns

```python
# NEVER - causes RPC errors for out-of-process servers
app.Quit()

# NEVER - leaks COM references, zombie processes
app = win32.Dispatch("Excel.Application")
# ... use app without context manager or cleanup ...

# NEVER - pass COM objects between threads without marshaling
threading.Thread(target=use_com_object, args=(excel,)).start()

# NEVER - ignore cleanup order
del app        # App first while inner refs exist
del workbook   # Inner refs now dangling!
```
