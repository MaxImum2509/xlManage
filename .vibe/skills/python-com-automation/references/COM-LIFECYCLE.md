# COM Lifecycle

Deep dive on COM reference counting, garbage collection interactions, deterministic cleanup, and zombie process prevention.

## Reference Counting Fundamentals

COM uses `AddRef()` / `Release()` for lifecycle management. Every COM interface pointer must be paired:

```
CreateObject()  -> AddRef (refcount = 1)
QueryInterface() -> AddRef (refcount = 2)
Release()       -> refcount = 1
Release()       -> refcount = 0 -> object destroyed
```

In Python, `win32com` wraps COM pointers in Python objects. When the Python wrapper is garbage-collected, it calls `Release()` on the underlying COM pointer.

## Python GC and COM Interaction

### The Problem

Python uses reference counting + cyclic garbage collector. COM `Release()` is called when:
1. Python wrapper's refcount drops to 0 (immediate)
2. Cyclic GC collects the wrapper (non-deterministic)

For **in-process** COM servers (DLLs), this is usually fine - the DLL stays loaded.

For **out-of-process** COM servers (Excel, Word, etc.), the external process shuts down when all COM references are released. If `Release()` is called after the process is already gone (e.g., after `Quit()`), you get RPC errors.

### The Quit() Problem

```
Timeline of the bug:
1. excel = Dispatch("Excel.Application")   # refcount=1, Excel.exe running
2. wb = excel.Workbooks.Add()              # refcount on Workbooks, Workbook
3. excel.Quit()                            # Excel.exe terminates immediately
4. del wb                                  # GC calls Release() on dead process
   -> pywintypes.com_error: 0x800706BE    # RPC server unavailable!
```

**Solution**: Never call `Quit()`. Let all Python wrappers be collected first. When the last `Release()` fires, the out-of-process server shuts down gracefully on its own.

## Deterministic Cleanup Pattern

```python
import win32com.client as win32
import gc


class COMSession:
    """Ensures COM objects are released in correct order."""

    def __init__(self, prog_id: str):
        self._refs = []  # Track all COM objects in creation order
        self._app = win32.Dispatch(prog_id)

    def track(self, obj):
        """Register a COM object for ordered cleanup."""
        self._refs.append(obj)
        return obj

    def cleanup(self):
        """Release all tracked objects in reverse creation order."""
        for ref in reversed(self._refs):
            try:
                del ref
            except Exception:
                pass
        self._refs.clear()
        del self._app
        self._app = None
        gc.collect()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.cleanup()
        return False
```

## Zombie Process Detection and Prevention

### What Causes Zombie Processes

1. Python crashes without releasing COM references
2. Unhandled exceptions skip cleanup code
3. COM references leaked in global/module scope
4. Circular references prevent GC from collecting COM wrappers

### Detection

```python
import subprocess


def find_zombie_com_processes(exe_name: str) -> list[int]:
    """Find processes that may be COM zombies (hidden, no windows)."""
    result = subprocess.run(
        ["tasklist", "/fi", f"imagename eq {exe_name}", "/fo", "csv", "/nh"],
        capture_output=True, text=True,
    )
    pids = []
    for line in result.stdout.strip().splitlines():
        parts = line.strip('"').split('","')
        if len(parts) >= 2:
            pids.append(int(parts[1]))
    return pids
```

### Prevention Checklist

1. **Always use context managers** (RAII pattern) for COM app lifecycle
2. **Release inner objects first** - worksheets before workbooks before app
3. **Call `gc.collect()`** after deleting COM wrappers for deterministic cleanup
4. **Set `Visible = False`** and `DisplayAlerts = False` to prevent blocking dialogs
5. **Avoid global COM references** - scope them to functions or context managers
6. **Handle exceptions** - ensure cleanup runs even on error

### Last Resort: Process Kill

Only use when all else fails (e.g., during development/debugging):

```python
import subprocess


def kill_com_process(exe_name: str):
    """Force-kill a COM server process. LAST RESORT ONLY."""
    subprocess.run(
        ["taskkill", "/f", "/im", exe_name],
        capture_output=True,
    )
```

## Thread Safety Patterns

### Per-Thread COM Initialization

Each thread that uses COM must initialize its own apartment:

```python
import threading
import pythoncom
import win32com.client as win32


def worker():
    pythoncom.CoInitialize()  # STA for this thread
    try:
        app = win32.Dispatch("SomeApp.Application")
        # ... do work ...
        del app
    finally:
        pythoncom.CoUninitialize()


thread = threading.Thread(target=worker)
thread.start()
```

### Sharing COM Objects Between Threads

COM objects are apartment-bound. To share across threads, marshal the interface:

```python
import pythoncom


def producer(queue):
    """Thread that creates the COM object."""
    pythoncom.CoInitialize()
    app = win32.Dispatch("SomeApp.Application")
    # Marshal the interface into a stream
    stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, app._oleobj_
    )
    queue.put(stream)
    # Keep thread alive while consumer uses the object
    event.wait()
    pythoncom.CoUninitialize()


def consumer(queue):
    """Thread that uses the COM object."""
    pythoncom.CoInitialize()
    stream = queue.get()
    # Unmarshal to get a proxy in this thread's apartment
    dispatch = pythoncom.CoGetInterfaceAndReleaseStream(
        stream, pythoncom.IID_IDispatch
    )
    app = win32.Dispatch(dispatch)
    # Now safe to use app in this thread
    pythoncom.CoUninitialize()
```

## ReleaseDispatch and Explicit Cleanup

For cases where you need explicit control over when `Release()` is called:

```python
import gc

# Method 1: del + gc.collect()
del com_object
gc.collect()  # Forces Release() immediately

# Method 2: Assign None (equivalent to del for local scope)
com_object = None
gc.collect()

# Method 3: Using weakref for monitoring (diagnostic)
import weakref
ref = weakref.ref(com_wrapper, lambda r: print("COM ref released"))
```
