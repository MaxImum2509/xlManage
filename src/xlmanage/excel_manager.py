"""
Excel lifecycle management for xlmanage.

This file is part of xlManage.

xlManage is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

xlManage is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with xlManage.  If not, see <https://www.gnu.org/licenses/>.
"""

import gc
import re
from dataclasses import dataclass
from typing import Any

try:
    import pythoncom
    import pywintypes
    import win32com.client
    from win32com.client import CDispatch
except ImportError:
    # Allow the module to be imported even if pywin32 is not available
    # This is useful for testing and documentation purposes
    CDispatch = Any
    pythoncom = None
    pywintypes = None

# Import subprocess for process management
import subprocess

from .exceptions import ExcelConnectionError, ExcelInstanceNotFoundError, ExcelRPCError


@dataclass
class InstanceInfo:
    """Information about a running Excel instance.

    Attributes:
        pid: Process ID of the Excel process
        visible: Whether the instance is visible on screen
        workbooks_count: Number of open workbooks
        hwnd: Window handle for unique identification
    """

    pid: int
    visible: bool
    workbooks_count: int
    hwnd: int


class ExcelManager:
    """Manager for Excel application lifecycle.

    Implements RAII pattern via context manager.
    Never call app.Quit() - use the stop() protocol instead.
    """

    def __init__(self, visible: bool = False):
        """Initialize Excel manager.

        Args:
            visible: If True, the Excel instance will be visible on screen.
                     Default False (automated mode).
        """
        self._app: CDispatch | None = None
        self._visible: bool = visible

    def __enter__(self) -> "ExcelManager":
        """Enter context manager - start Excel instance."""
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Exit context manager - stop Excel instance."""
        self.stop()

    @property
    def app(self) -> CDispatch:
        """Return the COM Application object.

        Returns:
            The Excel Application COM object.

        Raises:
            ExcelConnectionError: If Excel is not started.
        """
        if self._app is None:
            raise ExcelConnectionError(
                0x80080005, "Excel application not started. Call start() first."
            )
        return self._app

    def start(self, new: bool = False) -> InstanceInfo:
        """Start or connect to an Excel instance.

        Args:
            new: If False, win32.Dispatch() reuses an instance via ROT.
                 If True, win32.DispatchEx() creates an isolated process.

        Returns:
            InstanceInfo with information about the connected instance.

        Raises:
            ExcelConnectionError: If Excel is not installed or COM is unavailable.
        """
        try:
            if new:
                # Create a new isolated Excel instance
                self._app = win32com.client.DispatchEx("Excel.Application")
            else:
                # Reuse existing instance via Running Object Table (ROT)
                self._app = win32com.client.Dispatch("Excel.Application")

            # Set visibility
            self._app.Visible = self._visible

            # Get instance information
            return self.get_instance_info(self._app)

        except Exception as e:
            # Handle COM errors
            if hasattr(e, "hresult"):
                raise ExcelConnectionError(
                    getattr(e, "hresult"), f"Failed to start Excel: {str(e)}"
                ) from e
            else:
                raise ExcelConnectionError(
                    0x80080005, f"Failed to start Excel: {str(e)}"
                ) from e

    def get_instance_info(self, app: CDispatch) -> InstanceInfo:
        """Get information about an Excel instance.

        Args:
            app: Excel Application COM object

        Returns:
            InstanceInfo with the instance details.
        """
        # Get basic information
        visible = app.Visible
        workbooks_count = app.Workbooks.Count

        # Get HWND (window handle) and PID
        # This requires ctypes for Windows API calls
        try:
            import ctypes
            import ctypes.wintypes

            # Get window handle
            hwnd = app.Hwnd

            # Get process ID from window handle
            pid = ctypes.wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

            return InstanceInfo(
                pid=pid.value,
                visible=visible,
                workbooks_count=workbooks_count,
                hwnd=hwnd,
            )
        except Exception:
            # Fallback if we can't get HWND/PID
            return InstanceInfo(
                pid=-1, visible=visible, workbooks_count=workbooks_count, hwnd=-1
            )

    def stop(self, save: bool = True) -> None:
        """Stop the managed Excel instance properly.

        Shutdown protocol:
        1. Disable alerts
        2. Close all workbooks
        3. Release COM references (del)
        4. Garbage collection
        5. Set _app to None

        IMPORTANT: NEVER call app.Quit() - causes RPC error 0x800706BE.

        Args:
            save: If True, save each workbook before closing

        Example:
            >>> mgr = ExcelManager()
            >>> mgr.start()
            >>> # ... work ...
            >>> mgr.stop(save=True)
        """
        if self._app is None:
            # Already stopped, nothing to do
            return

        try:
            # 1. Disable alerts (avoid confirmation dialogs)
            self._app.DisplayAlerts = False

            # 2. Close all workbooks
            workbooks = []
            try:
                # Copy list to avoid iteration issues
                for wb in self._app.Workbooks:
                    workbooks.append(wb)
            except (pywintypes.com_error, Exception):
                # Error during enumeration, ignore
                pass

            for wb in workbooks:
                try:
                    wb.Close(SaveChanges=save)
                    del wb
                except (pywintypes.com_error, Exception):
                    # Workbook already closed or inaccessible
                    continue

            # 3. Release main reference
            del self._app

        except (pywintypes.com_error, Exception):
            # RPC error (server disconnected), ignore
            # Instance is probably already dead
            pass

        finally:
            # 4. Garbage collection to release all COM references
            gc.collect()

            # 5. Mark as stopped
            self._app = None

    def stop_instance(self, pid: int, save: bool = True) -> None:
        """Stop an Excel instance identified by its PID.

        Connects to the instance via ROT or HWND, then applies
        the stop() protocol.

        Args:
            pid: Process ID of the target Excel instance
            save: If True, save before closing

        Raises:
            ExcelInstanceNotFoundError: If PID doesn't exist or is not Excel
            ExcelRPCError: If instance is disconnected

        Example:
            >>> mgr = ExcelManager()
            >>> mgr.stop_instance(12345, save=False)
        """
        # Enumerate all instances
        instances = enumerate_excel_instances()

        # Find instance with the right PID
        target_app = None
        for app, info in instances:
            if info.pid == pid:
                target_app = app
                break

        if target_app is None:
            # Fallback: check via tasklist if PID exists
            all_pids = enumerate_excel_pids()
            if pid not in all_pids:
                raise ExcelInstanceNotFoundError(
                    str(pid), "Process ID not found or not an Excel instance"
                )

            # PID exists but inaccessible via COM
            raise ExcelRPCError(
                0x800706BE, f"Excel instance PID {pid} is disconnected or inaccessible"
            )

        # Apply shutdown protocol
        try:
            target_app.DisplayAlerts = False

            # Close all workbooks
            workbooks = []
            for wb in target_app.Workbooks:
                workbooks.append(wb)

            for wb in workbooks:
                try:
                    wb.Close(SaveChanges=save)
                    del wb
                except (pywintypes.com_error, Exception):
                    continue

            # Release reference
            del target_app

        except pywintypes.com_error as e:
            # RPC error
            raise ExcelRPCError(e.hresult, f"RPC error during shutdown: {e}") from e

        finally:
            gc.collect()

    def stop_all(self, save: bool = True) -> list[int]:
        """Stop all active Excel instances.

        Enumerates via ROT and applies stop_instance() for each.

        Args:
            save: If True, save before closing

        Returns:
            list[int]: List of PIDs stopped successfully

        Example:
            >>> mgr = ExcelManager()
            >>> stopped = mgr.stop_all(save=True)
            >>> print(f"{len(stopped)} instances stopped")
        """
        # Enumerate all instances
        instances = enumerate_excel_instances()

        stopped_pids: list[int] = []

        for app, info in instances:
            try:
                # Apply shutdown protocol
                app.DisplayAlerts = False

                workbooks = []
                for wb in app.Workbooks:
                    workbooks.append(wb)

                for wb in workbooks:
                    try:
                        wb.Close(SaveChanges=save)
                        del wb
                    except (pywintypes.com_error, Exception):
                        continue

                del app

                stopped_pids.append(info.pid)

            except (pywintypes.com_error, Exception):
                # Instance disconnected, ignore
                continue

        # Final garbage collection
        gc.collect()

        return stopped_pids

    def get_running_instance(self) -> InstanceInfo | None:
        """Get the active Excel instance.

        Returns:
            InstanceInfo if an Excel instance is running, None otherwise.

        Raises:
            ExcelConnectionError: If COM connection fails.
        """
        try:
            app = win32com.client.Dispatch("Excel.Application")
            return self.get_instance_info(app)
        except Exception as e:
            if hasattr(e, "hresult"):
                raise ExcelConnectionError(
                    getattr(e, "hresult"), f"Failed to get running instance: {str(e)}"
                ) from e
            else:
                raise ExcelConnectionError(
                    0x80080005, f"Failed to get running instance: {str(e)}"
                ) from e

    def list_running_instances(self) -> list[InstanceInfo]:
        """Enumerate all running Excel instances.

        Uses ROT as priority, then fallback to tasklist if ROT fails.

        Returns:
            list[InstanceInfo]: List of instances with their information

        Example:
            >>> mgr = ExcelManager()
            >>> instances = mgr.list_running_instances()
            >>> for inst in instances:
            ...     print(f"PID {inst.pid}: {inst.workbooks_count} classeurs")

        Note:
            Uses multiple methods to find instances:
            1. Running Object Table (ROT) enumeration
            2. Fallback to tasklist PID enumeration
        """
        # Try via ROT
        rot_instances = enumerate_excel_instances()

        if rot_instances:
            # Extract just the InstanceInfo
            return [info for app, info in rot_instances]

        # Fallback: tasklist to get PIDs
        try:
            pids = enumerate_excel_pids()

            # Convert PIDs to InstanceInfo (limited info)
            instances = []
            for pid in pids:
                # Cannot get visible/workbooks_count without COM
                info = InstanceInfo(
                    pid=pid,
                    visible=False,  # Unknown
                    workbooks_count=0,  # Unknown
                    hwnd=0,  # Unknown
                )
                instances.append(info)

            return instances

        except RuntimeError:
            # Fallback also failed, return empty list
            return []


def enumerate_excel_instances() -> list[tuple[CDispatch, InstanceInfo]]:
    """Enumerate Excel instances via Running Object Table (ROT).

    The Windows ROT contains all active COM objects. We filter for
    Excel.Application instances.

    Returns:
        list[tuple[CDispatch, InstanceInfo]]: List of (app, info) for each instance

    Note:
        This function may fail if ROT access is blocked. Use
        enumerate_excel_pids() as fallback.
    """
    instances: list[tuple[CDispatch, InstanceInfo]] = []

    try:
        # Get Running Object Table
        rot = pythoncom.GetRunningObjectTable()
        # Enumerate monikers (COM object identifiers)
        monikers = rot.EnumRunning()

        for moniker in monikers:
            try:
                # Get moniker name
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)

                # Filter for Excel.Application
                # Name contains "!Microsoft_Excel_Application" or similar
                if "Excel.Application" not in name:
                    continue

                # Get COM object from ROT
                obj = rot.GetObject(moniker)

                # Cast to CDispatch
                app = win32com.client.Dispatch(
                    obj.QueryInterface(pythoncom.IID_IDispatch)
                )

                # Extract instance info
                info = _get_instance_info_from_app(app)

                instances.append((app, info))

            except (pywintypes.com_error, Exception):
                # Instance inaccessible or disconnected, skip
                continue

    except (pywintypes.com_error, Exception):
        # ROT inaccessible, return empty list (fallback needed)
        return []

    return instances


def _get_instance_info_from_app(app: CDispatch) -> InstanceInfo:
    """Extract InstanceInfo from an Application object.

    Args:
        app: Excel Application COM object

    Returns:
        InstanceInfo: Instance information
    """
    import ctypes

    # Get HWND (window handle)
    hwnd = app.Hwnd

    # Extract PID from HWND via Windows API
    process_id = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
    pid = process_id.value

    # Get other info
    visible = app.Visible
    workbooks_count = app.Workbooks.Count

    return InstanceInfo(
        pid=pid, visible=visible, workbooks_count=workbooks_count, hwnd=hwnd
    )


def enumerate_excel_pids() -> list[int]:
    """Fallback: Enumerate Excel PIDs via tasklist command.

    Used when ROT is not accessible. Returns only PIDs, not COM objects.

    Returns:
        list[int]: List of EXCEL.EXE process IDs

    Raises:
        RuntimeError: If tasklist fails (command not found)

    Note:
        This is a fallback method when ROT enumeration fails.
        Uses Windows tasklist command to find EXCEL.EXE processes.
    """
    try:
        # Call tasklist with filter for EXCEL.EXE
        result = subprocess.run(
            ["tasklist", "/fi", "imagename eq EXCEL.EXE", "/fo", "csv", "/nh"],
            capture_output=True,
            text=True,
            check=True,
            timeout=10,
        )

        pids: list[int] = []

        # Parse CSV output
        # Format: "EXCEL.EXE","12345","Console","1","123,456 K"
        for line in result.stdout.strip().split("\n"):
            if not line or "INFO:" in line:
                continue

            # Extract PID (2nd column)
            match = re.search(r'"EXCEL\.EXE","(\d+)"', line)
            if match:
                pid = int(match.group(1))
                pids.append(pid)

        return pids

    except subprocess.TimeoutExpired:
        raise RuntimeError("Timeout lors de l'énumération des processus Excel")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Échec de tasklist: {e}")
    except FileNotFoundError:
        raise RuntimeError("Commande tasklist introuvable (Windows requis)")


def connect_by_hwnd(hwnd: int) -> CDispatch | None:
    """Connect to Excel instance by window handle.

    Used when instance is not in ROT but is still active.

    Args:
        hwnd: Window handle (HWND)

    Returns:
        CDispatch | None: Excel Application COM object, or None if failed

    Note:
        This method attempts to connect to an Excel instance
        by its window handle using Windows Accessibility API.
    """
    import ctypes
    from ctypes import c_void_p
    from ctypes.wintypes import DWORD

    try:
        # Load oleacc.dll (Accessibility API)
        oleacc = ctypes.windll.oleacc

        # Constants
        objid_nativeom = -16  # ID for native Office object

        # Get IDispatch from HWND
        ptr = c_void_p()
        result = oleacc.AccessibleObjectFromWindow(
            hwnd,
            DWORD(objid_nativeom),
            ctypes.byref(pythoncom.IID_IDispatch),
            ctypes.byref(ptr),
        )

        if result != 0 or not ptr:
            return None

        # Convert IDispatch to CDispatch
        dispatch = pythoncom.ObjectFromLresult(ptr.value, pythoncom.IID_IDispatch, 0)
        app = win32com.client.Dispatch(dispatch)

        return app

    except Exception:
        # Connection failed, return None
        return None
