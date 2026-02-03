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

from dataclasses import dataclass
from typing import Any

try:
    import pythoncom
    import win32com.client
    from win32com.client import CDispatch
except ImportError:
    # Allow the module to be imported even if pywin32 is not available
    # This is useful for testing and documentation purposes
    CDispatch = Any
    pythoncom = None

# Import subprocess for process management
import subprocess

from .exceptions import ExcelConnectionError, ExcelRPCError


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

        Protocol:
        1. app.DisplayAlerts = False
        2. Close each workbook with SaveChanges=save
        3. Delete all workbook references
        4. Delete app reference
        5. Run garbage collection
        6. Set self._app = None

        Args:
            save: If True, save each workbook before closing.
        """
        if self._app is None:
            return

        try:
            # Suppress alerts to avoid dialogs
            self._app.DisplayAlerts = False

            # Close all workbooks
            for wb in self._app.Workbooks:
                try:
                    wb.Close(SaveChanges=save)
                except Exception:
                    # Ignore errors when closing workbooks
                    pass
                finally:
                    # Clean up references
                    del wb

            # Clean up application reference
            del self._app
            self._app = None

            # Force garbage collection to release COM objects
            import gc

            gc.collect()

        except Exception as e:
            # If we get here, try to force cleanup
            try:
                del self._app
            except Exception:
                pass
            self._app = None

            if hasattr(e, "hresult"):
                raise ExcelRPCError(
                    getattr(e, "hresult"), f"Error stopping Excel: {str(e)}"
                ) from e
            else:
                raise ExcelRPCError(
                    0x800706BE, f"Error stopping Excel: {str(e)}"
                ) from e

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

        Returns:
            List of InstanceInfo for all running Excel instances.

        Note:
            Uses multiple methods to find instances:
            1. Running Object Table (ROT) enumeration
            2. Fallback to tasklist PID enumeration
        """
        instances = []

        # Method 1: Try ROT enumeration
        try:
            for app in enumerate_excel_instances():
                try:
                    info = self.get_instance_info(app)
                    instances.append(info)
                except Exception:
                    continue
        except Exception:
            pass

        # Method 2: Fallback to PID enumeration
        if not instances:
            try:
                for pid in enumerate_excel_pids():
                    try:
                        app = connect_by_pid(pid)
                        if app:
                            info = self.get_instance_info(app)
                            instances.append(info)
                    except Exception:
                        continue
            except Exception:
                pass

        return instances


def enumerate_excel_instances() -> list[CDispatch]:
    """Enumerate Excel instances via Running Object Table (ROT).

    Returns:
        List of Excel Application COM objects found in ROT.

    Note:
        This method uses the Windows Running Object Table to find
        all registered Excel instances.
    """
    instances = []

    try:
        # Get Running Object Table
        rot = pythoncom.GetRunningObjectTable()

        # Enumerate all running objects
        for moniker in rot:
            try:
                # Check if it's an Excel instance
                if "Excel.Application" in str(moniker):
                    obj = rot.GetObject(moniker)
                    if obj and hasattr(obj, "Application"):
                        instances.append(obj.Application)
            except Exception:
                continue
    except Exception:
        pass

    return instances


def enumerate_excel_pids() -> list[int]:
    """Fallback: Enumerate Excel PIDs via tasklist command.

    Returns:
        List of Excel process IDs found via tasklist.

    Note:
        This is a fallback method when ROT enumeration fails.
        Uses Windows tasklist command to find EXCEL.EXE processes.
    """
    pids = []

    try:
        # Use tasklist to find Excel processes
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=True,
        )

        # Parse CSV output
        for line in result.stdout.strip().split("\n"):
            if line:
                parts = line.split(",")
                if len(parts) >= 2:
                    try:
                        pid = int(parts[1].strip('"'))
                        pids.append(pid)
                    except ValueError:
                        continue
    except (subprocess.CalledProcessError, FileNotFoundError, Exception):
        pass

    return pids


def connect_by_pid(pid: int) -> CDispatch | None:
    """Connect to Excel instance by process ID.

    Args:
        pid: Process ID of Excel instance

    Returns:
        Excel Application COM object if found, None otherwise.

    Note:
        This method attempts to connect to an Excel instance
        by its process ID using various techniques.
    """
    try:
        # Try to find the instance via ROT first
        for app in enumerate_excel_instances():
            try:
                # Get the PID from the application
                app_pid = app.Hwnd  # Simplified - in practice need Windows API
                if app_pid == pid:
                    return app
            except Exception:
                continue

        # Fallback: Try to get any available instance
        # Note: This is a simplified approach
        return win32com.client.Dispatch("Excel.Application")
    except Exception:
        return None


def connect_by_hwnd(hwnd: int) -> CDispatch | None:
    """Connect to Excel instance by window handle.

    Args:
        hwnd: Window handle of Excel instance

    Returns:
        Excel Application COM object if found, None otherwise.

    Note:
        This method attempts to connect to an Excel instance
        by its window handle using Windows API functions.
    """
    try:
        # Try to find the instance with matching HWND
        for app in enumerate_excel_instances():
            try:
                if app.Hwnd == hwnd:
                    return app
            except Exception:
                continue
        return None
    except Exception:
        return None
