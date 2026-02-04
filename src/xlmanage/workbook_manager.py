"""
Workbook lifecycle management for xlmanage.

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
from pathlib import Path
from typing import Any

try:
    from win32com.client import CDispatch
except ImportError:
    CDispatch = Any

from .excel_manager import ExcelManager
from .exceptions import (
    ExcelConnectionError,
    WorkbookAlreadyOpenError,
    WorkbookNotFoundError,
    WorkbookSaveError,
)

# Excel file format constants
# See: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat
FILE_FORMAT_MAP: dict[str, int] = {
    ".xlsx": 51,  # xlOpenXMLWorkbook
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xls": 56,  # xlExcel8 (Excel 97-2003 format)
    ".xlsb": 50,  # xlExcel12 (Excel binary workbook)
    ".xltx": 54,  # xlOpenXMLTemplate
}


@dataclass
class WorkbookInfo:
    """Information about an Excel workbook.

    Attributes:
        name: Workbook filename (e.g., "data.xlsx")
        full_path: Full path to the workbook file
        read_only: Whether the workbook is opened in read-only mode
        saved: Whether all changes have been saved
        sheets_count: Number of worksheets in the workbook
    """

    name: str
    full_path: Path
    read_only: bool
    saved: bool
    sheets_count: int


def _detect_file_format(path: Path) -> int:
    """Detect Excel file format from file extension.

    Args:
        path: Path to the Excel file

    Returns:
        Excel FileFormat code (51, 52, 56, or 50)

    Raises:
        ValueError: If the file extension is not recognized

    Examples:
        >>> _detect_file_format(Path("data.xlsx"))
        51
        >>> _detect_file_format(Path("macro.xlsm"))
        52
    """
    extension = path.suffix.lower()

    if extension not in FILE_FORMAT_MAP:
        supported = ", ".join(FILE_FORMAT_MAP.keys())
        raise ValueError(
            f"Unsupported file extension '{extension}'. Supported formats: {supported}"
        )

    return FILE_FORMAT_MAP[extension]


def _find_open_workbook(app: CDispatch, path: Path) -> CDispatch | None:
    """Find an open workbook by path.

    Searches for a workbook in the Excel instance by comparing paths.
    First tries to match by FullName (complete path), then falls back
    to matching by Name (filename only).

    Args:
        app: Excel Application COM object
        path: Path to the workbook to find

    Returns:
        Workbook COM object if found, None otherwise

    Note:
        The search is case-insensitive on Windows.
        Paths are resolved to absolute paths before comparison.

    Examples:
        >>> app = win32com.client.Dispatch("Excel.Application")
        >>> wb = _find_open_workbook(app, Path("C:/data/test.xlsx"))
        >>> if wb:
        ...     print(f"Found: {wb.Name}")
    """
    # Resolve to absolute path for comparison
    resolved_path = path.resolve()
    filename = path.name

    # Iterate through all open workbooks
    for wb in app.Workbooks:
        try:
            # Method 1: Compare by full path (most reliable)
            wb_full_path = Path(wb.FullName).resolve()
            if wb_full_path == resolved_path:
                return wb

            # Method 2: Compare by filename only (fallback)
            # This handles cases where the path might be different
            # but it's actually the same file (network paths, etc.)
            if wb.Name.lower() == filename.lower():
                return wb

        except Exception:
            # If we can't read wb.FullName or wb.Name, skip this workbook
            continue

    return None


class WorkbookManager:
    """Manager for Excel workbook CRUD operations.

    This class provides methods to open, create, close, save, and list
    Excel workbooks. It depends on ExcelManager for COM access.

    Note:
        The ExcelManager instance must be started before using this manager.
    """

    def __init__(self, excel_manager: ExcelManager):
        """Initialize workbook manager.

        Args:
            excel_manager: An ExcelManager instance (must be started)

        Example:
            >>> with ExcelManager() as excel_mgr:
            ...     wb_mgr = WorkbookManager(excel_mgr)
            ...     info = wb_mgr.open(Path("data.xlsx"))
        """
        self._mgr = excel_manager

    def open(self, path: Path, read_only: bool = False) -> WorkbookInfo:
        """Open an existing workbook.

        Opens a workbook file and returns information about it.
        If the workbook is already open, raises an error.

        Args:
            path: Path to the Excel file to open
            read_only: If True, open in read-only mode (default: False)

        Returns:
            WorkbookInfo with details about the opened workbook

        Raises:
            WorkbookNotFoundError: If the file doesn't exist on disk
            WorkbookAlreadyOpenError: If the workbook is already open
            ExcelConnectionError: If COM connection fails

        Example:
            >>> manager = WorkbookManager(excel_mgr)
            >>> info = manager.open(Path("C:/data/sales.xlsx"), read_only=True)
            >>> print(f"Opened {info.name} with {info.sheets_count} sheets")
        """
        # Step 1: Verify file exists
        if not path.exists():
            raise WorkbookNotFoundError(path, f"File not found: {path}")

        # Step 2: Check if already open
        app = self._mgr.app  # Will raise if Excel not started
        existing_wb = _find_open_workbook(app, path)
        if existing_wb is not None:
            raise WorkbookAlreadyOpenError(
                path,
                existing_wb.Name,
                f"Workbook is already open: {existing_wb.Name}",
            )

        # Step 3: Open the workbook
        try:
            # Convert Path to string and resolve to absolute path
            abs_path = str(path.resolve())

            # Open via COM
            wb = app.Workbooks.Open(abs_path, ReadOnly=read_only)

            # Step 4: Build WorkbookInfo
            info = WorkbookInfo(
                name=wb.Name,
                full_path=Path(wb.FullName),
                read_only=wb.ReadOnly,
                saved=wb.Saved,
                sheets_count=wb.Worksheets.Count,
            )

            return info

        except Exception as e:
            # Wrap COM errors
            if hasattr(e, "hresult"):
                raise ExcelConnectionError(
                    getattr(e, "hresult"),
                    f"Failed to open workbook: {str(e)}",
                ) from e
            else:
                # Re-raise non-COM exceptions
                raise

    def create(self, path: Path, template: Path | None = None) -> WorkbookInfo:
        """Create a new workbook.

        Creates a new Excel workbook and saves it to the specified path.
        Optionally uses a template file as starting point.

        Args:
            path: Destination path for the new workbook
            template: Optional path to a template file (.xltx, .xltm, or .xlsx/.xlsm)

        Returns:
            WorkbookInfo with details about the created workbook

        Raises:
            WorkbookNotFoundError: If template file doesn't exist
            WorkbookSaveError: If save operation fails
            ExcelConnectionError: If COM connection fails

        Examples:
            >>> # Create blank workbook
            >>> info = manager.create(Path("C:/data/new.xlsx"))

            >>> # Create from template
            >>> template = Path("C:/templates/report.xltx")
            >>> info = manager.create(Path("C:/data/report.xlsx"), template=template)

        Note:
            The file format is automatically detected from the path extension.
            Supported formats: .xlsx, .xlsm, .xls, .xlsb
        """
        # Step 1: Validate template if provided
        if template is not None:
            if not template.exists():
                raise WorkbookNotFoundError(
                    template,
                    f"Template file not found: {template}",
                )

        # Step 2: Detect target file format
        try:
            file_format = _detect_file_format(path)
        except ValueError as e:
            raise WorkbookSaveError(
                path,
                message=f"Invalid file extension: {str(e)}",
            ) from e

        # Step 3: Create workbook
        app = self._mgr.app

        try:
            if template is None:
                # Create blank workbook
                wb = app.Workbooks.Add()
            else:
                # Create from template
                wb = app.Workbooks.Add(str(template.resolve()))

            # Step 4: Save to destination path
            abs_path = str(path.resolve())

            # Step 4: Save to destination path and build info
            try:
                wb.SaveAs(abs_path, FileFormat=file_format)

                # Step 5: Build WorkbookInfo - only reached if SaveAs succeeds
                info = WorkbookInfo(
                    name=wb.Name,
                    full_path=Path(wb.FullName),
                    read_only=wb.ReadOnly,
                    saved=wb.Saved,
                    sheets_count=wb.Worksheets.Count,
                )
                return info

            except Exception as e:
                # Clean up the unsaved workbook
                try:
                    wb.Close(SaveChanges=False)
                    del wb
                except Exception:
                    pass

                # Raise save error
                if hasattr(e, "hresult"):
                    raise WorkbookSaveError(
                        path,
                        hresult=getattr(e, "hresult"),
                        message=f"Failed to save workbook: {str(e)}",
                    ) from e
                else:
                    raise WorkbookSaveError(
                        path,
                        message=f"Failed to save workbook: {str(e)}",
                    ) from e

        except WorkbookSaveError:
            # Re-raise our own exceptions
            raise
        except Exception as e:
            # Wrap other COM errors
            if hasattr(e, "hresult"):
                raise ExcelConnectionError(
                    getattr(e, "hresult"),
                    f"Failed to create workbook: {str(e)}",
                ) from e
            else:
                raise

    def close(self, path: Path, save: bool = True, force: bool = False) -> None:
        """Close an open workbook.

        Closes a workbook that is currently open in Excel.
        Optionally saves changes before closing.

        Args:
            path: Path to the workbook to close
            save: If True, save changes before closing (default: True)
            force: If True, suppress confirmation dialogs (default: False)

        Raises:
            WorkbookNotFoundError: If the workbook is not currently open
            ExcelConnectionError: If COM connection fails

        Example:
            >>> # Close without saving
            >>> manager.close(Path("C:/data/temp.xlsx"), save=False)

            >>> # Close with save, no dialogs
            >>> manager.close(Path("C:/data/work.xlsx"), save=True, force=True)

        Note:
            If save=True and the workbook has never been saved,
            Excel may still show a "Save As" dialog unless force=True.
        """
        app = self._mgr.app

        # Step 1: Find the open workbook
        wb = _find_open_workbook(app, path)
        if wb is None:
            raise WorkbookNotFoundError(
                path,
                f"Workbook is not open: {path.name}",
            )

        # Step 2: Configure alerts
        if force:
            app.DisplayAlerts = False

        try:
            # Step 3: Close the workbook
            wb.Close(SaveChanges=save)

            # Step 4: Clean up COM reference
            del wb

        finally:
            # Step 5: Restore alerts
            if force:
                app.DisplayAlerts = True

    def save(self, path: Path, output: Path | None = None) -> None:
        """Save a workbook.

        Saves an open workbook. Can save to the same file (Save)
        or to a different file (SaveAs).

        Args:
            path: Path to the currently open workbook
            output: Optional destination path for SaveAs.
                    If None, saves to the current file (Save).

        Raises:
            WorkbookNotFoundError: If the workbook is not currently open
            WorkbookSaveError: If save operation fails
            ExcelConnectionError: If COM connection fails

        Examples:
            >>> # Save to current file
            >>> manager.save(Path("C:/data/work.xlsx"))

            >>> # Save to different file (SaveAs)
            >>> manager.save(
            ...     Path("C:/data/work.xlsx"),
            ...     output=Path("C:/backup/work_v2.xlsx")
            ... )

        Note:
            When using SaveAs with output parameter, the file format
            is automatically detected from the output file extension.
        """
        app = self._mgr.app

        # Step 1: Find the open workbook
        wb = _find_open_workbook(app, path)
        if wb is None:
            raise WorkbookNotFoundError(
                path,
                f"Workbook is not open: {path.name}",
            )

        try:
            if output is None:
                # Step 2a: Save to current file
                wb.Save()
            else:
                # Step 2b: SaveAs to different file

                # Detect file format from output extension
                try:
                    file_format = _detect_file_format(output)
                except ValueError as e:
                    raise WorkbookSaveError(
                        output,
                        message=f"Invalid file extension: {str(e)}",
                    ) from e

                # Convert to absolute path
                abs_path = str(output.resolve())

                # Save with format
                wb.SaveAs(abs_path, FileFormat=file_format)

        except WorkbookSaveError:
            # Re-raise our exceptions
            raise
        except Exception as e:
            target = output if output is not None else path
            if hasattr(e, "hresult"):
                raise WorkbookSaveError(
                    target,
                    hresult=getattr(e, "hresult"),
                    message=f"Failed to save workbook: {str(e)}",
                ) from e
            else:
                raise WorkbookSaveError(
                    target,
                    message=f"Failed to save workbook: {str(e)}",
                ) from e

    def list(self) -> list[WorkbookInfo]:
        """List all open workbooks.

        Returns information about all workbooks currently open
        in the Excel instance.

        Returns:
            List of WorkbookInfo for each open workbook.
            Returns empty list if no workbooks are open.

        Raises:
            ExcelConnectionError: If COM connection fails

        Example:
            >>> manager = WorkbookManager(excel_mgr)
            >>> workbooks = manager.list()
            >>> for wb in workbooks:
            ...     print(f"{wb.name}: {wb.sheets_count} sheets")
        """
        app = self._mgr.app
        workbooks = []

        for wb in app.Workbooks:
            try:
                info = WorkbookInfo(
                    name=wb.Name,
                    full_path=Path(wb.FullName),
                    read_only=wb.ReadOnly,
                    saved=wb.Saved,
                    sheets_count=wb.Worksheets.Count,
                )
                workbooks.append(info)
            except Exception:
                continue

        return workbooks
