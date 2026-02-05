"""
Worksheet management for xlmanage.

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

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    from win32com.client import CDispatch
except ImportError:
    CDispatch = Any

try:
    from .exceptions import WorksheetNameError
except ImportError:
    from xlmanage.exceptions import WorksheetNameError


SHEET_NAME_MAX_LENGTH: int = 31
SHEET_NAME_FORBIDDEN_CHARS: str = r"\\/\*\?:\[\]"


@dataclass
class WorksheetInfo:
    """Information about an Excel worksheet.

    Attributes:
        name: Name of the worksheet (e.g., "Sheet1")
        index: Position in the workbook (1-based as in Excel)
        visible: Whether the worksheet is visible to the user
        rows_used: Number of rows containing data
        columns_used: Number of columns containing data
    """

    name: str
    index: int
    visible: bool
    rows_used: int
    columns_used: int


def _validate_sheet_name(name: str) -> None:
    """Validate an Excel worksheet name.

    Checks that the name follows Excel's naming rules:
    - Not empty
    - Maximum 31 characters
    - No forbidden characters: \\ / * ? : [ ]

    Args:
        name: The worksheet name to validate

    Raises:
        WorksheetNameError: If the name violates any rule

    Examples:
        >>> _validate_sheet_name("Sheet1")  # OK
        >>> _validate_sheet_name("Data-2024_Q1")  # OK
        >>> _validate_sheet_name("A" * 32)  # Raises: too long
        >>> _validate_sheet_name("Sheet/1")  # Raises: forbidden char
    """
    if not name or not name.strip():
        raise WorksheetNameError(name, "name cannot be empty")

    if len(name) > SHEET_NAME_MAX_LENGTH:
        raise WorksheetNameError(
            name,
            f"name exceeds {SHEET_NAME_MAX_LENGTH} characters (length: {len(name)})",
        )

    forbidden_pattern = f"[{SHEET_NAME_FORBIDDEN_CHARS}]"
    match = re.search(forbidden_pattern, name)
    if match:
        forbidden_char = match.group(0)
        raise WorksheetNameError(
            name, f"contains forbidden character '{forbidden_char}'"
        )


def _resolve_workbook(app: CDispatch, workbook: Path | None) -> CDispatch:
    """Resolve the target workbook.

    If workbook is provided, finds or opens that specific workbook.
    If workbook is None, returns the active workbook.

    Args:
        app: Excel Application COM object
        workbook: Optional path to a specific workbook.
                  If None, uses the active workbook.

    Returns:
        Workbook COM object

    Raises:
        WorkbookNotFoundError: If the specified workbook is not open
        ExcelConnectionError: If no active workbook when workbook=None

    Examples:
        >>> # Use active workbook
        >>> wb = _resolve_workbook(app, None)

        >>> # Use specific workbook
        >>> wb = _resolve_workbook(app, Path("C:/data/test.xlsx"))

    Note:
        This function does NOT open the workbook if it's not already open.
        Use WorkbookManager.open() to open a workbook first.
        This function is shared by WorksheetManager, TableManager, and VBAManager.
    """
    if workbook is None:
        # Use active workbook
        try:
            wb = app.ActiveWorkbook
            if wb is None:
                from .exceptions import ExcelConnectionError

                raise ExcelConnectionError(
                    0x80080005, "No active workbook. Open a workbook first."
                )
            return wb
        except Exception as e:
            from .exceptions import ExcelConnectionError

            if hasattr(e, "hresult"):
                raise ExcelConnectionError(
                    getattr(e, "hresult"),
                    f"Failed to get active workbook: {str(e)}",
                ) from e
            else:
                raise
    else:
        # Find the specified workbook
        from .exceptions import WorkbookNotFoundError
        from .workbook_manager import _find_open_workbook

        wb = _find_open_workbook(app, workbook)
        if wb is None:
            raise WorkbookNotFoundError(
                workbook, f"Workbook is not open: {workbook.name}"
            )
        return wb


def _find_worksheet(wb: CDispatch, name: str) -> CDispatch | None:
    """Find a worksheet by name in a workbook.

    Searches for a worksheet with the given name.
    The search is case-insensitive (Excel behavior).

    Args:
        wb: Workbook COM object to search in
        name: Name of the worksheet to find

    Returns:
        Worksheet COM object if found, None otherwise

    Examples:
        >>> ws = _find_worksheet(wb, "Sheet1")
        >>> if ws:
        ...     print(f"Found: {ws.Name}")

        >>> # Case-insensitive
        >>> ws = _find_worksheet(wb, "SHEET1")  # Finds "Sheet1"

    Note:
        Excel worksheet names are case-insensitive but case-preserving.
        "Sheet1" and "SHEET1" refer to the same worksheet.
    """
    # Normalize search name to lowercase
    search_name = name.lower()

    # Iterate through all worksheets
    for ws in wb.Worksheets:
        try:
            # Compare case-insensitive
            if ws.Name.lower() == search_name:
                return ws
        except Exception:
            # Skip worksheets that can't be read
            continue

    return None


class WorksheetManager:
    """Manager for Excel worksheet CRUD operations.

    This class provides methods to create, delete, list, and copy
    worksheets. It depends on ExcelManager for COM access.

    Note:
        The ExcelManager instance must be started before using this manager.
    """

    def __init__(self, excel_manager):
        """Initialize worksheet manager.

        Args:
            excel_manager: An ExcelManager instance (must be started)

        Example:
            >>> with ExcelManager() as excel_mgr:
            ...     ws_mgr = WorksheetManager(excel_mgr)
            ...     info = ws_mgr.create("NewSheet")
        """
        self._mgr = excel_manager

    def _get_worksheet_info(self, ws: CDispatch) -> WorksheetInfo:
        """Extract information from a worksheet COM object.

        Args:
            ws: Worksheet COM object

        Returns:
            WorksheetInfo with worksheet details

        Note:
            If UsedRange fails (empty sheet), defaults to 0 rows/columns.
        """
        try:
            # Get used range to count rows/columns
            used_range = ws.UsedRange
            if used_range is not None:
                rows_used = used_range.Rows.Count
                columns_used = used_range.Columns.Count
            else:
                rows_used = 0
                columns_used = 0
        except Exception:
            # If UsedRange fails (empty sheet), default to 0
            rows_used = 0
            columns_used = 0

        return WorksheetInfo(
            name=ws.Name,
            index=ws.Index,
            visible=ws.Visible,
            rows_used=rows_used,
            columns_used=columns_used,
        )

    def create(self, name: str, workbook: Path | None = None) -> WorksheetInfo:
        """Create a new worksheet.

        Creates a new worksheet with the specified name in the target workbook.
        The worksheet is added at the end of the workbook.

        Args:
            name: Name for the new worksheet (must follow Excel naming rules)
            workbook: Optional path to target workbook.
                      If None, uses the active workbook.

        Returns:
            WorksheetInfo with details about the created worksheet

        Raises:
            WorksheetNameError: If the name is invalid
            WorksheetAlreadyExistsError: If a worksheet with this name already exists
            ExcelConnectionError: If COM connection fails
            WorkbookNotFoundError: If the specified workbook is not open

        Examples:
            >>> # Create in active workbook
            >>> manager = WorksheetManager(excel_mgr)
            >>> info = manager.create("Summary")
            >>> print(f"Created: {info.name} at index {info.index}")

            >>> # Create in specific workbook
            >>> info = manager.create("Data", Path("C:/work/report.xlsx"))

        Note:
            The worksheet is created at the end of the workbook.
            Use move() to reposition it if needed.
        """
        # Step 1: Validate the sheet name
        _validate_sheet_name(name)

        # Step 2: Get Excel app and resolve target workbook
        app = self._mgr.app
        wb = _resolve_workbook(app, workbook)

        # Step 3: Check if worksheet already exists
        existing = _find_worksheet(wb, name)
        if existing is not None:
            from .exceptions import WorksheetAlreadyExistsError

            raise WorksheetAlreadyExistsError(name, wb.Name)

        # Step 4: Create the worksheet at the end
        try:
            # Get the last worksheet
            last_ws = wb.Worksheets(wb.Worksheets.Count)

            # Add new worksheet after the last one
            ws = wb.Worksheets.Add(After=last_ws)

            # Set the name
            ws.Name = name

            # Step 5: Return WorksheetInfo
            return self._get_worksheet_info(ws)

        except Exception as e:
            from .exceptions import ExcelConnectionError

            if hasattr(e, "hresult"):
                raise ExcelConnectionError(
                    getattr(e, "hresult"),
                    f"Failed to create worksheet: {str(e)}",
                ) from e
            else:
                raise

    def delete(self, name: str, workbook: Path | None = None) -> None:
        """Delete a worksheet.

        Deletes the specified worksheet from the workbook.
        Excel always shows a confirmation dialog unless DisplayAlerts is disabled.

        Args:
            name: Name of the worksheet to delete
            workbook: Optional path to the target workbook.
                      If None, uses the active workbook.

        Raises:
            WorksheetNotFoundError: If the worksheet doesn't exist
            WorksheetDeleteError: If the worksheet cannot be deleted
            WorkbookNotFoundError: If the specified workbook is not open
            ExcelConnectionError: If COM connection fails

        Examples:
            >>> # Delete a worksheet
            >>> manager = WorksheetManager(excel_mgr)
            >>> manager.delete("OldSheet")

            >>> # Delete from specific workbook
            >>> manager.delete("TempData", Path("C:/work/report.xlsx"))

        Warning:
            You cannot delete the last visible worksheet in a workbook.
            Excel requires at least one visible worksheet.

        Note:
            DisplayAlerts is ALWAYS set to False to prevent Excel dialogs.
            The parameter is automatically managed and restored.
        """
        # Step 1: Resolve target workbook
        app = self._mgr.app
        wb = _resolve_workbook(app, workbook)

        # Step 2: Find the worksheet
        ws = _find_worksheet(wb, name)
        if ws is None:
            from .exceptions import WorksheetNotFoundError

            raise WorksheetNotFoundError(name, wb.Name)

        # Step 3: Check if it's the last visible sheet
        visible_count = 0
        for sheet in wb.Worksheets:
            try:
                if sheet.Visible:
                    visible_count += 1
                    if visible_count > 1:
                        break  # We have at least 2 visible sheets
            except Exception:
                continue

        if visible_count == 1 and ws.Visible:
            from .exceptions import WorksheetDeleteError

            raise WorksheetDeleteError(name, "cannot delete the last visible worksheet")

        # Step 4: Delete the worksheet
        # CRITICAL: DisplayAlerts MUST be False to avoid Excel dialog
        app.DisplayAlerts = False

        try:
            ws.Delete()
            # Clean up COM reference
            del ws
        finally:
            # Always restore DisplayAlerts
            app.DisplayAlerts = True

    def list(self, workbook: Path | None = None) -> list[WorksheetInfo]:
        """List all worksheets in a workbook.

        Returns information about all worksheets in the workbook,
        including hidden worksheets.

        Args:
            workbook: Optional path to the target workbook.
                      If None, uses the active workbook.

        Returns:
            List of WorksheetInfo for each worksheet.
            Returns empty list if workbook has no worksheets.

        Raises:
            WorkbookNotFoundError: If the specified workbook is not open
            ExcelConnectionError: If COM connection fails

        Examples:
            >>> manager = WorksheetManager(excel_mgr)
            >>> sheets = manager.list()
            >>> for sheet in sheets:
            ...     print(f"{sheet.index}. {sheet.name} ({sheet.rows_used} rows)")

            >>> # List from specific workbook
            >>> sheets = manager.list(Path("C:/work/report.xlsx"))

        Note:
            The list includes both visible and hidden worksheets.
            Hidden worksheets have visible=False.
        """
        app = self._mgr.app
        wb = _resolve_workbook(app, workbook)

        worksheets = []

        # Iterate through all worksheets
        for ws in wb.Worksheets:
            try:
                info = self._get_worksheet_info(ws)
                worksheets.append(info)
            except Exception:
                # Skip worksheets that can't be read
                continue

        return worksheets

    def copy(
        self, source: str, destination: str, workbook: Path | None = None
    ) -> WorksheetInfo:
        """Copy a worksheet and rename the copy.

        Creates a duplicate of the source worksheet and gives it a new name.
        The copy is placed immediately after the source worksheet.

        Args:
            source: Name of the worksheet to copy
            destination: Name for the copy
            workbook: Optional path to the target workbook.
                      If None, uses the active workbook.

        Returns:
            WorksheetInfo of the newly created copy

        Raises:
            WorksheetNotFoundError: If source worksheet doesn't exist
            WorksheetNameError: If destination name is invalid
            WorksheetAlreadyExistsError: If destination name already exists
            WorkbookNotFoundError: If the specified workbook is not open
            ExcelConnectionError: If COM connection fails

        Examples:
            >>> manager = WorksheetManager(excel_mgr)
            >>> info = manager.copy("Template", "January_Report")
            >>> print(f"Created copy: {info.name} at position {info.index}")

            >>> # Copy in specific workbook
            >>> path = Path("C:/work/data.xlsx")
            >>> info = manager.copy("Sheet1", "Sheet1_Backup", path)

        Note:
            Excel automatically activates the newly created copy.
            The copy contains all data, formatting, and formulas from the source.
        """
        # Step 1: Validate destination name
        _validate_sheet_name(destination)

        # Step 2: Resolve target workbook
        app = self._mgr.app
        wb = _resolve_workbook(app, workbook)

        # Step 3: Find source worksheet
        ws_source = _find_worksheet(wb, source)
        if ws_source is None:
            from .exceptions import WorksheetNotFoundError

            raise WorksheetNotFoundError(source, wb.Name)

        # Step 4: Check destination name doesn't exist
        from .exceptions import WorksheetAlreadyExistsError

        ws_existing = _find_worksheet(wb, destination)
        if ws_existing is not None:
            raise WorksheetAlreadyExistsError(destination, wb.Name)

        # Step 5: Copy the worksheet
        try:
            # Copy after the source worksheet
            ws_source.Copy(After=ws_source)

            # The copied worksheet becomes the active sheet
            ws_copy = wb.ActiveSheet

            # Rename the copy
            ws_copy.Name = destination

            # Step 6: Get worksheet information
            info = self._get_worksheet_info(ws_copy)

            return info

        except WorksheetNameError:
            # Re-raise our own exceptions
            raise
        except WorksheetAlreadyExistsError:
            raise
        except Exception as e:
            # Wrap COM errors
            if hasattr(e, "hresult"):
                from .exceptions import ExcelConnectionError

                raise ExcelConnectionError(
                    getattr(e, "hresult"),
                    f"Failed to copy worksheet '{source}': {str(e)}",
                ) from e
            else:
                raise
