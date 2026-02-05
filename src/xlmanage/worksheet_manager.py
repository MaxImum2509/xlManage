"""
Worksheet information and validation for xlmanage.

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
