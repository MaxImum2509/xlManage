"""
Table management operations for Excel workbooks.

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

from .exceptions import TableAlreadyExistsError, TableNameError, TableRangeError
from .worksheet_manager import _find_worksheet, _resolve_workbook

# Excel table name constraints
TABLE_NAME_MAX_LENGTH: int = 255
# Must start with letter or underscore, contains only alphanumeric and underscores
TABLE_NAME_PATTERN: str = r"^[a-zA-Z_][a-zA-Z0-9_]*$"


@dataclass
class TableInfo:
    """Information about an Excel table (ListObject).

    Attributes:
        name: Name of the table (e.g., "tbl_Sales")
        worksheet_name: Name of the worksheet containing the table
        range_ref: Range reference (e.g., "A1:D100")
        header_row_range: Range of the header row
        rows_count: Number of data rows (excluding header)
    """

    name: str
    worksheet_name: str
    range_ref: str
    header_row_range: str
    rows_count: int


def _validate_table_name(name: str) -> None:
    """Validate an Excel table name.

    Checks that the name follows Excel's naming rules.

    Args:
        name: The table name to validate

    Raises:
        TableNameError: If the name violates any rule

    Examples:
        >>> _validate_table_name("tbl_Sales")  # OK
        >>> _validate_table_name("Data_2024")  # OK
        >>> _validate_table_name("A" * 256)  # Raises: too long
        >>> _validate_table_name("1Data")  # Raises: starts with digit
    """
    # Rule 1: Name cannot be empty
    if not name or not name.strip():
        raise TableNameError(name, "name cannot be empty")

    # Rule 2: Maximum 255 characters
    if len(name) > TABLE_NAME_MAX_LENGTH:
        raise TableNameError(
            name,
            f"name exceeds {TABLE_NAME_MAX_LENGTH} characters (length: {len(name)})",
        )

    # Rule 3: Must match pattern (start with letter or _, only alphanumeric and _)
    if not re.match(TABLE_NAME_PATTERN, name):
        raise TableNameError(
            name,
            "must start with letter or underscore, "
            "contain only alphanumeric characters and underscores",
        )

    # Rule 4: Cannot be a cell reference
    if re.match(r"^[A-Z]+\d+$|^[rR]\d+[cC]\d+$", name):
        raise TableNameError(name, "cannot be a cell reference")


def _find_table(ws: "CDispatch", name: str) -> "CDispatch | None":
    """Find a table by name in a worksheet.

    Searches for a table (ListObject) with the given name.
    Note: Table names are case-SENSITIVE in Excel.

    Args:
        ws: Worksheet COM object to search in
        name: Name of the table to find

    Returns:
        Table COM object if found, None otherwise

    Examples:
        >>> table = _find_table(ws, "tbl_Sales")
        >>> if table:
        ...     print(f"Found: {table.Name}")

    Note:
        Unlike worksheet names, Excel table names are case-sensitive.
        "tbl_Sales" and "TBL_SALES" are different tables.
    """
    # Iterate through all tables in the worksheet
    for table in ws.ListObjects:
        try:
            if table.Name == name:  # Case-sensitive comparison
                return table
        except Exception:
            # Skip tables that can't be read
            continue

    return None


def _validate_range(range_ref: str) -> None:
    """Validate an Excel range reference.

    Checks that the range has valid syntax and structure.

    Args:
        range_ref: Range reference to validate (e.g., "A1:D10")

    Raises:
        TableRangeError: If the range is invalid

    Examples:
        >>> _validate_range("A1:D10")  # OK
        >>> _validate_range("Sheet1!A1:D10")  # OK
        >>> _validate_range("A1:Z")  # Raises: invalid syntax
    """
    if not range_ref or not range_ref.strip():
        raise TableRangeError(range_ref, "range cannot be empty")

    # Remove sheet reference if present (e.g., "Sheet1!" or "'Sheet Name'!")
    clean_range = range_ref
    if "!" in clean_range:
        parts = clean_range.split("!", 1)
        if len(parts) == 2:
            clean_range = parts[1]

    # Must contain at least one colon (for start:end range)
    if ":" not in clean_range:
        raise TableRangeError(range_ref, "range must have format A1:Z99")

    # Basic pattern check for Excel ranges
    pattern = r"^[A-Z]+\d+:[A-Z]+\d+$|^[rR]\d+[cC]\d+:[rR]\d+[cC]\d+$"
    if not re.match(pattern, clean_range.replace("$", "")):
        raise TableRangeError(range_ref, "invalid range syntax")


class TableManager:
    """Manager for Excel table (ListObject) CRUD operations.

    This class provides methods to create, delete, and list tables.
    It depends on ExcelManager for COM access.

    Note:
        The ExcelManager instance must be started before using this manager.
    """

    def __init__(self, excel_manager):
        """Initialize table manager.

        Args:
            excel_manager: An ExcelManager instance (must be started)

        Example:
            >>> with ExcelManager() as excel_mgr:
            ...     table_mgr = TableManager(excel_mgr)
            ...     info = table_mgr.create("tbl_Sales", "A1:D100", worksheet="Data")
        """
        self._mgr = excel_manager

    def _get_table_info(self, table: "CDispatch", ws: "CDispatch") -> TableInfo:
        """Extract information from a table COM object.

        Args:
            table: Table COM object
            ws: Worksheet COM object

        Returns:
            TableInfo with table details
        """
        return TableInfo(
            name=table.Name,
            worksheet_name=ws.Name,
            range_ref=table.Range.Address,
            header_row_range=table.HeaderRowRange.Address,
            rows_count=table.DataBodyRange.Rows.Count if table.DataBodyRange else 0,
        )

    def create(
        self,
        name: str,
        range_ref: str,
        worksheet: str | None = None,
        workbook: Path | None = None,
    ) -> TableInfo:
        """Create a new table in a worksheet.

        Args:
            name: Name for the new table (e.g., "tbl_Sales")
            range_ref: Range reference (e.g., "A1:D100")
            worksheet: Worksheet name (if None, uses active worksheet)
            workbook: Workbook path (if None, uses active workbook)

        Returns:
            TableInfo with details of the created table

        Raises:
            TableNameError: If the table name is invalid
            TableRangeError: If the range is invalid
            TableAlreadyExistsError: If a table with this name already exists
            WorksheetNotFoundError: If the worksheet doesn't exist
            WorkbookNotFoundError: If the workbook is not open

        Examples:
            >>> manager = TableManager(excel_mgr)
            >>> info = manager.create("tbl_Sales", "A1:D100", worksheet="Data")
            >>> print(f"{info.name}: {info.rows_count} rows")
        """
        # Validate table name
        _validate_table_name(name)

        # Validate range
        _validate_range(range_ref)

        # Resolve workbook and worksheet
        wb = _resolve_workbook(self._mgr.app, workbook)

        if worksheet is None:
            ws = wb.ActiveSheet
        else:
            ws = _find_worksheet(wb, worksheet)

        # Check if table name already exists in workbook
        for sheet in wb.Worksheets:
            for existing_table in sheet.ListObjects:
                if existing_table.Name == name:
                    raise TableAlreadyExistsError(name, wb.Name)

        # Create the table
        table = ws.ListObjects.Add(
            SourceType=1,  # xlSrcRange
            Source=ws.Range(range_ref),
            XlListObjectHasHeaders=1,  # xlYes
        )
        table.Name = name

        return self._get_table_info(table, ws)
