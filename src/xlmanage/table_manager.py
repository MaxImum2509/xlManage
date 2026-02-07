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

from .exceptions import (
    TableAlreadyExistsError,
    TableNameError,
    TableNotFoundError,
    TableRangeError,
)
from .worksheet_manager import _find_worksheet, _resolve_workbook

# Excel table name constraints
TABLE_NAME_MAX_LENGTH: int = 255
# Must start with letter or underscore, contains only alphanumeric and underscores
TABLE_NAME_PATTERN: str = r"^[A-Za-z_][A-Za-z0-9_]*$"


@dataclass
class TableInfo:
    """Information about an Excel table (ListObject).

    Attributes:
        name: Name of the table (e.g., "tbl_Sales")
        worksheet_name: Name of the worksheet containing the table
        range_address: Range address (e.g., "$A$1:$D$100")
        columns: List of column header names
        rows_count: Number of data rows (excluding header)
        header_row: Address of the header row (e.g., "$A$1:$D$1")
    """

    name: str
    worksheet_name: str
    range_address: str
    columns: list[str]
    rows_count: int
    header_row: str


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


def _find_table(wb: "CDispatch", name: str) -> "tuple[CDispatch, CDispatch] | None":
    """Find a table by name in a workbook (searches all worksheets).

    Table names are unique across the entire workbook, not just within sheets.
    Searches all worksheets in the workbook for a table with the given name.

    Args:
        wb: Workbook COM object to search in
        name: Name of the table to find

    Returns:
        Tuple of (worksheet, table) if found, None otherwise

    Examples:
        >>> result = _find_table(wb, "tbl_Sales")
        >>> if result:
        ...     ws, table = result
        ...     print(f"Found {table.Name} in {ws.Name}")

    Note:
        Table names are case-SENSITIVE in Excel.
        "tbl_Sales" and "TBL_SALES" are different tables.
    """
    # Iterate through all worksheets in the workbook
    for ws in wb.Worksheets:
        try:
            # Iterate through all tables in the worksheet
            for table in ws.ListObjects:
                try:
                    if table.Name == name:  # Case-sensitive comparison
                        return (ws, table)
                except Exception:
                    # Skip tables that can't be read
                    continue
        except Exception:
            # Skip sheets that can't be read
            continue

    return None


def _ranges_overlap(range1: "CDispatch", range2: "CDispatch") -> bool:
    """Check if two COM Range objects overlap using Application.Intersect.

    Uses Excel's built-in Intersect method to determine overlap.

    Args:
        range1: First Range COM object
        range2: Second Range COM object

    Returns:
        True if ranges overlap, False otherwise

    Examples:
        >>> overlap = _ranges_overlap(ws.Range("A1:D10"), ws.Range("C5:F15"))
        >>> print(overlap)  # True (overlaps at C5:D10)
    """
    try:
        app = range1.Application
        intersection = app.Intersect(range1, range2)
        return intersection is not None
    except Exception:
        return False


def _validate_range(ws: "CDispatch", range_ref: str) -> "CDispatch":
    """Validate an Excel range reference and return the COM Range object.

    Validates both syntax (via ws.Range) and checks for overlap with
    existing tables in the worksheet.

    Args:
        ws: Worksheet COM object
        range_ref: Range reference to validate (e.g., "A1:D10")

    Returns:
        Validated Range COM object

    Raises:
        TableRangeError: If the range is invalid or overlaps with an existing table

    Examples:
        >>> range_obj = _validate_range(ws, "A1:D10")
        >>> print(range_obj.Address)  # "$A$1:$D$10"
    """
    if not range_ref or not range_ref.strip():
        raise TableRangeError(range_ref, "range cannot be empty")

    # Attempt to create Range object (validates syntax)
    try:
        range_obj = ws.Range(range_ref)
    except Exception:
        raise TableRangeError(range_ref, "invalid range syntax")

    # Check for overlap with existing tables
    for table in ws.ListObjects:
        try:
            existing_range = table.Range
            if _ranges_overlap(range_obj, existing_range):
                raise TableRangeError(
                    range_ref, f"range overlaps with existing table '{table.Name}'"
                )
        except TableRangeError:
            raise
        except Exception:
            # Skip tables that can't be read
            continue

    return range_obj


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
            TableInfo with table details including column names
        """
        # Extract column names from table headers
        columns = [col.Name for col in table.ListColumns]

        return TableInfo(
            name=table.Name,
            worksheet_name=ws.Name,
            range_address=table.Range.Address,
            columns=columns,
            rows_count=table.DataBodyRange.Rows.Count if table.DataBodyRange else 0,
            header_row=table.HeaderRowRange.Address,
        )

    def create(
        self,
        name: str,
        range_ref: str,
        workbook: Path | None = None,
        worksheet: str | None = None,
    ) -> TableInfo:
        """Create a new table in a worksheet.

        Args:
            name: Name for the new table (e.g., "tbl_Sales")
            range_ref: Range reference (e.g., "A1:D100")
            workbook: Workbook path (if None, uses active workbook)
            worksheet: Worksheet name (if None, uses active worksheet)

        Returns:
            TableInfo with details of the created table

        Raises:
            TableNameError: If the table name is invalid
            TableRangeError: If the range is invalid or overlaps
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

        # Resolve workbook and worksheet
        wb = _resolve_workbook(self._mgr.app, workbook)

        if worksheet is None:
            ws = wb.ActiveSheet
        else:
            ws = _find_worksheet(wb, worksheet)

        # Check if table name already exists in workbook
        # Use the new _find_table() that searches the entire workbook
        if _find_table(wb, name) is not None:
            raise TableAlreadyExistsError(name, wb.Name)

        # Validate range (checks syntax and overlap with existing tables)
        range_obj = _validate_range(ws, range_ref)

        # Create the table
        table = ws.ListObjects.Add(
            SourceType=1,  # xlSrcRange
            Source=range_obj,
            XlListObjectHasHeaders=1,  # xlYes
        )
        table.Name = name

        return self._get_table_info(table, ws)

    def delete(
        self,
        name: str,
        workbook: Path | None = None,
        worksheet: str | None = None,
        force: bool = False,
    ) -> None:
        """Delete a table.

        By default (force=False), removes the table structure but keeps the data.
        With force=True, deletes both the table structure and the data.

        Args:
            name: Name of the table to delete
            workbook: Target workbook path (if None, uses active workbook)
            worksheet: Worksheet containing the table (if None, search all)
            force: If True, deletes table and data; if False, only removes
                   table structure (keeps data as normal range)

        Raises:
            TableNotFoundError: If the table doesn't exist
            WorkbookNotFoundError: If the specified workbook is not open
            ExcelConnectionError: If COM connection fails

        Examples:
            >>> manager = TableManager(excel_mgr)
            >>> manager.delete("tbl_Sales")  # Keeps data, removes structure
            >>> manager.delete("tbl_Old", force=True)  # Deletes everything
        """
        # Resolve workbook
        wb = _resolve_workbook(self._mgr.app, workbook)

        # Search for table using new signature that searches entire workbook
        result = None

        if worksheet is None:
            # Search all worksheets using new _find_table(wb, name)
            result = _find_table(wb, name)
        else:
            # Search specific worksheet
            ws = _find_worksheet(wb, worksheet)
            if ws is not None:
                # Check if table exists in this specific worksheet
                for table in ws.ListObjects:
                    try:
                        if table.Name == name:
                            result = (ws, table)
                            break
                    except Exception:
                        continue

        if result is None:
            worksheet_context = worksheet if worksheet else "any worksheet"
            raise TableNotFoundError(name, worksheet_context)

        _ws, table_found = result

        # Delete the table
        if force:
            # Delete table structure AND data
            table_found.Delete()
        else:
            # Remove table structure but keep data
            table_found.Unlist()

    def list(
        self,
        worksheet: str | None = None,
        workbook: Path | None = None,
    ) -> list[TableInfo]:
        """List all tables.

        Returns information about all tables in the worksheet(s).

        Args:
            worksheet: Worksheet name to search (if None, list all in workbook)
            workbook: Target workbook path (if None, uses active workbook)

        Returns:
            List of TableInfo for each table
            Returns empty list if no tables found

        Raises:
            WorkbookNotFoundError: If the specified workbook is not open
            ExcelConnectionError: If COM connection fails

        Examples:
            >>> manager = TableManager(excel_mgr)
            >>> tables = manager.list(worksheet="Data")
            >>> for table in tables:
            ...     print(f"{table.name}: {table.rows_count} rows")
        """
        # Resolve workbook
        wb = _resolve_workbook(self._mgr.app, workbook)

        tables = []

        if worksheet is None:
            # List all tables in workbook
            for sheet in wb.Worksheets:
                try:
                    for table in sheet.ListObjects:
                        try:
                            tables.append(self._get_table_info(table, sheet))
                        except Exception:
                            # Skip tables that can't be read
                            continue
                except Exception:
                    # Skip sheets that can't be read
                    continue
        else:
            # List tables in specific worksheet
            ws = _find_worksheet(wb, worksheet)
            if ws:
                for table in ws.ListObjects:
                    try:
                        tables.append(self._get_table_info(table, ws))
                    except Exception:
                        # Skip tables that can't be read
                        continue

        return tables
