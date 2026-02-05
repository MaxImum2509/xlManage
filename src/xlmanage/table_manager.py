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

from .exceptions import TableNameError

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
