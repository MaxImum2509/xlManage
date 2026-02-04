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
