"""
xlmanage package initialization.
"""

__version__ = "0.1.0"
__all__ = [
    "cli",
    "excel_manager",
    "exceptions",
    "ExcelManager",
    "InstanceInfo",
    "ExcelConnectionError",
    "ExcelInstanceNotFoundError",
    "ExcelManageError",
    "ExcelRPCError",
    "WorkbookNotFoundError",
    "WorkbookAlreadyOpenError",
    "WorkbookSaveError",
]

# Import exceptions for easy access
# Import main classes
from .excel_manager import ExcelManager, InstanceInfo
from .exceptions import (
    ExcelConnectionError,
    ExcelInstanceNotFoundError,
    ExcelManageError,
    ExcelRPCError,
    WorkbookAlreadyOpenError,
    WorkbookNotFoundError,
    WorkbookSaveError,
)
