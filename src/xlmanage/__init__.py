"""
xlmanage package initialization.
"""

__version__ = "0.1.0"
__all__ = [
    "cli",
    "excel_manager",
    "exceptions",
    "workbook_manager",
    "worksheet_manager",
    "ExcelManager",
    "InstanceInfo",
    "WorkbookManager",
    "WorkbookInfo",
    "WorksheetInfo",
    "ExcelConnectionError",
    "ExcelInstanceNotFoundError",
    "ExcelManageError",
    "ExcelRPCError",
    "WorkbookNotFoundError",
    "WorkbookAlreadyOpenError",
    "WorkbookSaveError",
    "WorksheetAlreadyExistsError",
    "WorksheetDeleteError",
    "WorksheetNameError",
    "WorksheetNotFoundError",
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
    WorksheetAlreadyExistsError,
    WorksheetDeleteError,
    WorksheetNameError,
    WorksheetNotFoundError,
)
from .workbook_manager import WorkbookInfo as WorkbookInfoClass
from .workbook_manager import WorkbookManager
from .worksheet_manager import WorksheetInfo as WorksheetInfoData

# Export both for clarity
# Note: Use WorkbookInfoClass for workbook info, WorksheetInfoData for worksheet info
WorkbookInfo = WorkbookInfoClass
WorksheetInfo = WorksheetInfoData
