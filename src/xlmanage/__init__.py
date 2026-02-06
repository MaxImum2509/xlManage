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
    "table_manager",
    "ExcelManager",
    "InstanceInfo",
    "WorkbookManager",
    "WorkbookInfo",
    "WorksheetManager",
    "WorksheetInfo",
    "TableManager",
    "TableInfo",
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
    "TableNotFoundError",
    "TableAlreadyExistsError",
    "TableRangeError",
    "TableNameError",
    "VBAProjectAccessError",
    "VBAModuleNotFoundError",
    "VBAModuleAlreadyExistsError",
    "VBAImportError",
    "VBAExportError",
    "VBAMacroError",
    "VBAWorkbookFormatError",
]

# Import exceptions for easy access
# Import main classes
from .excel_manager import ExcelManager, InstanceInfo
from .exceptions import (
    ExcelConnectionError,
    ExcelInstanceNotFoundError,
    ExcelManageError,
    ExcelRPCError,
    TableAlreadyExistsError,
    TableNameError,
    TableNotFoundError,
    TableRangeError,
    VBAExportError,
    VBAImportError,
    VBAMacroError,
    VBAModuleAlreadyExistsError,
    VBAModuleNotFoundError,
    VBAProjectAccessError,
    VBAWorkbookFormatError,
    WorkbookAlreadyOpenError,
    WorkbookNotFoundError,
    WorkbookSaveError,
    WorksheetAlreadyExistsError,
    WorksheetDeleteError,
    WorksheetNameError,
    WorksheetNotFoundError,
)
from .table_manager import TableInfo as TableInfoData
from .table_manager import TableManager
from .workbook_manager import WorkbookInfo as WorkbookInfoClass
from .workbook_manager import WorkbookManager
from .worksheet_manager import WorksheetInfo as WorksheetInfoData
from .worksheet_manager import WorksheetManager

# Export both for clarity
# Note: Use WorkbookInfoClass for workbook info, WorksheetInfoData for
# worksheet info, TableInfoData for table info
WorkbookInfo = WorkbookInfoClass
WorksheetInfo = WorksheetInfoData
TableInfo = TableInfoData
