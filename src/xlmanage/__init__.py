"""
xlmanage package initialization - Exports publics et version.

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
    "VBAManager",
    "VBAModuleInfo",
    "MacroRunner",
    "MacroResult",
    "ExcelOptimizer",
    "ScreenOptimizer",
    "CalculationOptimizer",
    "OptimizationState",
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

# Import main classes
from .calculation_optimizer import CalculationOptimizer
from .excel_manager import ExcelManager, InstanceInfo
from .excel_optimizer import ExcelOptimizer, OptimizationState
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
from .macro_runner import MacroResult, MacroRunner
from .screen_optimizer import ScreenOptimizer
from .table_manager import TableInfo, TableManager
from .vba_manager import VBAManager, VBAModuleInfo
from .workbook_manager import WorkbookInfo, WorkbookManager
from .worksheet_manager import WorksheetInfo, WorksheetManager
