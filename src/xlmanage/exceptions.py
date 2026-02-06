"""
Exception classes for xlmanage COM error handling.

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

from pathlib import Path


class ExcelManageError(Exception):
    """Base class for all xlmanage exceptions."""

    pass


class ExcelConnectionError(ExcelManageError):
    """Excel COM connection failed.

    Raised when Excel is not installed or COM server is unavailable.
    """

    def __init__(self, hresult: int, message: str = "Excel connection failed"):
        """Initialize Excel connection error.

        Args:
            hresult: COM HRESULT error code (e.g., 0x80080005)
            message: Human-readable error message
        """
        self.hresult = hresult
        self.message = message
        super().__init__(f"{message} (HRESULT: {hresult:#010x})")


class ExcelInstanceNotFoundError(ExcelManageError):
    """Excel instance not found.

    Raised when a requested Excel instance cannot be found.
    """

    def __init__(self, instance_id: str, message: str = "Instance not found"):
        """Initialize instance not found error.

        Args:
            instance_id: Identifier of the instance that was not found
            message: Human-readable error message
        """
        self.instance_id = instance_id
        self.message = message
        super().__init__(f"{message}: {instance_id}")


class ExcelRPCError(ExcelManageError):
    """Excel RPC error.

    Raised when COM server is disconnected or unavailable.
    """

    def __init__(self, hresult: int, message: str = "RPC error"):
        """Initialize RPC error.

        Args:
            hresult: COM HRESULT error code (e.g., 0x800706BE, 0x80010108)
            message: Human-readable error message
        """
        self.hresult = hresult
        self.message = message
        super().__init__(f"{message} (HRESULT: {hresult:#010x})")


class WorkbookNotFoundError(ExcelManageError):
    """Classeur introuvable sur le disque.

    Raised when attempting to open a workbook file that doesn't exist.
    """

    def __init__(self, path: Path, message: str = "Workbook not found"):
        """Initialize workbook not found error.

        Args:
            path: Path to the workbook file that was not found
            message: Human-readable error message
        """
        self.path = path
        self.message = message
        super().__init__(f"{message}: {path}")


class WorkbookAlreadyOpenError(ExcelManageError):
    """Classeur déjà ouvert dans l'instance Excel.

    Raised when attempting to open a workbook that is already open.
    """

    def __init__(self, path: Path, name: str, message: str = "Workbook already open"):
        """Initialize workbook already open error.

        Args:
            path: Path to the workbook file
            name: Name of the workbook (e.g., "data.xlsx")
            message: Human-readable error message
        """
        self.path = path
        self.name = name
        self.message = message
        super().__init__(f"{message}: {name} at {path}")


class WorkbookSaveError(ExcelManageError):
    """Échec de sauvegarde du classeur.

    Raised when save operation fails due to permissions, invalid path, or format issues.
    """

    def __init__(self, path: Path, hresult: int = 0, message: str = "Save failed"):
        """Initialize workbook save error.

        Args:
            path: Path where the save was attempted
            hresult: COM HRESULT error code (0 if not a COM error)
            message: Human-readable error message
        """
        self.path = path
        self.hresult = hresult
        self.message = message

        if hresult != 0:
            super().__init__(f"{message}: {path} (HRESULT: {hresult:#010x})")
        else:
            super().__init__(f"{message}: {path}")


class WorksheetNotFoundError(ExcelManageError):
    """Feuille introuvable dans le classeur.

    Raised when attempting to access a worksheet that doesn't exist.
    """

    def __init__(self, name: str, workbook_name: str):
        """Initialize worksheet not found error.

        Args:
            name: Name of the worksheet that was not found
            workbook_name: Name of the workbook that was searched
        """
        self.name = name
        self.workbook_name = workbook_name
        super().__init__(f"Worksheet '{name}' not found in workbook '{workbook_name}'")


class WorksheetAlreadyExistsError(ExcelManageError):
    """Nom de feuille déjà utilisé.

    Raised when attempting to create a worksheet with a name that already exists.
    """

    def __init__(self, name: str, workbook_name: str):
        """Initialize worksheet already exists error.

        Args:
            name: Name of the worksheet that already exists
            workbook_name: Name of the workbook
        """
        self.name = name
        self.workbook_name = workbook_name
        super().__init__(
            f"Worksheet '{name}' already exists in workbook '{workbook_name}'"
        )


class WorksheetDeleteError(ExcelManageError):
    """Suppression de feuille impossible.

    Raised when a worksheet cannot be deleted (e.g., last visible sheet).
    """

    def __init__(self, name: str, reason: str):
        """Initialize worksheet delete error.

        Args:
            name: Name of the worksheet that cannot be deleted
            reason: Explanation of why deletion failed
        """
        self.name = name
        self.reason = reason
        super().__init__(f"Cannot delete worksheet '{name}': {reason}")


class WorksheetNameError(ExcelManageError):
    """Nom de feuille invalide.

    Raised when a worksheet name violates Excel naming rules.
    """

    def __init__(self, name: str, reason: str):
        """Initialize worksheet name error.

        Args:
            name: The invalid worksheet name
            reason: Explanation of why name is invalid
        """
        self.name = name
        self.reason = reason
        super().__init__(f"Invalid worksheet name '{name}': {reason}")


class TableNotFoundError(ExcelManageError):
    """Table introuvable dans la feuille.

    Raised when attempting to access a table that doesn't exist.
    """

    def __init__(self, name: str, worksheet_name: str):
        """Initialize table not found error.

        Args:
            name: Name of the table that was not found
            worksheet_name: Name of the worksheet that was searched
        """
        self.name = name
        self.worksheet_name = worksheet_name
        super().__init__(f"Table '{name}' not found in worksheet '{worksheet_name}'")


class TableAlreadyExistsError(ExcelManageError):
    """Nom de table déjà utilisé.

    Raised when attempting to create a table with a name that already exists.
    """

    def __init__(self, name: str, workbook_name: str):
        """Initialize table already exists error.

        Args:
            name: Name of the table that already exists
            workbook_name: Name of the workbook (tables are unique per workbook)
        """
        self.name = name
        self.workbook_name = workbook_name
        super().__init__(f"Table '{name}' already exists in workbook '{workbook_name}'")


class TableRangeError(ExcelManageError):
    """Plage de table invalide.

    Raised when a table range is invalid (syntax error, empty, overlaps, etc.).
    """

    def __init__(self, range_ref: str, reason: str):
        """Initialize table range error.

        Args:
            range_ref: The invalid range reference (e.g., "A1:D10")
            reason: Explanation of why the range is invalid
        """
        self.range_ref = range_ref
        self.reason = reason
        super().__init__(f"Invalid table range '{range_ref}': {reason}")


class TableNameError(ExcelManageError):
    """Nom de table invalide.

    Raised when a table name violates Excel naming rules.
    """

    def __init__(self, name: str, reason: str):
        """Initialize table name error.

        Args:
            name: The invalid table name
            reason: Explanation of why the name is invalid
        """
        self.name = name
        self.reason = reason
        super().__init__(f"Invalid table name '{name}': {reason}")


class VBAProjectAccessError(ExcelManageError):
    """Accès au projet VBA refusé par le Trust Center.

    Raised when Excel's Trust Center blocks programmatic access to VBA.
    """

    def __init__(self, workbook_name: str):
        """Initialize VBA project access error.

        Args:
            workbook_name: Name of the workbook with blocked VBA access
        """
        self.workbook_name = workbook_name
        super().__init__(
            f"Access to VBA project in '{workbook_name}' denied. "
            "Enable 'Trust access to the VBA project object model' in "
            "Excel Trust Center."
        )


class VBAModuleNotFoundError(ExcelManageError):
    """Module VBA introuvable dans le projet.

    Raised when attempting to access a VBA module that doesn't exist,
    or when trying to delete a non-deletable module (document modules).
    """

    def __init__(self, module_name: str, workbook_name: str, reason: str = ""):
        """Initialize VBA module not found error.

        Args:
            module_name: Name of the missing module
            workbook_name: Name of the workbook that was searched
            reason: Optional additional context (e.g., "Cannot delete document module")
        """
        self.module_name = module_name
        self.workbook_name = workbook_name
        self.reason = reason

        if reason:
            message = f"Module '{module_name}' in '{workbook_name}': {reason}"
        else:
            message = (
                f"VBA module '{module_name}' not found in workbook '{workbook_name}'"
            )

        super().__init__(message)


class VBAModuleAlreadyExistsError(ExcelManageError):
    """Module VBA avec ce nom existe déjà.

    Raised when attempting to import a module with a duplicate name.
    """

    def __init__(self, module_name: str, workbook_name: str):
        """Initialize VBA module already exists error.

        Args:
            module_name: Name of the duplicate module
            workbook_name: Name of the workbook containing the duplicate
        """
        self.module_name = module_name
        self.workbook_name = workbook_name
        super().__init__(
            f"VBA module '{module_name}' already exists in workbook '{workbook_name}'"
        )


class VBAImportError(ExcelManageError):
    """Échec d'import de module VBA.

    Raised when importing a VBA module fails (invalid file, wrong encoding, etc.).
    """

    def __init__(self, module_file: str, reason: str):
        """Initialize VBA import error.

        Args:
            module_file: Path to the module file that failed to import
            reason: Explanation of why the import failed
        """
        self.module_file = module_file
        self.reason = reason
        super().__init__(f"Failed to import VBA module from '{module_file}': {reason}")


class VBAExportError(ExcelManageError):
    """Échec d'export de module VBA.

    Raised when exporting a VBA module fails (permissions, invalid path, etc.).
    """

    def __init__(self, module_name: str, output_path: str, reason: str):
        """Initialize VBA export error.

        Args:
            module_name: Name of the module that failed to export
            output_path: Destination path where export was attempted
            reason: Explanation of why the export failed
        """
        self.module_name = module_name
        self.output_path = output_path
        self.reason = reason
        super().__init__(
            f"Failed to export VBA module '{module_name}' to '{output_path}': {reason}"
        )


class VBAMacroError(ExcelManageError):
    """Échec d'exécution ou de parsing de macro VBA.

    Raised when a VBA macro execution fails, macro is not found,
    or argument parsing fails.

    Attributes:
        macro_name: Name of the macro (optional for parsing errors)
        reason: Explanation of the failure
    """

    def __init__(self, macro_name: str = "", reason: str = "") -> None:
        """Initialize VBA macro error.

        Args:
            macro_name: Name of the macro that failed (empty for parsing errors)
            reason: Explanation of the failure (from COM excepinfo[2] or parsing error)
        """
        self.macro_name = macro_name
        self.reason = reason

        message = "Macro error"
        if macro_name:
            message += f" '{macro_name}'"
        if reason:
            message += f": {reason}"

        super().__init__(message)


class VBAWorkbookFormatError(ExcelManageError):
    """Classeur au format .xlsx ne supportant pas les macros.

    Raised when attempting VBA operations on a macro-free workbook format.
    """

    def __init__(self, workbook_name: str):
        """Initialize VBA workbook format error.

        Args:
            workbook_name: Name of the .xlsx workbook
        """
        self.workbook_name = workbook_name
        super().__init__(
            f"Workbook '{workbook_name}' is in .xlsx format which doesn't support VBA. "
            "Convert to .xlsm format to use macros."
        )
