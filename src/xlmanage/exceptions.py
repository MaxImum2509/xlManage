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
