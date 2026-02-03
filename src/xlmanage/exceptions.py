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
