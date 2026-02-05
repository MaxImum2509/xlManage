"""
Tests for xlmanage exception classes.

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

import pytest
from pathlib import Path

from xlmanage.exceptions import (
    ExcelConnectionError,
    ExcelInstanceNotFoundError,
    ExcelManageError,
    ExcelRPCError,
    TableAlreadyExistsError,
    TableNameError,
    TableNotFoundError,
    TableRangeError,
    WorkbookAlreadyOpenError,
    WorkbookNotFoundError,
    WorkbookSaveError,
    WorksheetAlreadyExistsError,
    WorksheetDeleteError,
    WorksheetNameError,
    WorksheetNotFoundError,
)


class TestExcelManageError:
    """Test base exception class."""

    def test_base_exception(self):
        """Test that ExcelManageError is a proper base exception."""
        with pytest.raises(ExcelManageError):
            raise ExcelManageError("Base error")


class TestExcelConnectionError:
    """Test Excel connection error exception."""

    def test_connection_error_default_message(self):
        """Test ExcelConnectionError with default message."""
        with pytest.raises(ExcelConnectionError) as exc_info:
            raise ExcelConnectionError(hresult=0x80080005)

        error = exc_info.value
        assert error.hresult == 0x80080005
        assert error.message == "Excel connection failed"
        assert "HRESULT: 0x80080005" in str(error)

    def test_connection_error_custom_message(self):
        """Test ExcelConnectionError with custom message."""
        with pytest.raises(ExcelConnectionError) as exc_info:
            raise ExcelConnectionError(
                hresult=0x80080005, message="Excel not installed"
            )

        error = exc_info.value
        assert error.hresult == 0x80080005
        assert error.message == "Excel not installed"
        assert "Excel not installed" in str(error)
        assert "HRESULT: 0x80080005" in str(error)

    def test_connection_error_inheritance(self):
        """Test that ExcelConnectionError inherits from ExcelManageError."""
        with pytest.raises(ExcelManageError):
            raise ExcelConnectionError(hresult=0x80080005)


class TestExcelInstanceNotFoundError:
    """Test Excel instance not found error exception."""

    def test_instance_not_found_default_message(self):
        """Test ExcelInstanceNotFoundError with default message."""
        with pytest.raises(ExcelInstanceNotFoundError) as exc_info:
            raise ExcelInstanceNotFoundError(instance_id="Excel.Application.1")

        error = exc_info.value
        assert error.instance_id == "Excel.Application.1"
        assert error.message == "Instance not found"
        assert "Instance not found: Excel.Application.1" in str(error)

    def test_instance_not_found_custom_message(self):
        """Test ExcelInstanceNotFoundError with custom message."""
        with pytest.raises(ExcelInstanceNotFoundError) as exc_info:
            raise ExcelInstanceNotFoundError(
                instance_id="Excel.Application.1",
                message="Requested Excel instance not available",
            )

        error = exc_info.value
        assert error.instance_id == "Excel.Application.1"
        assert error.message == "Requested Excel instance not available"
        assert "Requested Excel instance not available: Excel.Application.1" in str(
            error
        )

    def test_instance_not_found_inheritance(self):
        """Test that ExcelInstanceNotFoundError inherits from ExcelManageError."""
        with pytest.raises(ExcelManageError):
            raise ExcelInstanceNotFoundError(instance_id="Excel.Application.1")


class TestExcelRPCError:
    """Test Excel RPC error exception."""

    def test_rpc_error_default_message(self):
        """Test ExcelRPCError with default message."""
        with pytest.raises(ExcelRPCError) as exc_info:
            raise ExcelRPCError(hresult=0x800706BE)

        error = exc_info.value
        assert error.hresult == 0x800706BE
        assert error.message == "RPC error"
        assert "HRESULT: 0x800706be" in str(error)

    def test_rpc_error_custom_message(self):
        """Test ExcelRPCError with custom message."""
        with pytest.raises(ExcelRPCError) as exc_info:
            raise ExcelRPCError(hresult=0x800706BE, message="RPC server unavailable")

        error = exc_info.value
        assert error.hresult == 0x800706BE
        assert error.message == "RPC server unavailable"
        assert "RPC server unavailable" in str(error)
        assert "HRESULT: 0x800706be" in str(error)

    def test_rpc_error_inheritance(self):
        """Test that ExcelRPCError inherits from ExcelManageError."""
        with pytest.raises(ExcelManageError):
            raise ExcelRPCError(hresult=0x800706BE)


class TestExceptionAttributes:
    """Test exception attributes and string representations."""

    def test_connection_error_attributes(self):
        """Test ExcelConnectionError attributes."""
        error = ExcelConnectionError(
            hresult=0x80080005, message="Test connection error"
        )
        assert hasattr(error, "hresult")
        assert hasattr(error, "message")
        assert error.hresult == 0x80080005
        assert error.message == "Test connection error"
        assert "HRESULT: 0x80080005" in str(error)

    def test_instance_not_found_attributes(self):
        """Test ExcelInstanceNotFoundError attributes."""
        error = ExcelInstanceNotFoundError(
            instance_id="test_instance", message="Test instance error"
        )
        assert hasattr(error, "instance_id")
        assert hasattr(error, "message")
        assert error.instance_id == "test_instance"
        assert error.message == "Test instance error"
        assert "Test instance error: test_instance" in str(error)

    def test_rpc_error_attributes(self):
        """Test ExcelRPCError attributes."""
        error = ExcelRPCError(hresult=0x800706BE, message="Test RPC error")
        assert hasattr(error, "hresult")
        assert hasattr(error, "message")
        assert error.hresult == 0x800706BE
        assert error.message == "Test RPC error"
        assert "HRESULT: 0x800706be" in str(error)


class TestExceptionEdgeCases:
    """Test exception edge cases and special scenarios."""

    def test_connection_error_zero_hresult(self):
        """Test ExcelConnectionError with HRESULT=0."""
        error = ExcelConnectionError(hresult=0, message="Success but unexpected")
        assert error.hresult == 0
        assert "HRESULT: 0x00000000" in str(error)

    def test_connection_error_negative_hresult(self):
        """Test ExcelConnectionError with negative HRESULT."""
        error = ExcelConnectionError(hresult=-2147467259, message="Negative HRESULT")
        assert error.hresult == -2147467259
        assert "HRESULT: -0x7fffbffb" in str(error)

    def test_instance_not_found_empty_instance_id(self):
        """Test ExcelInstanceNotFoundError with empty instance ID."""
        error = ExcelInstanceNotFoundError(instance_id="", message="Empty ID")
        assert error.instance_id == ""
        assert "Empty ID: " in str(error)

    def test_rpc_error_common_hresults(self):
        """Test ExcelRPCError with common HRESULT values."""
        common_hresults = [
            (0x800706BE, "RPC server unavailable"),
            (0x80010108, "COM object disconnected"),
            (0x80080005, "Server execution failed"),
        ]

        for hresult, message in common_hresults:
            error = ExcelRPCError(hresult=hresult, message=message)
            assert error.hresult == hresult
            assert message in str(error)


class TestExceptionInheritance:
    """Test exception inheritance and polymorphism."""

    def test_all_exceptions_inherit_from_base(self):
        """Test that all exceptions inherit from ExcelManageError."""
        exceptions = [
            ExcelConnectionError(hresult=0x80080005),
            ExcelInstanceNotFoundError(instance_id="test"),
            ExcelRPCError(hresult=0x800706BE),
        ]

        for exc in exceptions:
            assert isinstance(exc, ExcelManageError)
            assert isinstance(exc, Exception)

    def test_catch_base_exception(self):
        """Test catching base exception catches all derived exceptions."""
        exceptions_to_test = [
            ExcelConnectionError(hresult=0x80080005, message="Connection failed"),
            ExcelInstanceNotFoundError(instance_id="test", message="Not found"),
            ExcelRPCError(hresult=0x800706BE, message="RPC error"),
        ]

        for exc in exceptions_to_test:
            with pytest.raises(ExcelManageError):
                raise exc

    def test_exception_hierarchy(self):
        """Test the complete exception hierarchy."""
        # Test that we can catch specific exceptions
        with pytest.raises(ExcelConnectionError):
            raise ExcelConnectionError(hresult=0x80080005)

        with pytest.raises(ExcelInstanceNotFoundError):
            raise ExcelInstanceNotFoundError(instance_id="test")

        with pytest.raises(ExcelRPCError):
            raise ExcelRPCError(hresult=0x800706BE)

        # Test that we can catch the base exception
        with pytest.raises(ExcelManageError):
            raise ExcelConnectionError(hresult=0x80080005)


class TestExceptionEquality:
    """Test exception equality and comparison."""

    def test_exceptions_with_same_values_are_equal(self):
        """Test that exceptions with same values are equal."""
        error1 = ExcelConnectionError(hresult=0x80080005, message="Test")
        error2 = ExcelConnectionError(hresult=0x80080005, message="Test")

        # Note: Exceptions are not typically compared for equality,
        # but we test that their attributes are equal
        assert error1.hresult == error2.hresult
        assert error1.message == error2.message
        assert str(error1) == str(error2)

    def test_exceptions_with_different_values_are_not_equal(self):
        """Test that exceptions with different values are not equal."""
        error1 = ExcelConnectionError(hresult=0x80080005, message="Test1")
        error2 = ExcelConnectionError(hresult=0x800706BE, message="Test2")

        assert error1.hresult != error2.hresult
        assert error1.message != error2.message
        assert str(error1) != str(error2)


class TestExceptionSerialization:
    """Test exception serialization and representation."""

    def test_exception_string_representation(self):
        """Test string representation of exceptions."""
        # Test ExcelConnectionError
        error1 = ExcelConnectionError(hresult=0x80080005, message="Connection failed")
        assert "Connection failed" in str(error1)
        assert "0x80080005" in str(error1)

        # Test ExcelInstanceNotFoundError
        error2 = ExcelInstanceNotFoundError(
            instance_id="Excel.App.1", message="Not found"
        )
        assert "Not found" in str(error2)
        assert "Excel.App.1" in str(error2)

        # Test ExcelRPCError
        error3 = ExcelRPCError(hresult=0x800706BE, message="RPC error")
        assert "RPC error" in str(error3)
        assert "0x800706be" in str(error3)

    def test_exception_repr(self):
        """Test repr() of exceptions."""
        error = ExcelConnectionError(hresult=0x80080005, message="Test")
        repr_str = repr(error)
        assert "ExcelConnectionError" in repr_str
        assert "0x80080005" in repr_str or "Test" in repr_str


class TestExceptionInstantiation:
    """Test exception instantiation and attribute access."""

    def test_instance_not_found_error_instantiation(self):
        """Test ExcelInstanceNotFoundError instantiation with attributes."""
        # Test with default message
        error1 = ExcelInstanceNotFoundError(instance_id="test_instance_1")
        assert error1.instance_id == "test_instance_1"
        assert error1.message == "Instance not found"
        assert "Instance not found: test_instance_1" in str(error1)

        # Test with custom message
        error2 = ExcelInstanceNotFoundError(
            instance_id="test_instance_2", message="Custom message"
        )
        assert error2.instance_id == "test_instance_2"
        assert error2.message == "Custom message"
        assert "Custom message: test_instance_2" in str(error2)

        # Test with empty instance_id
        error3 = ExcelInstanceNotFoundError(instance_id="")
        assert error3.instance_id == ""
        assert "Instance not found: " in str(error3)

    def test_rpc_error_instantiation(self):
        """Test ExcelRPCError instantiation with attributes."""
        # Test with default message
        error1 = ExcelRPCError(hresult=0x800706BE)
        assert error1.hresult == 0x800706BE
        assert error1.message == "RPC error"
        assert "RPC error" in str(error1)

        # Test with custom message
        error2 = ExcelRPCError(hresult=0x80010108, message="Custom RPC message")
        assert error2.hresult == 0x80010108
        assert error2.message == "Custom RPC message"
        assert "Custom RPC message" in str(error2)


class TestWorkbookNotFoundError:
    """Tests for WorkbookNotFoundError."""

    def test_workbook_not_found_default_message(self):
        """Test WorkbookNotFoundError with default message."""
        path = Path("C:/test/missing.xlsx")
        error = WorkbookNotFoundError(path)

        assert error.path == path
        assert error.message == "Workbook not found"
        assert "Workbook not found" in str(error)
        assert "missing.xlsx" in str(error)

    def test_workbook_not_found_custom_message(self):
        """Test WorkbookNotFoundError with custom message."""
        path = Path("D:/data/file.xlsm")
        error = WorkbookNotFoundError(path, "File does not exist")

        assert error.path == path
        assert error.message == "File does not exist"
        assert "File does not exist" in str(error)

    def test_workbook_not_found_inheritance(self):
        """Test WorkbookNotFoundError inherits from ExcelManageError."""
        error = WorkbookNotFoundError(Path("test.xlsx"))
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestWorkbookAlreadyOpenError:
    """Tests for WorkbookAlreadyOpenError."""

    def test_workbook_already_open_default_message(self):
        """Test WorkbookAlreadyOpenError with default message."""
        path = Path("C:/test/open.xlsx")
        name = "open.xlsx"
        error = WorkbookAlreadyOpenError(path, name)

        assert error.path == path
        assert error.name == name
        assert error.message == "Workbook already open"
        assert name in str(error)

    def test_workbook_already_open_custom_message(self):
        """Test WorkbookAlreadyOpenError with custom message."""
        path = Path("D:/data/file.xlsm")
        name = "file.xlsm"
        error = WorkbookAlreadyOpenError(path, name, "Already loaded")

        assert error.message == "Already loaded"
        assert "Already loaded" in str(error)

    def test_workbook_already_open_inheritance(self):
        """Test WorkbookAlreadyOpenError inherits from ExcelManageError."""
        error = WorkbookAlreadyOpenError(Path("test.xlsx"), "test.xlsx")
        assert isinstance(error, ExcelManageError)


class TestWorkbookSaveError:
    """Tests for WorkbookSaveError."""

    def test_workbook_save_error_without_hresult(self):
        """Test WorkbookSaveError without COM error."""
        path = Path("C:/test/readonly.xlsx")
        error = WorkbookSaveError(path)

        assert error.path == path
        assert error.hresult == 0
        assert error.message == "Save failed"
        assert "readonly.xlsx" in str(error)
        assert "HRESULT" not in str(error)  # No HRESULT when 0

    def test_workbook_save_error_with_hresult(self):
        """Test WorkbookSaveError with COM error."""
        path = Path("D:/test/file.xlsx")
        error = WorkbookSaveError(path, hresult=0x80070005)

        assert error.hresult == 0x80070005
        assert "0x80070005" in str(error)
        assert "HRESULT" in str(error)

    def test_workbook_save_error_custom_message(self):
        """Test WorkbookSaveError with custom message."""
        path = Path("E:/data/protected.xlsx")
        error = WorkbookSaveError(path, message="Access denied")

        assert error.message == "Access denied"
        assert "Access denied" in str(error)

    def test_workbook_save_error_inheritance(self):
        """Test WorkbookSaveError inherits from ExcelManageError."""
        error = WorkbookSaveError(Path("test.xlsx"))
        assert isinstance(error, ExcelManageError)


class TestWorksheetNotFoundError:
    """Tests for WorksheetNotFoundError."""

    def test_worksheet_not_found_default_message(self):
        """Test WorksheetNotFoundError with default message."""
        error = WorksheetNotFoundError("Sheet1", "workbook.xlsx")

        assert error.name == "Sheet1"
        assert error.workbook_name == "workbook.xlsx"
        assert "not found" in str(error).lower()
        assert "Sheet1" in str(error)
        assert "workbook.xlsx" in str(error)

    def test_worksheet_not_found_custom_workbook_name(self):
        """Test WorksheetNotFoundError with different workbook name."""
        error = WorksheetNotFoundError("Data", "reports.xlsx")

        assert error.name == "Data"
        assert error.workbook_name == "reports.xlsx"
        assert "not found" in str(error).lower()

    def test_worksheet_not_found_empty_sheet_name(self):
        """Test WorksheetNotFoundError with empty sheet name."""
        error = WorksheetNotFoundError("", "test.xlsx")

        assert error.name == ""
        assert error.workbook_name == "test.xlsx"
        assert "Worksheet ''" in str(error)

    def test_worksheet_not_found_inheritance(self):
        """Test WorksheetNotFoundError inherits from ExcelManageError."""
        error = WorksheetNotFoundError("Sheet1", "workbook.xlsx")
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestWorksheetAlreadyExistsError:
    """Tests for WorksheetAlreadyExistsError."""

    def test_worksheet_already_exists_default_message(self):
        """Test WorksheetAlreadyExistsError with default message."""
        error = WorksheetAlreadyExistsError("Summary", "report.xlsx")

        assert error.name == "Summary"
        assert error.workbook_name == "report.xlsx"
        assert "already exists" in str(error).lower()
        assert "Summary" in str(error)
        assert "report.xlsx" in str(error)

    def test_worksheet_already_exists_with_spaces(self):
        """Test WorksheetAlreadyExistsError with sheet name containing spaces."""
        error = WorksheetAlreadyExistsError("My Data", "workbook.xlsx")

        assert error.name == "My Data"
        assert error.workbook_name == "workbook.xlsx"
        assert "already exists" in str(error).lower()

    def test_worksheet_already_exists_inheritance(self):
        """Test WorksheetAlreadyExistsError inherits from ExcelManageError."""
        error = WorksheetAlreadyExistsError("Sheet1", "workbook.xlsx")
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestWorksheetDeleteError:
    """Tests for WorksheetDeleteError."""

    def test_worksheet_delete_error_default_reason(self):
        """Test WorksheetDeleteError with default reason."""
        error = WorksheetDeleteError("Sheet1", "Cannot delete")

        assert error.name == "Sheet1"
        assert error.reason == "Cannot delete"
        assert "Cannot delete" in str(error)
        assert "Sheet1" in str(error)

    def test_worksheet_delete_last_visible_sheet(self):
        """Test WorksheetDeleteError for last visible sheet."""
        error = WorksheetDeleteError("LastSheet", "last visible sheet")

        assert error.name == "LastSheet"
        assert error.reason == "last visible sheet"
        assert "last visible sheet" in str(error)

    def test_worksheet_delete_error_inheritance(self):
        """Test WorksheetDeleteError inherits from ExcelManageError."""
        error = WorksheetDeleteError("Sheet1", "protected")
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestWorksheetNameError:
    """Tests for WorksheetNameError."""

    def test_worksheet_name_error_default_reason(self):
        """Test WorksheetNameError with default reason."""
        error = WorksheetNameError("Sheet?", "invalid character")

        assert error.name == "Sheet?"
        assert error.reason == "invalid character"
        assert "Invalid worksheet name" in str(error)
        assert "Sheet?" in str(error)

    def test_worksheet_name_error_too_long(self):
        """Test WorksheetNameError for name too long."""
        error = WorksheetNameError("A" * 32, "name exceeds 31 characters")

        assert error.name == "A" * 32
        assert error.reason == "name exceeds 31 characters"
        assert "31 characters" in str(error)

    def test_worksheet_name_error_invalid_character(self):
        """Test WorksheetNameError for invalid character."""
        error = WorksheetNameError("Data/Sheet", "contains invalid character '/'")

        assert error.name == "Data/Sheet"
        assert error.reason == "contains invalid character '/'"
        assert "Invalid worksheet name" in str(error)

    def test_worksheet_name_error_inheritance(self):
        """Test WorksheetNameError inherits from ExcelManageError."""
        error = WorksheetNameError("Sheet!", "invalid character")
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestTableNotFoundError:
    """Tests for TableNotFoundError."""

    def test_table_not_found_default_message(self):
        """Test TableNotFoundError with default message."""
        error = TableNotFoundError("tbl_Sales", "Data")

        assert error.name == "tbl_Sales"
        assert error.worksheet_name == "Data"
        assert "not found" in str(error).lower()
        assert "tbl_Sales" in str(error)
        assert "Data" in str(error)

    def test_table_not_found_different_names(self):
        """Test TableNotFoundError with different table and worksheet names."""
        error = TableNotFoundError("tbl_Customers", "Sheet1")

        assert error.name == "tbl_Customers"
        assert error.worksheet_name == "Sheet1"
        assert "not found" in str(error).lower()

    def test_table_not_found_empty_table_name(self):
        """Test TableNotFoundError with empty table name."""
        error = TableNotFoundError("", "Data")

        assert error.name == ""
        assert error.worksheet_name == "Data"
        assert "Table ''" in str(error)

    def test_table_not_found_inheritance(self):
        """Test TableNotFoundError inherits from ExcelManageError."""
        error = TableNotFoundError("tbl_Test", "Sheet1")
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestTableAlreadyExistsError:
    """Tests for TableAlreadyExistsError."""

    def test_table_already_exists_default_message(self):
        """Test TableAlreadyExistsError with default message."""
        error = TableAlreadyExistsError("tbl_Sales", "workbook.xlsx")

        assert error.name == "tbl_Sales"
        assert error.workbook_name == "workbook.xlsx"
        assert "already exists" in str(error).lower()
        assert "tbl_Sales" in str(error)
        assert "workbook.xlsx" in str(error)

    def test_table_already_exists_different_names(self):
        """Test TableAlreadyExistsError with different names."""
        error = TableAlreadyExistsError("tbl_Data", "report.xlsx")

        assert error.name == "tbl_Data"
        assert error.workbook_name == "report.xlsx"
        assert "already exists" in str(error).lower()

    def test_table_already_exists_inheritance(self):
        """Test TableAlreadyExistsError inherits from ExcelManageError."""
        error = TableAlreadyExistsError("tbl_Test", "test.xlsx")
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestTableRangeError:
    """Tests for TableRangeError."""

    def test_table_range_error_default_reason(self):
        """Test TableRangeError with default reason."""
        error = TableRangeError("A1:D", "invalid syntax")

        assert error.range_ref == "A1:D"
        assert error.reason == "invalid syntax"
        assert "Invalid table range" in str(error)
        assert "A1:D" in str(error)
        assert "invalid syntax" in str(error)

    def test_table_range_error_empty_range(self):
        """Test TableRangeError for empty range."""
        error = TableRangeError("", "range cannot be empty")

        assert error.range_ref == ""
        assert error.reason == "range cannot be empty"
        assert "range cannot be empty" in str(error)

    def test_table_range_error_overlapping(self):
        """Test TableRangeError for overlapping range."""
        error = TableRangeError("A1:D10", "overlaps with existing table")

        assert error.range_ref == "A1:D10"
        assert error.reason == "overlaps with existing table"
        assert "overlaps" in str(error)

    def test_table_range_error_inheritance(self):
        """Test TableRangeError inherits from ExcelManageError."""
        error = TableRangeError("A1:Z", "invalid")
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestTableNameError:
    """Tests for TableNameError."""

    def test_table_name_error_default_reason(self):
        """Test TableNameError with default reason."""
        error = TableNameError("tbl Sales", "contains spaces")

        assert error.name == "tbl Sales"
        assert error.reason == "contains spaces"
        assert "Invalid table name" in str(error)
        assert "tbl Sales" in str(error)
        assert "contains spaces" in str(error)

    def test_table_name_error_too_long(self):
        """Test TableNameError for name too long."""
        long_name = "A" * 256
        error = TableNameError(long_name, "name exceeds 255 characters")

        assert error.name == long_name
        assert error.reason == "name exceeds 255 characters"
        assert "255 characters" in str(error)

    def test_table_name_error_starts_with_digit(self):
        """Test TableNameError for name starting with digit."""
        error = TableNameError("1Data", "must start with letter or underscore")

        assert error.name == "1Data"
        assert error.reason == "must start with letter or underscore"
        assert "Invalid table name" in str(error)

    def test_table_name_error_cell_reference(self):
        """Test TableNameError for name being a cell reference."""
        error = TableNameError("A1", "cannot be a cell reference")

        assert error.name == "A1"
        assert error.reason == "cannot be a cell reference"
        assert "cell reference" in str(error)

    def test_table_name_error_inheritance(self):
        """Test TableNameError inherits from ExcelManageError."""
        error = TableNameError("tbl!", "invalid character")
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)
