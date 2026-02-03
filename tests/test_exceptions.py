"""
Test exceptions for xlmanage COM error handling.

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
from xlmanage.exceptions import (
    ExcelManageError,
    ExcelConnectionError,
    ExcelInstanceNotFoundError,
    ExcelRPCError,
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
