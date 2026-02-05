"""
Tests for table manager functionality.

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
from unittest.mock import MagicMock

from xlmanage.exceptions import TableNameError, TableRangeError
from xlmanage.table_manager import (
    TABLE_NAME_MAX_LENGTH,
    TABLE_NAME_PATTERN,
    TableInfo,
    _find_table,
    _validate_range,
    _validate_table_name,
)


class TestTableInfo:
    """Tests for TableInfo dataclass."""

    def test_table_info_creation(self):
        """Test creating a TableInfo instance."""
        info = TableInfo(
            name="tbl_Sales",
            worksheet_name="Data",
            range_ref="A1:D100",
            header_row_range="A1:D1",
            rows_count=99,
        )

        assert info.name == "tbl_Sales"
        assert info.worksheet_name == "Data"
        assert info.range_ref == "A1:D100"
        assert info.header_row_range == "A1:D1"
        assert info.rows_count == 99

    def test_table_info_zero_rows(self):
        """Test TableInfo with zero data rows."""
        info = TableInfo(
            name="tbl_Empty",
            worksheet_name="Sheet1",
            range_ref="A1:D1",
            header_row_range="A1:D1",
            rows_count=0,
        )

        assert info.rows_count == 0

    def test_table_info_equality(self):
        """Test TableInfo equality comparison."""
        info1 = TableInfo("tbl_Test", "Sheet1", "A1:B10", "A1:B1", 9)
        info2 = TableInfo("tbl_Test", "Sheet1", "A1:B10", "A1:B1", 9)

        assert info1 == info2

    def test_table_info_inequality(self):
        """Test TableInfo inequality comparison."""
        info1 = TableInfo("tbl_Test1", "Sheet1", "A1:B10", "A1:B1", 9)
        info2 = TableInfo("tbl_Test2", "Sheet1", "A1:B10", "A1:B1", 9)

        assert info1 != info2


class TestTableNameConstants:
    """Tests for table name validation constants."""

    def test_table_name_max_length(self):
        """Test TABLE_NAME_MAX_LENGTH constant."""
        assert TABLE_NAME_MAX_LENGTH == 255

    def test_table_name_pattern(self):
        """Test TABLE_NAME_PATTERN constant."""
        import re

        pattern = re.compile(TABLE_NAME_PATTERN)

        # Valid names
        assert pattern.match("tbl_Sales")
        assert pattern.match("Data_2024")
        assert pattern.match("_PrivateTable")
        assert pattern.match("MyTable123")

        # Invalid names
        assert not pattern.match("1Data")  # Starts with digit
        assert not pattern.match("tbl Sales")  # Contains space
        assert not pattern.match("tbl-Sales")  # Contains hyphen
        assert not pattern.match("tbl.Sales")  # Contains dot


class TestValidateTableName:
    """Tests for _validate_table_name function."""

    def test_validate_valid_names(self):
        """Test _validate_table_name with valid names."""
        valid_names = [
            "tbl_Sales",
            "Data_2024",
            "_PrivateTable",
            "MyTable",
            "T",
            "_",
            "Table123",
            "a" * 255,  # Max length
        ]

        for name in valid_names:
            _validate_table_name(name)  # Should not raise

    def test_validate_empty_name(self):
        """Test _validate_table_name with empty name."""
        with pytest.raises(TableNameError) as exc_info:
            _validate_table_name("")

        assert exc_info.value.name == ""
        assert "cannot be empty" in exc_info.value.reason

    def test_validate_whitespace_only_name(self):
        """Test _validate_table_name with whitespace-only name."""
        with pytest.raises(TableNameError) as exc_info:
            _validate_table_name("   ")

        assert "cannot be empty" in exc_info.value.reason

    def test_validate_name_too_long(self):
        """Test _validate_table_name with name exceeding max length."""
        long_name = "A" * 256

        with pytest.raises(TableNameError) as exc_info:
            _validate_table_name(long_name)

        assert exc_info.value.name == long_name
        assert "exceeds 255 characters" in exc_info.value.reason
        assert "256" in exc_info.value.reason

    def test_validate_name_starts_with_digit(self):
        """Test _validate_table_name with name starting with digit."""
        with pytest.raises(TableNameError) as exc_info:
            _validate_table_name("1Data")

        assert exc_info.value.name == "1Data"
        assert "must start with letter or underscore" in exc_info.value.reason

    def test_validate_name_with_space(self):
        """Test _validate_table_name with name containing space."""
        with pytest.raises(TableNameError) as exc_info:
            _validate_table_name("tbl Sales")

        assert exc_info.value.name == "tbl Sales"
        assert "must start with letter or underscore" in exc_info.value.reason

    def test_validate_name_with_hyphen(self):
        """Test _validate_table_name with name containing hyphen."""
        with pytest.raises(TableNameError) as exc_info:
            _validate_table_name("tbl-Sales")

        assert exc_info.value.name == "tbl-Sales"
        assert "alphanumeric" in exc_info.value.reason

    def test_validate_name_with_dot(self):
        """Test _validate_table_name with name containing dot."""
        with pytest.raises(TableNameError) as exc_info:
            _validate_table_name("tbl.Sales")

        assert "alphanumeric" in exc_info.value.reason

    def test_validate_cell_reference_a1(self):
        """Test _validate_table_name with A1-style cell reference."""
        cell_refs = ["A1", "Z99", "AA100", "XFD1048576"]

        for ref in cell_refs:
            with pytest.raises(TableNameError) as exc_info:
                _validate_table_name(ref)

            assert exc_info.value.name == ref
            assert "cannot be a cell reference" in exc_info.value.reason

    def test_validate_cell_reference_r1c1(self):
        """Test _validate_table_name with R1C1-style cell reference."""
        cell_refs = ["R1C1", "R10C5", "r1c1", "r100c200"]

        for ref in cell_refs:
            with pytest.raises(TableNameError) as exc_info:
                _validate_table_name(ref)

            assert exc_info.value.name == ref
            assert "cannot be a cell reference" in exc_info.value.reason

    def test_validate_name_with_special_characters(self):
        """Test _validate_table_name with various special characters."""
        invalid_names = [
            "tbl@Sales",
            "tbl#Sales",
            "tbl$Sales",
            "tbl%Sales",
            "tbl&Sales",
            "tbl*Sales",
            "tbl(Sales)",
            "tbl[Sales]",
            "tbl{Sales}",
        ]

        for name in invalid_names:
            with pytest.raises(TableNameError):
                _validate_table_name(name)

    def test_validate_max_length_boundary(self):
        """Test _validate_table_name at max length boundary."""
        # Exactly 255 characters - should pass
        name_255 = "A" * 255
        _validate_table_name(name_255)  # Should not raise

        # 256 characters - should fail
        name_256 = "A" * 256
        with pytest.raises(TableNameError) as exc_info:
            _validate_table_name(name_256)

        assert "exceeds 255 characters" in exc_info.value.reason

    def test_validate_underscore_variations(self):
        """Test _validate_table_name with various underscore patterns."""
        valid_names = [
            "_Table",
            "__Table",
            "Table_",
            "Table__",
            "_Table_",
            "My_Table_Name",
        ]

        for name in valid_names:
            _validate_table_name(name)  # Should not raise

    def test_validate_mixed_case(self):
        """Test _validate_table_name with mixed case names."""
        valid_names = [
            "TblSales",
            "tblSales",
            "TBLSALES",
            "MyTableName",
            "my_Table_Name",
        ]

        for name in valid_names:
            _validate_table_name(name)  # Should not raise


class TestFindTable:
    """Tests for _find_table function."""

    def test_find_table_success(self):
        """Test _find_table finds a table by name."""
        # Create mock worksheet with tables
        ws = MagicMock()
        table1 = MagicMock()
        table1.Name = "tbl_Sales"
        table2 = MagicMock()
        table2.Name = "tbl_Products"
        ws.ListObjects = [table1, table2]

        result = _find_table(ws, "tbl_Sales")

        assert result == table1

    def test_find_table_not_found(self):
        """Test _find_table returns None when table doesn't exist."""
        ws = MagicMock()
        table1 = MagicMock()
        table1.Name = "tbl_Sales"
        ws.ListObjects = [table1]

        result = _find_table(ws, "tbl_Missing")

        assert result is None

    def test_find_table_case_sensitive(self):
        """Test _find_table is case-sensitive."""
        ws = MagicMock()
        table1 = MagicMock()
        table1.Name = "tbl_Sales"
        ws.ListObjects = [table1]

        # Different case should not match
        result = _find_table(ws, "TBL_SALES")
        assert result is None

        # Exact case should match
        result = _find_table(ws, "tbl_Sales")
        assert result == table1

    def test_find_table_empty_worksheet(self):
        """Test _find_table with no tables in worksheet."""
        ws = MagicMock()
        ws.ListObjects = []

        result = _find_table(ws, "tbl_Any")

        assert result is None

    def test_find_table_handles_error(self):
        """Test _find_table continues when table can't be read."""
        ws = MagicMock()
        table1 = MagicMock()
        table1.Name = Exception("COM error")
        table2 = MagicMock()
        table2.Name = "tbl_Valid"
        ws.ListObjects = [table1, table2]

        result = _find_table(ws, "tbl_Valid")

        assert result == table2


class TestValidateRange:
    """Tests for _validate_range function."""

    def test_validate_valid_ranges(self):
        """Test _validate_range with valid range references."""
        valid_ranges = [
            "A1:D10",
            "B5:Z100",
            "AA1:ZZ999",
            "A1:A1",  # Single cell as range
            "$A$1:$D$10",  # With $ signs
        ]

        for range_ref in valid_ranges:
            _validate_range(range_ref)  # Should not raise

    def test_validate_range_with_sheet_reference(self):
        """Test _validate_range with sheet name prefix."""
        valid_ranges = [
            "Sheet1!A1:D10",
            "'My Sheet'!B5:Z100",
            "Data!A1:Z999",
        ]

        for range_ref in valid_ranges:
            _validate_range(range_ref)  # Should not raise

    def test_validate_empty_range(self):
        """Test _validate_range with empty range."""
        with pytest.raises(TableRangeError) as exc_info:
            _validate_range("")

        assert exc_info.value.range_ref == ""
        assert "cannot be empty" in exc_info.value.reason

    def test_validate_whitespace_only_range(self):
        """Test _validate_range with whitespace-only range."""
        with pytest.raises(TableRangeError) as exc_info:
            _validate_range("   ")

        assert "cannot be empty" in exc_info.value.reason

    def test_validate_range_missing_colon(self):
        """Test _validate_range with no colon."""
        with pytest.raises(TableRangeError) as exc_info:
            _validate_range("A1")

        assert exc_info.value.range_ref == "A1"
        assert "must have format A1:Z99" in exc_info.value.reason

    def test_validate_range_invalid_syntax(self):
        """Test _validate_range with invalid syntax."""
        invalid_ranges = [
            "A1:D",  # Incomplete end cell
            "1:D10",  # Invalid start cell
            "A:D10",  # Column reference without row
            "A1:10",  # Number only for end cell
        ]

        for range_ref in invalid_ranges:
            with pytest.raises(TableRangeError) as exc_info:
                _validate_range(range_ref)

            assert "invalid range syntax" in exc_info.value.reason

    def test_validate_range_no_colon(self):
        """Test _validate_range with no colon at all."""
        with pytest.raises(TableRangeError) as exc_info:
            _validate_range("ABC")

        assert "must have format A1:Z99" in exc_info.value.reason

    def test_validate_r1c1_range(self):
        """Test _validate_range with R1C1 notation."""
        valid_ranges = [
            "R1C1:R10C5",
            "r1c1:r100c200",
        ]

        for range_ref in valid_ranges:
            _validate_range(range_ref)  # Should not raise

    def test_validate_range_with_dollar_signs(self):
        """Test _validate_range handles $ signs correctly."""
        valid_ranges = [
            "$A$1:$D$10",
            "A$1:D$10",
            "$A1:$D10",
        ]

        for range_ref in valid_ranges:
            _validate_range(range_ref)  # Should not raise
