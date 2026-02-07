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
from unittest.mock import MagicMock, Mock

from xlmanage.exceptions import (
    TableAlreadyExistsError,
    TableNameError,
    TableNotFoundError,
    TableRangeError,
)
from xlmanage.table_manager import (
    TABLE_NAME_MAX_LENGTH,
    TABLE_NAME_PATTERN,
    TableInfo,
    TableManager,
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
            range_address="$A$1:$D$100",
            columns=["Col1", "Col2", "Col3", "Col4"],
            rows_count=99,
            header_row="$A$1:$D$1",
        )

        assert info.name == "tbl_Sales"
        assert info.worksheet_name == "Data"
        assert info.range_address == "$A$1:$D$100"
        assert info.columns == ["Col1", "Col2", "Col3", "Col4"]
        assert info.rows_count == 99
        assert info.header_row == "$A$1:$D$1"

    def test_table_info_zero_rows(self):
        """Test TableInfo with zero data rows."""
        info = TableInfo(
            name="tbl_Empty",
            worksheet_name="Sheet1",
            range_address="$A$1:$D$1",
            columns=["Col1", "Col2", "Col3", "Col4"],
            rows_count=0,
            header_row="$A$1:$D$1",
        )

        assert info.rows_count == 0

    def test_table_info_equality(self):
        """Test TableInfo equality comparison."""
        info1 = TableInfo("tbl_Test", "Sheet1", "$A$1:$B$10", ["A", "B"], 9, "$A$1:$B$1")
        info2 = TableInfo("tbl_Test", "Sheet1", "$A$1:$B$10", ["A", "B"], 9, "$A$1:$B$1")

        assert info1 == info2

    def test_table_info_inequality(self):
        """Test TableInfo inequality comparison."""
        info1 = TableInfo("tbl_Test1", "Sheet1", "$A$1:$B$10", ["A", "B"], 9, "$A$1:$B$1")
        info2 = TableInfo("tbl_Test2", "Sheet1", "$A$1:$B$10", ["A", "B"], 9, "$A$1:$B$1")

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
        """Test _find_table finds a table by name across worksheets."""
        # _find_table(wb, name) takes a workbook, searches all worksheets
        table1 = MagicMock()
        table1.Name = "tbl_Sales"
        table2 = MagicMock()
        table2.Name = "tbl_Products"

        ws = MagicMock()
        ws.ListObjects = [table1, table2]

        wb = MagicMock()
        wb.Worksheets = [ws]

        result = _find_table(wb, "tbl_Sales")

        assert result is not None
        found_ws, found_table = result
        assert found_table == table1

    def test_find_table_not_found(self):
        """Test _find_table returns None when table doesn't exist."""
        table1 = MagicMock()
        table1.Name = "tbl_Sales"

        ws = MagicMock()
        ws.ListObjects = [table1]

        wb = MagicMock()
        wb.Worksheets = [ws]

        result = _find_table(wb, "tbl_Missing")

        assert result is None

    def test_find_table_case_sensitive(self):
        """Test _find_table is case-sensitive."""
        table1 = MagicMock()
        table1.Name = "tbl_Sales"

        ws = MagicMock()
        ws.ListObjects = [table1]

        wb = MagicMock()
        wb.Worksheets = [ws]

        # Different case should not match
        result = _find_table(wb, "TBL_SALES")
        assert result is None

        # Exact case should match
        result = _find_table(wb, "tbl_Sales")
        assert result is not None
        assert result[1] == table1

    def test_find_table_empty_worksheet(self):
        """Test _find_table with no tables in worksheet."""
        ws = MagicMock()
        ws.ListObjects = []

        wb = MagicMock()
        wb.Worksheets = [ws]

        result = _find_table(wb, "tbl_Any")

        assert result is None

    def test_find_table_handles_error(self):
        """Test _find_table continues when table can't be read."""
        table1 = MagicMock()
        type(table1).Name = property(
            lambda self: (_ for _ in ()).throw(Exception("COM error"))
        )
        table2 = MagicMock()
        table2.Name = "tbl_Valid"

        ws = MagicMock()
        ws.ListObjects = [table1, table2]

        wb = MagicMock()
        wb.Worksheets = [ws]

        result = _find_table(wb, "tbl_Valid")

        assert result is not None
        assert result[1] == table2


class TestValidateRange:
    """Tests for _validate_range function."""

    def _make_ws_mock(self):
        """Create a worksheet mock for _validate_range tests."""
        ws = MagicMock()
        ws.Range.return_value = MagicMock()
        ws.ListObjects = []  # No existing tables
        return ws

    def test_validate_valid_ranges(self):
        """Test _validate_range with valid range references."""
        ws = self._make_ws_mock()
        valid_ranges = [
            "A1:D10",
            "B5:Z100",
            "AA1:ZZ999",
            "A1:A1",  # Single cell as range
            "$A$1:$D$10",  # With $ signs
        ]

        for range_ref in valid_ranges:
            _validate_range(ws, range_ref)  # Should not raise

    def test_validate_range_with_sheet_reference(self):
        """Test _validate_range with sheet name prefix."""
        ws = self._make_ws_mock()
        valid_ranges = [
            "Sheet1!A1:D10",
            "'My Sheet'!B5:Z100",
            "Data!A1:Z999",
        ]

        for range_ref in valid_ranges:
            _validate_range(ws, range_ref)  # Should not raise

    def test_validate_empty_range(self):
        """Test _validate_range with empty range."""
        ws = self._make_ws_mock()
        with pytest.raises(TableRangeError) as exc_info:
            _validate_range(ws, "")

        assert exc_info.value.range_ref == ""
        assert "cannot be empty" in exc_info.value.reason

    def test_validate_whitespace_only_range(self):
        """Test _validate_range with whitespace-only range."""
        ws = self._make_ws_mock()
        with pytest.raises(TableRangeError) as exc_info:
            _validate_range(ws, "   ")

        assert "cannot be empty" in exc_info.value.reason

    def test_validate_range_missing_colon(self):
        """Test _validate_range with invalid range syntax."""
        ws = self._make_ws_mock()
        ws.Range.side_effect = Exception("Invalid range")
        with pytest.raises(TableRangeError) as exc_info:
            _validate_range(ws, "A1")

        assert exc_info.value.range_ref == "A1"
        assert "invalid range syntax" in exc_info.value.reason

    def test_validate_range_invalid_syntax(self):
        """Test _validate_range with invalid syntax."""
        ws = self._make_ws_mock()
        ws.Range.side_effect = Exception("Invalid range")
        invalid_ranges = [
            "A1:D",  # Incomplete end cell
            "1:D10",  # Invalid start cell
            "A:D10",  # Column reference without row
            "A1:10",  # Number only for end cell
        ]

        for range_ref in invalid_ranges:
            with pytest.raises(TableRangeError) as exc_info:
                _validate_range(ws, range_ref)

            assert "invalid range syntax" in exc_info.value.reason

    def test_validate_range_no_colon(self):
        """Test _validate_range with no colon at all."""
        ws = self._make_ws_mock()
        ws.Range.side_effect = Exception("Invalid range")
        with pytest.raises(TableRangeError) as exc_info:
            _validate_range(ws, "ABC")

        assert "invalid range syntax" in exc_info.value.reason

    def test_validate_r1c1_range(self):
        """Test _validate_range with R1C1 notation."""
        ws = self._make_ws_mock()
        valid_ranges = [
            "R1C1:R10C5",
            "r1c1:r100c200",
        ]

        for range_ref in valid_ranges:
            _validate_range(ws, range_ref)  # Should not raise

    def test_validate_range_with_dollar_signs(self):
        """Test _validate_range handles $ signs correctly."""
        ws = self._make_ws_mock()
        valid_ranges = [
            "$A$1:$D$10",
            "A$1:D$10",
            "$A1:$D10",
        ]

        for range_ref in valid_ranges:
            _validate_range(ws, range_ref)  # Should not raise


class TestTableManager:
    """Tests for TableManager class."""

    def test_table_manager_initialization(self):
        """Test TableManager initialization."""
        mock_excel_mgr = Mock()
        manager = TableManager(mock_excel_mgr)

        assert manager._mgr == mock_excel_mgr

    def test_get_table_info_with_data(self):
        """Test _get_table_info with table containing data."""
        mock_excel_mgr = Mock()
        manager = TableManager(mock_excel_mgr)

        # Create mock columns
        col1, col2, col3, col4 = Mock(), Mock(), Mock(), Mock()
        col1.Name, col2.Name = "Col1", "Col2"
        col3.Name, col4.Name = "Col3", "Col4"

        # Create mock table
        mock_table = Mock()
        mock_table.Name = "tbl_Sales"
        mock_table.Range.Address = "$A$1:$D$100"
        mock_table.HeaderRowRange.Address = "$A$1:$D$1"
        mock_table.DataBodyRange.Rows.Count = 99
        mock_table.ListColumns = [col1, col2, col3, col4]

        # Create mock worksheet
        mock_ws = Mock()
        mock_ws.Name = "Data"

        info = manager._get_table_info(mock_table, mock_ws)

        assert info.name == "tbl_Sales"
        assert info.worksheet_name == "Data"
        assert info.range_address == "$A$1:$D$100"
        assert info.header_row == "$A$1:$D$1"
        assert info.columns == ["Col1", "Col2", "Col3", "Col4"]
        assert info.rows_count == 99

    def test_get_table_info_empty_table(self):
        """Test _get_table_info with empty table (no data rows)."""
        mock_excel_mgr = Mock()
        manager = TableManager(mock_excel_mgr)

        # Create mock columns
        col1 = Mock()
        col1.Name = "Col1"

        # Create mock table with no data
        mock_table = Mock()
        mock_table.Name = "tbl_Empty"
        mock_table.Range.Address = "$A$1:$D$1"
        mock_table.HeaderRowRange.Address = "$A$1:$D$1"
        mock_table.DataBodyRange = None  # No data rows
        mock_table.ListColumns = [col1]

        mock_ws = Mock()
        mock_ws.Name = "Sheet1"

        info = manager._get_table_info(mock_table, mock_ws)

        assert info.name == "tbl_Empty"
        assert info.rows_count == 0


class TestTableManagerCreate:
    """Tests for TableManager.create() method."""

    def test_create_in_active_worksheet(self):
        """Test creating table in active worksheet."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock active workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock active worksheet
        mock_ws = Mock()
        mock_ws.Name = "Data"
        mock_wb.ActiveSheet = mock_ws

        # Mock Range and table creation
        mock_range = Mock()
        mock_ws.Range.return_value = mock_range

        # Mock columns
        col1, col2, col3, col4 = Mock(), Mock(), Mock(), Mock()
        col1.Name, col2.Name = "A", "B"
        col3.Name, col4.Name = "C", "D"

        mock_table = Mock()
        mock_table.Name = "tbl_Sales"
        mock_table.Range.Address = "$A$1:$D$100"
        mock_table.HeaderRowRange.Address = "$A$1:$D$1"
        mock_table.DataBodyRange.Rows.Count = 99
        mock_table.ListColumns = [col1, col2, col3, col4]

        # Mock ListObjects with Add method
        mock_list_objects = Mock()
        mock_list_objects.Add.return_value = mock_table
        mock_list_objects.__iter__ = Mock(return_value=iter([]))  # Empty iterator
        mock_ws.ListObjects = mock_list_objects

        # Mock Worksheets with single worksheet (no existing tables)
        mock_wb.Worksheets = [mock_ws]

        manager = TableManager(mock_excel_mgr)
        info = manager.create("tbl_Sales", "A1:D100")

        assert info.name == "tbl_Sales"
        assert info.worksheet_name == "Data"
        assert info.range_address == "$A$1:$D$100"
        assert info.rows_count == 99

        # Verify Add was called with correct parameters
        mock_list_objects.Add.assert_called_once()
        call_args = mock_list_objects.Add.call_args
        assert call_args[1]["SourceType"] == 1
        assert call_args[1]["XlListObjectHasHeaders"] == 1

    def test_create_invalid_table_name(self):
        """Test create with invalid table name."""
        mock_excel_mgr = Mock()
        manager = TableManager(mock_excel_mgr)

        with pytest.raises(TableNameError):
            manager.create("1InvalidName", "A1:D10")

    def test_create_invalid_range(self):
        """Test create with invalid range (empty string)."""
        from unittest.mock import patch

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        mock_ws = Mock()
        mock_ws.Name = "Sheet1"
        mock_ws.Range.side_effect = Exception("Invalid range")
        mock_ws.ListObjects = []
        mock_wb.ActiveSheet = mock_ws
        mock_wb.Worksheets = [mock_ws]

        manager = TableManager(mock_excel_mgr)

        with pytest.raises(TableRangeError):
            manager.create("tbl_Valid", "")

    def test_create_duplicate_name(self):
        """Test create with duplicate table name."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock worksheet
        mock_ws = Mock()
        mock_ws.Name = "Data"
        mock_wb.ActiveSheet = mock_ws

        # Mock existing table with same name
        mock_existing_table = Mock()
        mock_existing_table.Name = "tbl_Sales"

        mock_sheet = Mock()
        mock_sheet.ListObjects = [mock_existing_table]
        mock_wb.Worksheets = [mock_sheet]

        manager = TableManager(mock_excel_mgr)

        with pytest.raises(TableAlreadyExistsError) as exc_info:
            manager.create("tbl_Sales", "A1:D10")

        assert exc_info.value.name == "tbl_Sales"
        assert exc_info.value.workbook_name == "Test.xlsx"

    def test_create_with_specific_worksheet(self):
        """Test creating table in a specific worksheet."""
        from unittest.mock import patch

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock columns
        col1, col2, col3, col4 = Mock(), Mock(), Mock(), Mock()
        col1.Name, col2.Name = "A", "B"
        col3.Name, col4.Name = "C", "D"

        # Mock table
        mock_table = Mock()
        mock_table.Name = "tbl_Test"
        mock_table.Range.Address = "$A$1:$D$10"
        mock_table.HeaderRowRange.Address = "$A$1:$D$1"
        mock_table.DataBodyRange.Rows.Count = 9
        mock_table.ListColumns = [col1, col2, col3, col4]

        # Mock specific worksheet
        mock_ws = Mock()
        mock_ws.Name = "CustomSheet"
        mock_range = Mock()
        mock_ws.Range.return_value = mock_range
        mock_ws.ListObjects = MagicMock()
        mock_ws.ListObjects.Add.return_value = mock_table
        mock_ws.ListObjects.__iter__ = Mock(return_value=iter([]))

        # No existing tables in workbook
        mock_wb.Worksheets = []

        manager = TableManager(mock_excel_mgr)

        with patch("xlmanage.table_manager._resolve_workbook", return_value=mock_wb):
            with patch(
                "xlmanage.table_manager._find_worksheet", return_value=mock_ws
            ):
                info = manager.create("tbl_Test", "A1:D10", worksheet="CustomSheet")

                assert info.name == "tbl_Test"
                assert info.worksheet_name == "CustomSheet"
                assert info.rows_count == 9


class TestTableManagerDelete:
    """Tests for TableManager.delete() method."""

    def test_delete_from_active_workbook(self):
        """Test deleting table searching all worksheets (default: Unlist)."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock worksheet with table
        mock_ws = Mock()
        mock_ws.Name = "Data"

        # Mock table
        mock_table = Mock()
        mock_table.Name = "tbl_Sales"
        mock_ws.ListObjects = [mock_table]

        mock_wb.Worksheets = [mock_ws]

        manager = TableManager(mock_excel_mgr)
        manager.delete("tbl_Sales")

        # Default (force=False) calls Unlist, not Delete
        mock_table.Unlist.assert_called_once()

    def test_delete_from_specific_worksheet(self):
        """Test deleting table from specific worksheet with force=True."""
        from unittest.mock import patch

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock worksheet
        mock_ws = Mock()
        mock_ws.Name = "Data"

        # Mock table
        mock_table = Mock()
        mock_table.Name = "tbl_Sales"
        mock_ws.ListObjects = [mock_table]

        mock_wb.Worksheets = [mock_ws]

        manager = TableManager(mock_excel_mgr)

        with patch("xlmanage.table_manager._resolve_workbook", return_value=mock_wb):
            with patch("xlmanage.table_manager._find_worksheet", return_value=mock_ws):
                manager.delete("tbl_Sales", worksheet="Data", force=True)

        # force=True calls Delete
        mock_table.Delete.assert_called_once()

    def test_delete_table_not_found(self):
        """Test delete with non-existent table."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook with empty worksheet
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        mock_ws = Mock()
        mock_ws.Name = "Data"
        mock_ws.ListObjects = []
        mock_wb.Worksheets = [mock_ws]

        manager = TableManager(mock_excel_mgr)

        with pytest.raises(TableNotFoundError) as exc_info:
            manager.delete("tbl_Missing")

        assert exc_info.value.name == "tbl_Missing"


class TestTableManagerList:
    """Tests for TableManager.list() method."""

    def _make_mock_table(self, name, range_addr, header_addr, rows, col_names):
        """Helper to create a mock table with ListColumns."""
        mock_table = Mock()
        mock_table.Name = name
        mock_table.Range.Address = range_addr
        mock_table.HeaderRowRange.Address = header_addr
        mock_table.DataBodyRange.Rows.Count = rows
        cols = []
        for cn in col_names:
            c = Mock()
            c.Name = cn
            cols.append(c)
        mock_table.ListColumns = cols
        return mock_table

    def test_list_all_tables_in_workbook(self):
        """Test listing all tables in workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock first worksheet with one table
        mock_ws1 = Mock()
        mock_ws1.Name = "Sheet1"
        mock_table1 = self._make_mock_table(
            "tbl_Sales", "$A$1:$D$10", "$A$1:$D$1", 9, ["A", "B", "C", "D"]
        )
        mock_ws1.ListObjects = [mock_table1]

        # Mock second worksheet with one table
        mock_ws2 = Mock()
        mock_ws2.Name = "Sheet2"
        mock_table2 = self._make_mock_table(
            "tbl_Products", "$A$1:$C$20", "$A$1:$C$1", 19, ["A", "B", "C"]
        )
        mock_ws2.ListObjects = [mock_table2]

        mock_wb.Worksheets = [mock_ws1, mock_ws2]

        manager = TableManager(mock_excel_mgr)
        tables = manager.list()

        assert len(tables) == 2
        assert tables[0].name == "tbl_Sales"
        assert tables[0].worksheet_name == "Sheet1"
        assert tables[1].name == "tbl_Products"
        assert tables[1].worksheet_name == "Sheet2"

    def test_list_tables_in_specific_worksheet(self):
        """Test listing tables in specific worksheet."""
        from unittest.mock import patch

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock worksheet
        mock_ws = Mock()
        mock_ws.Name = "Data"
        mock_table = self._make_mock_table(
            "tbl_Sales", "$A$1:$D$10", "$A$1:$D$1", 9, ["A", "B", "C", "D"]
        )
        mock_ws.ListObjects = [mock_table]

        mock_wb.Worksheets = [mock_ws]

        manager = TableManager(mock_excel_mgr)

        with patch("xlmanage.table_manager._resolve_workbook", return_value=mock_wb):
            with patch("xlmanage.table_manager._find_worksheet", return_value=mock_ws):
                tables = manager.list(worksheet="Data")

        assert len(tables) == 1
        assert tables[0].name == "tbl_Sales"
        assert tables[0].worksheet_name == "Data"

    def test_list_empty_workbook(self):
        """Test listing tables in workbook with no tables."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook with no tables
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        mock_ws = Mock()
        mock_ws.Name = "Sheet1"
        mock_ws.ListObjects = []
        mock_wb.Worksheets = [mock_ws]

        manager = TableManager(mock_excel_mgr)
        tables = manager.list()

        assert len(tables) == 0

    def test_list_handles_corrupted_table(self):
        """Test list continues when a table can't be read."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock worksheet
        mock_ws = Mock()
        mock_ws.Name = "Data"

        # Mock corrupted table (raises exception when accessing Name)
        mock_table1 = Mock()
        type(mock_table1).Name = property(lambda self: (_ for _ in ()).throw(Exception("Corrupted")))

        # Mock valid table
        mock_table2 = Mock()
        mock_table2.Name = "tbl_Valid"
        mock_table2.Range.Address = "$A$1:$D$10"
        mock_table2.HeaderRowRange.Address = "$A$1:$D$1"
        mock_table2.DataBodyRange.Rows.Count = 9
        cols = [Mock(Name="A"), Mock(Name="B"), Mock(Name="C"), Mock(Name="D")]
        mock_table2.ListColumns = cols

        mock_ws.ListObjects = [mock_table1, mock_table2]
        mock_wb.Worksheets = [mock_ws]

        manager = TableManager(mock_excel_mgr)
        tables = manager.list()

        # Should skip corrupted table and return only valid one
        assert len(tables) == 1
        assert tables[0].name == "tbl_Valid"

    def test_list_handles_corrupted_sheet(self):
        """Test list() continues when a sheet is corrupted."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock good worksheet with table
        mock_good_ws = Mock()
        mock_good_ws.Name = "GoodSheet"
        mock_good_table = Mock()
        mock_good_table.Name = "tbl_Good"
        mock_good_table.Range.Address = "$A$1:$B$5"
        mock_good_table.HeaderRowRange.Address = "$A$1:$B$1"
        mock_good_table.DataBodyRange.Rows.Count = 4
        mock_good_table.ListColumns = [Mock(Name="A"), Mock(Name="B")]
        mock_good_ws.ListObjects = [mock_good_table]

        # Mock corrupted sheet that raises exception when accessing ListObjects
        mock_bad_ws = Mock()
        mock_bad_ws.Name = "BadSheet"
        type(mock_bad_ws).ListObjects = property(
            lambda self: (_ for _ in ()).throw(Exception("Corrupted sheet"))
        )

        mock_wb.Worksheets = [mock_good_ws, mock_bad_ws]

        manager = TableManager(mock_excel_mgr)
        tables = manager.list()

        # Should return only the good table, skipping corrupted sheet
        assert len(tables) == 1
        assert tables[0].name == "tbl_Good"
        assert tables[0].worksheet_name == "GoodSheet"

    def test_list_specific_worksheet_with_corrupted_table(self):
        """Test list() skips corrupted tables in specific worksheet."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        # Mock worksheet
        mock_ws = Mock()
        mock_ws.Name = "DataSheet"

        # Good table
        mock_good_table = Mock()
        mock_good_table.Name = "tbl_Good"
        mock_good_table.Range.Address = "$A$1:$C$10"
        mock_good_table.HeaderRowRange.Address = "$A$1:$C$1"
        mock_good_table.DataBodyRange.Rows.Count = 9
        mock_good_table.ListColumns = [Mock(Name="A"), Mock(Name="B"), Mock(Name="C")]

        # Corrupted table that raises exception when accessing properties
        mock_bad_table = Mock()
        type(mock_bad_table).Name = property(
            lambda self: (_ for _ in ()).throw(Exception("Corrupted table"))
        )

        mock_ws.ListObjects = [mock_good_table, mock_bad_table]

        # Mock _find_worksheet to return the worksheet
        from unittest.mock import patch

        manager = TableManager(mock_excel_mgr)

        with patch("xlmanage.table_manager._resolve_workbook", return_value=mock_wb):
            with patch("xlmanage.table_manager._find_worksheet", return_value=mock_ws):
                tables = manager.list(worksheet="DataSheet")

                # Should return only good table, skip corrupted one
                assert len(tables) == 1
                assert tables[0].name == "tbl_Good"


class TestFindTableEdgeCases:
    """Tests for _find_table edge cases (corrupted sheet)."""

    def test_find_table_skips_corrupted_sheet(self):
        """Test _find_table skips sheets that raise on ListObjects."""
        from xlmanage.table_manager import _find_table

        mock_wb = Mock()

        # Bad sheet raises on ListObjects
        mock_bad_ws = Mock()
        mock_bad_ws.Name = "Bad"
        type(mock_bad_ws).ListObjects = property(
            lambda self: (_ for _ in ()).throw(Exception("Corrupted"))
        )

        # Good sheet has the target table
        mock_good_ws = Mock()
        mock_good_ws.Name = "Good"
        mock_table = Mock()
        mock_table.Name = "tbl_Target"
        mock_good_ws.ListObjects = [mock_table]

        mock_wb.Worksheets = [mock_bad_ws, mock_good_ws]

        result = _find_table(mock_wb, "tbl_Target")
        assert result is not None
        ws, table = result
        assert table.Name == "tbl_Target"

    def test_find_table_skips_corrupted_table_name(self):
        """Test _find_table skips tables whose Name raises."""
        from xlmanage.table_manager import _find_table

        mock_wb = Mock()
        mock_ws = Mock()
        mock_ws.Name = "Sheet1"

        mock_bad_table = Mock()
        type(mock_bad_table).Name = property(
            lambda self: (_ for _ in ()).throw(Exception("Corrupted"))
        )
        mock_good_table = Mock()
        mock_good_table.Name = "tbl_Target"

        mock_ws.ListObjects = [mock_bad_table, mock_good_table]
        mock_wb.Worksheets = [mock_ws]

        result = _find_table(mock_wb, "tbl_Target")
        assert result is not None
        _, table = result
        assert table.Name == "tbl_Target"


class TestRangesOverlap:
    """Tests for _ranges_overlap function."""

    def test_ranges_overlap_true(self):
        """Test overlapping ranges return True."""
        from xlmanage.table_manager import _ranges_overlap

        mock_range1 = Mock()
        mock_range2 = Mock()
        mock_app = Mock()
        mock_range1.Application = mock_app
        mock_app.Intersect.return_value = Mock()  # non-None = overlap

        assert _ranges_overlap(mock_range1, mock_range2) is True

    def test_ranges_overlap_false(self):
        """Test non-overlapping ranges return False."""
        from xlmanage.table_manager import _ranges_overlap

        mock_range1 = Mock()
        mock_range2 = Mock()
        mock_app = Mock()
        mock_range1.Application = mock_app
        mock_app.Intersect.return_value = None

        assert _ranges_overlap(mock_range1, mock_range2) is False

    def test_ranges_overlap_exception(self):
        """Test exception returns False."""
        from xlmanage.table_manager import _ranges_overlap

        mock_range1 = Mock()
        mock_range2 = Mock()
        mock_range1.Application.Intersect.side_effect = Exception("COM error")

        assert _ranges_overlap(mock_range1, mock_range2) is False


class TestValidateRangeOverlap:
    """Tests for _validate_range overlap detection."""

    def test_validate_range_overlap_raises(self):
        """Test _validate_range raises when range overlaps existing table."""
        from unittest.mock import patch

        from xlmanage.table_manager import _validate_range

        mock_ws = Mock()
        mock_range = Mock()
        mock_ws.Range.return_value = mock_range

        mock_table = Mock()
        mock_table.Name = "tbl_Existing"
        mock_table.Range = Mock()
        mock_ws.ListObjects = [mock_table]

        with patch("xlmanage.table_manager._ranges_overlap", return_value=True):
            with pytest.raises(TableRangeError) as exc_info:
                _validate_range(mock_ws, "A1:D10")
            assert "overlaps" in str(exc_info.value)

    def test_validate_range_overlap_skip_unreadable_table(self):
        """Test _validate_range skips tables that raise on Range access."""
        from unittest.mock import patch

        from xlmanage.table_manager import _validate_range

        mock_ws = Mock()
        mock_range = Mock()
        mock_ws.Range.return_value = mock_range

        # Table whose Range property raises
        mock_bad_table = Mock()
        type(mock_bad_table).Range = property(
            lambda self: (_ for _ in ()).throw(Exception("Corrupted"))
        )
        mock_ws.ListObjects = [mock_bad_table]

        with patch("xlmanage.table_manager._ranges_overlap", return_value=False):
            result = _validate_range(mock_ws, "A1:D10")
        assert result == mock_range


class TestDeleteWithWorksheet:
    """Tests for delete() with specific worksheet parameter."""

    def test_delete_from_specific_worksheet_not_found(self):
        """Test delete raises when table not in specific worksheet."""
        from unittest.mock import patch

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_ws = Mock()
        mock_ws.Name = "Data"
        mock_ws.ListObjects = []

        manager = TableManager(mock_excel_mgr)
        with patch("xlmanage.table_manager._resolve_workbook", return_value=mock_wb):
            with patch("xlmanage.table_manager._find_worksheet", return_value=mock_ws):
                with pytest.raises(TableNotFoundError):
                    manager.delete("tbl_Missing", worksheet="Data")

    def test_delete_specific_worksheet_corrupted_table_skipped(self):
        """Test delete skips corrupted tables when searching specific worksheet."""
        from unittest.mock import patch

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_ws = Mock()
        mock_ws.Name = "Data"

        # Corrupted table
        mock_bad = Mock()
        type(mock_bad).Name = property(
            lambda self: (_ for _ in ()).throw(Exception("Corrupted"))
        )

        # Good table
        mock_good = Mock()
        mock_good.Name = "tbl_Target"

        mock_ws.ListObjects = [mock_bad, mock_good]

        manager = TableManager(mock_excel_mgr)
        with patch("xlmanage.table_manager._resolve_workbook", return_value=mock_wb):
            with patch("xlmanage.table_manager._find_worksheet", return_value=mock_ws):
                manager.delete("tbl_Target", worksheet="Data")

        mock_good.Unlist.assert_called_once()
