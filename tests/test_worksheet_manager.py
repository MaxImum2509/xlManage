"""
Tests for WorksheetManager functionality.

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
import re

try:
    from xlmanage.worksheet_manager import (
        WorksheetInfo,
        WorksheetManager,
        _validate_sheet_name,
        _resolve_workbook,
        _find_worksheet,
        SHEET_NAME_MAX_LENGTH,
        SHEET_NAME_FORBIDDEN_CHARS,
    )
except ImportError:
    from xlmanage.worksheet_manager import (
        WorksheetInfo,
        WorksheetManager,
        _validate_sheet_name,
        _resolve_workbook,
        _find_worksheet,
        SHEET_NAME_MAX_LENGTH,
        SHEET_NAME_FORBIDDEN_CHARS,
    )

from xlmanage.exceptions import (
    WorksheetNameError,
    ExcelManageError,
    WorkbookNotFoundError,
    ExcelConnectionError,
    WorksheetAlreadyExistsError,
    WorksheetNotFoundError,
    WorksheetDeleteError,
)
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch, PropertyMock


class TestWorksheetInfo:
    """Tests for WorksheetInfo dataclass."""

    def test_worksheet_info_creation(self):
        """Test creating WorksheetInfo instance."""
        info = WorksheetInfo(
            name="Sheet1",
            index=1,
            visible=True,
            rows_used=100,
            columns_used=10,
        )

        assert info.name == "Sheet1"
        assert info.index == 1
        assert info.visible is True
        assert info.rows_used == 100
        assert info.columns_used == 10

    def test_worksheet_info_fields(self):
        """Test all fields are accessible."""
        info = WorksheetInfo(
            name="DataSheet",
            index=2,
            visible=False,
            rows_used=0,
            columns_used=0,
        )

        assert isinstance(info.name, str)
        assert isinstance(info.index, int)
        assert isinstance(info.visible, bool)
        assert isinstance(info.rows_used, int)
        assert isinstance(info.columns_used, int)

    def test_worksheet_info_hidden_sheet(self):
        """Test hidden sheet info."""
        info = WorksheetInfo(
            name="HiddenSheet",
            index=3,
            visible=False,
            rows_used=50,
            columns_used=5,
        )

        assert info.visible is False

    def test_worksheet_info_zero_rows_columns(self):
        """Test empty worksheet info."""
        info = WorksheetInfo(
            name="EmptySheet",
            index=4,
            visible=True,
            rows_used=0,
            columns_used=0,
        )

        assert info.rows_used == 0
        assert info.columns_used == 0


class TestValidateSheetName:
    """Tests for _validate_sheet_name function."""

    def test_validate_sheet_name_simple_valid(self):
        """Test simple valid sheet names."""
        valid_names = ["Sheet1", "Data", "Summary", "Report-Q1", "Test_A"]
        for name in valid_names:
            _validate_sheet_name(name)

    def test_validate_sheet_name_max_length(self):
        """Test sheet name with exactly 31 characters."""
        name = "A" * 31
        _validate_sheet_name(name)

    def test_validate_sheet_name_too_long(self):
        """Test sheet name exceeding 31 characters raises error."""
        name = "B" * 32
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "exceeds 31 characters" in str(exc_info.value).lower()
        assert "32" in str(exc_info.value)

    def test_validate_sheet_name_empty(self):
        """Test empty sheet name raises error."""
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name("")

        assert exc_info.value.name == ""
        assert "cannot be empty" in str(exc_info.value).lower()

    def test_validate_sheet_name_whitespace_only(self):
        """Test whitespace-only sheet name raises error."""
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name("   ")

        assert exc_info.value.name == "   "
        assert "cannot be empty" in str(exc_info.value).lower()

    def test_validate_sheet_name_forbidden_backslash(self):
        """Test sheet name with backslash raises error."""
        name = "Sheet\\1"
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "forbidden character" in str(exc_info.value).lower()
        assert "\\" in str(exc_info.value)

    def test_validate_sheet_name_forbidden_forward_slash(self):
        """Test sheet name with forward slash raises error."""
        name = "Sheet/1"
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "forbidden character" in str(exc_info.value).lower()
        assert "/" in str(exc_info.value)

    def test_validate_sheet_name_forbidden_asterisk(self):
        """Test sheet name with asterisk raises error."""
        name = "Sheet*1"
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "forbidden character" in str(exc_info.value).lower()
        assert "*" in str(exc_info.value)

    def test_validate_sheet_name_forbidden_question_mark(self):
        """Test sheet name with question mark raises error."""
        name = "Sheet?1"
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "forbidden character" in str(exc_info.value).lower()
        assert "?" in str(exc_info.value)

    def test_validate_sheet_name_forbidden_colon(self):
        """Test sheet name with colon raises error."""
        name = "Sheet:1"
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "forbidden character" in str(exc_info.value).lower()
        assert ":" in str(exc_info.value)

    def test_validate_sheet_name_forbidden_bracket(self):
        """Test sheet name with bracket raises error."""
        name = "Sheet[1]"
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "forbidden character" in str(exc_info.value).lower()
        assert "[" in str(exc_info.value)

    def test_validate_sheet_name_forbidden_bracket_close(self):
        """Test sheet name with close bracket raises error."""
        name = "Sheet]1"
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "forbidden character" in str(exc_info.value).lower()
        assert "]" in str(exc_info.value)

    def test_validate_sheet_name_multiple_forbidden_chars(self):
        """Test sheet name with multiple forbidden characters."""
        name = "Sheet/1*?2"
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name(name)

        assert exc_info.value.name == name
        assert "forbidden character" in str(exc_info.value).lower()

    def test_validate_sheet_name_complex_valid(self):
        """Test complex valid sheet names."""
        valid_names = [
            "Data-2024",
            "Q1_Sales",
            "Budget-Test",
            "Summary_A",
            "Q1",
            "Q2",
            "Q3",
            "Q4",
            "Sheet 1",
            "Data (Jan)",
            "Sales Report",
        ]
        for name in valid_names:
            _validate_sheet_name(name)

    def test_validate_sheet_name_unicode_valid(self):
        """Test sheet names with Unicode characters."""
        valid_names = ["Données", "Rapport", "Été", "Résumé", "Test_été"]
        for name in valid_names:
            _validate_sheet_name(name)

    def test_validate_sheet_name_error_inheritance(self):
        """Test WorksheetNameError inherits from ExcelManageError."""
        with pytest.raises(WorksheetNameError) as exc_info:
            _validate_sheet_name("Sheet*Invalid")

        assert isinstance(exc_info.value, WorksheetNameError)
        assert isinstance(exc_info.value, ExcelManageError)
        assert isinstance(exc_info.value, Exception)


class TestValidationConstants:
    """Tests for validation constants."""

    def test_sheet_name_max_length_constant(self):
        """Test SHEET_NAME_MAX_LENGTH constant."""
        assert SHEET_NAME_MAX_LENGTH == 31

    def test_sheet_name_forbidden_chars_constant(self):
        """Test SHEET_NAME_FORBIDDEN_CHARS constant."""
        assert r"\/\*\?:\[\]" in SHEET_NAME_FORBIDDEN_CHARS

    def test_forbidden_chars_regex(self):
        """Test forbidden characters pattern works."""
        pattern = f"[{SHEET_NAME_FORBIDDEN_CHARS}]"
        assert re.search(pattern, "Sheet/1")
        assert re.search(pattern, "Sheet*1")
        assert re.search(pattern, "Sheet?1")
        assert re.search(pattern, "Sheet:1")
        assert re.search(pattern, "Sheet[1]")
        assert not re.search(pattern, "Sheet1")

    def test_forbidden_chars_coverage(self):
        """Test all forbidden characters are covered."""
        forbidden = r"\\/\*\?:\[\]"
        for char in ["\\", "/", "*", "?", ":", "[", "]"]:
            assert char in forbidden


class TestResolveWorkbook:
    """Tests for _resolve_workbook function."""

    def test_resolve_workbook_with_none_returns_active(self):
        """Test resolving with None returns active workbook."""
        # Mock Excel app with active workbook
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.Name = "Active.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        result = _resolve_workbook(mock_app, None)

        assert result == mock_wb
        assert mock_app.ActiveWorkbook == mock_wb

    def test_resolve_workbook_with_none_no_active_raises(self):
        """Test resolving with None when no active workbook raises error."""
        # Mock Excel app with no active workbook
        mock_app = Mock()
        mock_app.ActiveWorkbook = None

        with pytest.raises(ExcelConnectionError) as exc_info:
            _resolve_workbook(mock_app, None)

        assert "No active workbook" in str(exc_info.value)
        assert exc_info.value.hresult == 0x80080005

    def test_resolve_workbook_with_none_com_error_raises(self):
        """Test resolving with None when COM error occurs."""

        # Mock Excel app that raises COM error
        class COMError(Exception):
            def __init__(self):
                self.hresult = 0x800401E4

        mock_app = Mock()
        # Configure the property to raise exception when accessed
        type(mock_app).ActiveWorkbook = property(lambda self: (_ for _ in ()).throw(COMError()))

        with pytest.raises(ExcelConnectionError) as exc_info:
            _resolve_workbook(mock_app, None)

        assert exc_info.value.hresult == 0x800401E4

    def test_resolve_workbook_with_path_finds_open(self):
        """Test resolving with path finds open workbook."""
        # Mock Excel app with open workbook
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.Name = "test.xlsx"

        # Patch _find_open_workbook in workbook_manager module
        with patch("xlmanage.workbook_manager._find_open_workbook") as mock_find:
            mock_find.return_value = mock_wb

            result = _resolve_workbook(mock_app, Path("C:/data/test.xlsx"))

            assert result == mock_wb
            mock_find.assert_called_once_with(mock_app, Path("C:/data/test.xlsx"))

    def test_resolve_workbook_with_path_not_open_raises(self):
        """Test resolving with path when workbook not open raises error."""
        # Mock Excel app
        mock_app = Mock()

        # Patch _find_open_workbook to return None
        with patch("xlmanage.workbook_manager._find_open_workbook") as mock_find:
            mock_find.return_value = None

            with pytest.raises(WorkbookNotFoundError) as exc_info:
                _resolve_workbook(mock_app, Path("C:/data/missing.xlsx"))

            assert "is not open" in str(exc_info.value)
            assert exc_info.value.path == Path("C:/data/missing.xlsx")

    def test_resolve_workbook_preserves_workbook_object(self):
        """Test that resolved workbook object is returned unchanged."""
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.Name = "test.xlsx"
        mock_wb.FullName = "C:/data/test.xlsx"
        mock_app.ActiveWorkbook = mock_wb

        result = _resolve_workbook(mock_app, None)

        assert result is mock_wb
        assert result.Name == "test.xlsx"
        assert result.FullName == "C:/data/test.xlsx"

    def test_resolve_workbook_with_none_non_com_error_raises(self):
        """Test resolving with None when non-COM error occurs."""
        # Mock Excel app that raises exception without hresult
        mock_app = Mock()
        type(mock_app).ActiveWorkbook = property(
            lambda self: (_ for _ in ()).throw(RuntimeError("Some error"))
        )

        with pytest.raises(RuntimeError) as exc_info:
            _resolve_workbook(mock_app, None)

        assert "Some error" in str(exc_info.value)


class TestFindWorksheet:
    """Tests for _find_worksheet function."""

    def test_find_worksheet_exact_match(self):
        """Test finding worksheet with exact name match."""
        # Mock workbook with worksheets
        mock_wb = Mock()
        mock_ws1 = Mock()
        mock_ws1.Name = "Sheet1"
        mock_ws2 = Mock()
        mock_ws2.Name = "Sheet2"
        mock_wb.Worksheets = [mock_ws1, mock_ws2]

        result = _find_worksheet(mock_wb, "Sheet1")

        assert result == mock_ws1
        assert result.Name == "Sheet1"

    def test_find_worksheet_case_insensitive(self):
        """Test finding worksheet is case-insensitive."""
        # Mock workbook with worksheets
        mock_wb = Mock()
        mock_ws = Mock()
        mock_ws.Name = "Sheet1"
        mock_wb.Worksheets = [mock_ws]

        # Test various cases
        assert _find_worksheet(mock_wb, "SHEET1") == mock_ws
        assert _find_worksheet(mock_wb, "sheet1") == mock_ws
        assert _find_worksheet(mock_wb, "ShEeT1") == mock_ws

    def test_find_worksheet_not_found(self):
        """Test finding non-existent worksheet returns None."""
        # Mock workbook with worksheets
        mock_wb = Mock()
        mock_ws = Mock()
        mock_ws.Name = "Sheet1"
        mock_wb.Worksheets = [mock_ws]

        result = _find_worksheet(mock_wb, "NonExistent")

        assert result is None

    def test_find_worksheet_empty_workbook(self):
        """Test finding worksheet in empty workbook returns None."""
        # Mock workbook with no worksheets
        mock_wb = Mock()
        mock_wb.Worksheets = []

        result = _find_worksheet(mock_wb, "Sheet1")

        assert result is None

    def test_find_worksheet_multiple_sheets(self):
        """Test finding worksheet among multiple sheets."""
        # Mock workbook with multiple worksheets
        mock_wb = Mock()
        mock_ws1 = Mock()
        mock_ws1.Name = "Data"
        mock_ws2 = Mock()
        mock_ws2.Name = "Summary"
        mock_ws3 = Mock()
        mock_ws3.Name = "Report"
        mock_wb.Worksheets = [mock_ws1, mock_ws2, mock_ws3]

        result = _find_worksheet(mock_wb, "Summary")

        assert result == mock_ws2
        assert result.Name == "Summary"

    def test_find_worksheet_handles_read_error(self):
        """Test finding worksheet handles errors when reading names."""
        # Mock workbook with worksheets, one that raises error
        mock_wb = Mock()

        # Create mock worksheet that raises error when accessing Name
        mock_ws1 = Mock()

        def raise_error():
            raise Exception("Read error")

        type(mock_ws1).Name = property(lambda self: raise_error())

        mock_ws2 = Mock()
        mock_ws2.Name = "Sheet2"
        mock_wb.Worksheets = [mock_ws1, mock_ws2]

        # Should skip ws1 and find ws2
        result = _find_worksheet(mock_wb, "Sheet2")

        assert result == mock_ws2

    def test_find_worksheet_all_error_returns_none(self):
        """Test finding worksheet when all sheets error returns None."""
        # Mock workbook with worksheets that all raise errors
        mock_wb = Mock()

        mock_ws1 = Mock()

        def raise_error1():
            raise Exception("Read error 1")

        type(mock_ws1).Name = property(lambda self: raise_error1())

        mock_ws2 = Mock()

        def raise_error2():
            raise Exception("Read error 2")

        type(mock_ws2).Name = property(lambda self: raise_error2())

        mock_wb.Worksheets = [mock_ws1, mock_ws2]

        result = _find_worksheet(mock_wb, "Sheet1")

        assert result is None

    def test_find_worksheet_unicode_names(self):
        """Test finding worksheet with Unicode names."""
        # Mock workbook with Unicode worksheet names
        mock_wb = Mock()
        mock_ws1 = Mock()
        mock_ws1.Name = "Données"
        mock_ws2 = Mock()
        mock_ws2.Name = "Été"
        mock_wb.Worksheets = [mock_ws1, mock_ws2]

        result1 = _find_worksheet(mock_wb, "Données")
        assert result1 == mock_ws1

        result2 = _find_worksheet(mock_wb, "été")  # Case-insensitive
        assert result2 == mock_ws2

    def test_find_worksheet_special_characters(self):
        """Test finding worksheet with special characters in name."""
        # Mock workbook with special character names
        mock_wb = Mock()
        mock_ws = Mock()
        mock_ws.Name = "Data (2024)"
        mock_wb.Worksheets = [mock_ws]

        result = _find_worksheet(mock_wb, "Data (2024)")

        assert result == mock_ws

    def test_find_worksheet_returns_first_match(self):
        """Test that finding worksheet returns first match found."""
        # Mock workbook with worksheets
        mock_wb = Mock()
        mock_ws1 = Mock()
        mock_ws1.Name = "Sheet1"
        mock_ws2 = Mock()
        mock_ws2.Name = "Sheet2"
        mock_ws3 = Mock()
        mock_ws3.Name = "Sheet1"  # Duplicate name (shouldn't happen in real Excel)
        mock_wb.Worksheets = [mock_ws1, mock_ws2, mock_ws3]

        result = _find_worksheet(mock_wb, "Sheet1")

        # Should return first match
        assert result == mock_ws1

    def test_find_worksheet_preserves_worksheet_object(self):
        """Test that found worksheet object is returned unchanged."""
        mock_wb = Mock()
        mock_ws = Mock()
        mock_ws.Name = "TestSheet"
        mock_ws.Index = 1
        mock_ws.Visible = True
        mock_wb.Worksheets = [mock_ws]

        result = _find_worksheet(mock_wb, "TestSheet")

        assert result is mock_ws
        assert result.Name == "TestSheet"
        assert result.Index == 1
        assert result.Visible is True


class TestWorksheetManager:
    """Tests for WorksheetManager class."""

    def test_worksheet_manager_initialization(self):
        """Test WorksheetManager initialization."""
        mock_excel_mgr = Mock()
        manager = WorksheetManager(mock_excel_mgr)

        assert manager._mgr == mock_excel_mgr

    def test_get_worksheet_info_with_data(self):
        """Test _get_worksheet_info with worksheet containing data."""
        mock_excel_mgr = Mock()
        manager = WorksheetManager(mock_excel_mgr)

        # Create mock worksheet with data
        mock_ws = Mock()
        mock_ws.Name = "DataSheet"
        mock_ws.Index = 2
        mock_ws.Visible = True

        # Mock UsedRange
        mock_used_range = Mock()
        mock_used_range.Rows.Count = 100
        mock_used_range.Columns.Count = 10
        mock_ws.UsedRange = mock_used_range

        info = manager._get_worksheet_info(mock_ws)

        assert info.name == "DataSheet"
        assert info.index == 2
        assert info.visible is True
        assert info.rows_used == 100
        assert info.columns_used == 10

    def test_get_worksheet_info_empty_sheet(self):
        """Test _get_worksheet_info with empty worksheet."""
        mock_excel_mgr = Mock()
        manager = WorksheetManager(mock_excel_mgr)

        # Create mock worksheet with None UsedRange
        mock_ws = Mock()
        mock_ws.Name = "EmptySheet"
        mock_ws.Index = 1
        mock_ws.Visible = True
        mock_ws.UsedRange = None

        info = manager._get_worksheet_info(mock_ws)

        assert info.name == "EmptySheet"
        assert info.index == 1
        assert info.visible is True
        assert info.rows_used == 0
        assert info.columns_used == 0

    def test_get_worksheet_info_used_range_error(self):
        """Test _get_worksheet_info when UsedRange raises exception."""
        mock_excel_mgr = Mock()
        manager = WorksheetManager(mock_excel_mgr)

        # Create mock worksheet that raises exception on UsedRange
        mock_ws = Mock()
        mock_ws.Name = "ErrorSheet"
        mock_ws.Index = 3
        mock_ws.Visible = False
        type(mock_ws).UsedRange = PropertyMock(
            side_effect=Exception("UsedRange failed")
        )

        info = manager._get_worksheet_info(mock_ws)

        assert info.name == "ErrorSheet"
        assert info.index == 3
        assert info.visible is False
        assert info.rows_used == 0
        assert info.columns_used == 0


class TestWorksheetManagerCreate:
    """Tests for WorksheetManager.create() method."""

    def test_create_in_active_workbook(self):
        """Test creating worksheet in active workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock active workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_wb.Worksheets.Count = 2

        # Mock last worksheet
        mock_last_ws = Mock()
        mock_last_ws.Name = "Sheet2"
        mock_wb.Worksheets = Mock(return_value=mock_last_ws)

        # Mock new worksheet
        mock_new_ws = Mock()
        mock_new_ws.Name = "NewSheet"
        mock_new_ws.Index = 3
        mock_new_ws.Visible = True
        mock_new_ws.UsedRange = None

        # Configure Worksheets.Add to return new worksheet
        mock_wb.Worksheets.Add = Mock(return_value=mock_new_ws)

        manager = WorksheetManager(mock_excel_mgr)

        with patch(
            "xlmanage.worksheet_manager._resolve_workbook"
        ) as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch(
                "xlmanage.worksheet_manager._find_worksheet"
            ) as mock_find:
                mock_find.return_value = None  # Sheet doesn't exist

                info = manager.create("NewSheet")

                # Verify worksheet was created
                assert info.name == "NewSheet"
                assert info.index == 3
                assert info.visible is True
                assert info.rows_used == 0
                assert info.columns_used == 0

                # Verify calls
                mock_resolve.assert_called_once_with(mock_app, None)
                mock_find.assert_called_once_with(mock_wb, "NewSheet")
                mock_wb.Worksheets.Add.assert_called_once_with(
                    After=mock_last_ws
                )
                assert mock_new_ws.Name == "NewSheet"

    def test_create_in_specific_workbook(self):
        """Test creating worksheet in specific workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock specific workbook
        mock_wb = Mock()
        mock_wb.Name = "Specific.xlsx"
        mock_wb.Worksheets.Count = 1

        # Mock last worksheet
        mock_last_ws = Mock()
        mock_wb.Worksheets = Mock(return_value=mock_last_ws)

        # Mock new worksheet
        mock_new_ws = Mock()
        mock_new_ws.Name = "DataSheet"
        mock_new_ws.Index = 2
        mock_new_ws.Visible = True
        mock_new_ws.UsedRange = None

        mock_wb.Worksheets.Add = Mock(return_value=mock_new_ws)

        manager = WorksheetManager(mock_excel_mgr)
        workbook_path = Path("C:/data/Specific.xlsx")

        with patch(
            "xlmanage.worksheet_manager._resolve_workbook"
        ) as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch(
                "xlmanage.worksheet_manager._find_worksheet"
            ) as mock_find:
                mock_find.return_value = None

                info = manager.create("DataSheet", workbook_path)

                assert info.name == "DataSheet"
                assert info.index == 2
                mock_resolve.assert_called_once_with(mock_app, workbook_path)

    def test_create_invalid_name(self):
        """Test creating worksheet with invalid name."""
        mock_excel_mgr = Mock()
        manager = WorksheetManager(mock_excel_mgr)

        # Test empty name
        with pytest.raises(WorksheetNameError) as exc_info:
            manager.create("")

        assert "cannot be empty" in str(exc_info.value).lower()

        # Test name with forbidden character
        with pytest.raises(WorksheetNameError) as exc_info:
            manager.create("Sheet/Invalid")

        assert "forbidden character" in str(exc_info.value).lower()

    def test_create_duplicate_name(self):
        """Test creating worksheet with name that already exists."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock existing worksheet
        mock_existing_ws = Mock()
        mock_existing_ws.Name = "Existing"

        manager = WorksheetManager(mock_excel_mgr)

        with patch(
            "xlmanage.worksheet_manager._resolve_workbook"
        ) as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch(
                "xlmanage.worksheet_manager._find_worksheet"
            ) as mock_find:
                mock_find.return_value = mock_existing_ws  # Sheet exists

                with pytest.raises(WorksheetAlreadyExistsError) as exc_info:
                    manager.create("Existing")

                assert exc_info.value.name == "Existing"
                assert exc_info.value.workbook_name == "Test.xlsx"
                assert "already exists" in str(exc_info.value).lower()

    def test_create_workbook_not_found(self):
        """Test creating worksheet in non-existent workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        manager = WorksheetManager(mock_excel_mgr)
        workbook_path = Path("C:/data/missing.xlsx")

        with patch(
            "xlmanage.worksheet_manager._resolve_workbook"
        ) as mock_resolve:
            mock_resolve.side_effect = WorkbookNotFoundError(
                workbook_path, "Workbook not found"
            )

            with pytest.raises(WorkbookNotFoundError) as exc_info:
                manager.create("NewSheet", workbook_path)

            assert exc_info.value.path == workbook_path

    def test_create_com_error(self):
        """Test creating worksheet when COM error occurs."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_wb.Worksheets.Count = 1

        # Mock last worksheet
        mock_last_ws = Mock()
        mock_wb.Worksheets = Mock(return_value=mock_last_ws)

        # Mock Worksheets.Add to raise COM error
        class COMError(Exception):
            def __init__(self):
                self.hresult = 0x800A03EC

        mock_wb.Worksheets.Add = Mock(side_effect=COMError())

        manager = WorksheetManager(mock_excel_mgr)

        with patch(
            "xlmanage.worksheet_manager._resolve_workbook"
        ) as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch(
                "xlmanage.worksheet_manager._find_worksheet"
            ) as mock_find:
                mock_find.return_value = None

                with pytest.raises(ExcelConnectionError) as exc_info:
                    manager.create("NewSheet")

                assert exc_info.value.hresult == 0x800A03EC
                assert "Failed to create worksheet" in str(exc_info.value)

    def test_create_at_end_of_workbook(self):
        """Test that worksheet is created at the end of workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook with 3 sheets
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_wb.Worksheets.Count = 3

        # Mock last worksheet (index 3)
        mock_last_ws = Mock()
        mock_last_ws.Name = "Sheet3"
        mock_last_ws.Index = 3
        mock_wb.Worksheets = Mock(return_value=mock_last_ws)

        # Mock new worksheet (should be index 4)
        mock_new_ws = Mock()
        mock_new_ws.Name = "LastSheet"
        mock_new_ws.Index = 4
        mock_new_ws.Visible = True
        mock_new_ws.UsedRange = None

        mock_wb.Worksheets.Add = Mock(return_value=mock_new_ws)

        manager = WorksheetManager(mock_excel_mgr)

        with patch(
            "xlmanage.worksheet_manager._resolve_workbook"
        ) as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch(
                "xlmanage.worksheet_manager._find_worksheet"
            ) as mock_find:
                mock_find.return_value = None

                info = manager.create("LastSheet")

                # Verify worksheet was created at index 4 (after 3)
                assert info.index == 4
                assert info.name == "LastSheet"

                # Verify Add was called with After=last_ws
                mock_wb.Worksheets.Add.assert_called_once_with(
                    After=mock_last_ws
                )

    def test_create_preserves_worksheet_info(self):
        """Test that create returns accurate WorksheetInfo."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"
        mock_wb.Worksheets.Count = 1
        mock_last_ws = Mock()
        mock_wb.Worksheets = Mock(return_value=mock_last_ws)

        # Mock new worksheet with specific properties
        mock_new_ws = Mock()
        mock_new_ws.Name = "TestSheet"
        mock_new_ws.Index = 2
        mock_new_ws.Visible = True

        # Mock UsedRange with data
        mock_used_range = Mock()
        mock_used_range.Rows.Count = 50
        mock_used_range.Columns.Count = 5
        mock_new_ws.UsedRange = mock_used_range

        mock_wb.Worksheets.Add = Mock(return_value=mock_new_ws)

        manager = WorksheetManager(mock_excel_mgr)

        with patch(
            "xlmanage.worksheet_manager._resolve_workbook"
        ) as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch(
                "xlmanage.worksheet_manager._find_worksheet"
            ) as mock_find:
                mock_find.return_value = None

                info = manager.create("TestSheet")

                # Verify all fields are correct
                assert info.name == "TestSheet"
                assert info.index == 2
                assert info.visible is True
                assert info.rows_used == 50
                assert info.columns_used == 5
                assert isinstance(info, WorksheetInfo)


class TestWorksheetManagerDelete:
    """Tests for WorksheetManager.delete() method."""

    def test_delete_worksheet_success(self):
        """Test deleting a worksheet successfully."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock worksheet to delete
        mock_ws = Mock()
        mock_ws.Name = "ToDelete"
        mock_ws.Visible = True
        mock_ws.Delete = Mock()

        # Mock other visible sheets (so we have at least 2 visible)
        mock_ws2 = Mock()
        mock_ws2.Visible = True
        mock_wb.Worksheets = [mock_ws, mock_ws2]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = mock_ws

                manager.delete("ToDelete")

                # Verify Delete was called
                mock_ws.Delete.assert_called_once()

                # Verify DisplayAlerts was managed
                assert mock_app.DisplayAlerts is True  # Restored

    def test_delete_from_specific_workbook(self):
        """Test deleting worksheet from specific workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Specific.xlsx"

        # Mock worksheets
        mock_ws = Mock()
        mock_ws.Name = "OldSheet"
        mock_ws.Visible = True
        mock_ws.Delete = Mock()

        mock_ws2 = Mock()
        mock_ws2.Visible = True
        mock_wb.Worksheets = [mock_ws, mock_ws2]

        manager = WorksheetManager(mock_excel_mgr)
        workbook_path = Path("C:/data/Specific.xlsx")

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = mock_ws

                manager.delete("OldSheet", workbook_path)

                mock_resolve.assert_called_once_with(mock_app, workbook_path)
                mock_ws.Delete.assert_called_once()

    def test_delete_worksheet_not_found(self):
        """Test deleting non-existent worksheet raises error."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = None  # Worksheet not found

                with pytest.raises(WorksheetNotFoundError) as exc_info:
                    manager.delete("Missing")

                assert exc_info.value.name == "Missing"
                assert exc_info.value.workbook_name == "Test.xlsx"

    def test_delete_last_visible_sheet_raises_error(self):
        """Test deleting last visible sheet raises error."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook with only 1 visible sheet
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock the only visible worksheet
        mock_ws = Mock()
        mock_ws.Name = "LastSheet"
        mock_ws.Visible = True

        # Mock hidden sheets (not counted)
        mock_ws_hidden = Mock()
        mock_ws_hidden.Visible = False

        mock_wb.Worksheets = [mock_ws, mock_ws_hidden]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = mock_ws

                with pytest.raises(WorksheetDeleteError) as exc_info:
                    manager.delete("LastSheet")

                assert exc_info.value.name == "LastSheet"
                assert "last visible worksheet" in str(exc_info.value).lower()

    def test_delete_hidden_sheet_when_only_one_visible(self):
        """Test deleting hidden sheet when only one visible sheet exists."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock visible sheet
        mock_ws_visible = Mock()
        mock_ws_visible.Visible = True

        # Mock hidden sheet to delete
        mock_ws_hidden = Mock()
        mock_ws_hidden.Name = "HiddenSheet"
        mock_ws_hidden.Visible = False
        mock_ws_hidden.Delete = Mock()

        mock_wb.Worksheets = [mock_ws_visible, mock_ws_hidden]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = mock_ws_hidden

                # Should succeed because we're deleting a hidden sheet
                manager.delete("HiddenSheet")

                mock_ws_hidden.Delete.assert_called_once()

    def test_delete_display_alerts_restored_on_error(self):
        """Test DisplayAlerts is restored even if Delete raises error."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock worksheet that raises error on Delete
        mock_ws = Mock()
        mock_ws.Name = "ErrorSheet"
        mock_ws.Visible = True
        mock_ws.Delete = Mock(side_effect=Exception("Delete failed"))

        mock_ws2 = Mock()
        mock_ws2.Visible = True
        mock_wb.Worksheets = [mock_ws, mock_ws2]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = mock_ws

                # Delete should raise, but DisplayAlerts should be restored
                with pytest.raises(Exception) as exc_info:
                    manager.delete("ErrorSheet")

                assert "Delete failed" in str(exc_info.value)
                # Verify DisplayAlerts was restored
                assert mock_app.DisplayAlerts is True

    def test_delete_with_multiple_visible_sheets(self):
        """Test deleting when multiple visible sheets exist."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook with 3 visible sheets
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        mock_ws1 = Mock()
        mock_ws1.Name = "Sheet1"
        mock_ws1.Visible = True
        mock_ws1.Delete = Mock()

        mock_ws2 = Mock()
        mock_ws2.Visible = True

        mock_ws3 = Mock()
        mock_ws3.Visible = True

        mock_wb.Worksheets = [mock_ws1, mock_ws2, mock_ws3]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = mock_ws1

                manager.delete("Sheet1")

                # Should succeed with 3 visible sheets
                mock_ws1.Delete.assert_called_once()

    def test_delete_handles_worksheet_iteration_error(self):
        """Test delete handles errors during worksheet iteration."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock worksheet to delete
        mock_ws = Mock()
        mock_ws.Name = "ToDelete"
        mock_ws.Visible = True
        mock_ws.Delete = Mock()

        # Mock another visible sheet
        mock_ws2 = Mock()
        mock_ws2.Visible = True

        # Mock sheet that raises error when accessing Visible
        mock_ws_error = Mock()
        type(mock_ws_error).Visible = PropertyMock(side_effect=Exception("Error"))

        mock_wb.Worksheets = [mock_ws, mock_ws2, mock_ws_error]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = mock_ws

                # Should succeed, ignoring the error sheet
                manager.delete("ToDelete")

                mock_ws.Delete.assert_called_once()


class TestWorksheetManagerList:
    """Tests for WorksheetManager.list() method."""

    def test_list_worksheets_success(self):
        """Test listing worksheets successfully."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock worksheets
        mock_ws1 = Mock()
        mock_ws1.Name = "Sheet1"
        mock_ws1.Index = 1
        mock_ws1.Visible = True
        mock_ws1.UsedRange = None

        mock_ws2 = Mock()
        mock_ws2.Name = "Sheet2"
        mock_ws2.Index = 2
        mock_ws2.Visible = False
        mock_ws2.UsedRange = None

        mock_wb.Worksheets = [mock_ws1, mock_ws2]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            sheets = manager.list()

            assert len(sheets) == 2
            assert sheets[0].name == "Sheet1"
            assert sheets[0].index == 1
            assert sheets[0].visible is True
            assert sheets[1].name == "Sheet2"
            assert sheets[1].index == 2
            assert sheets[1].visible is False

    def test_list_from_specific_workbook(self):
        """Test listing worksheets from specific workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Specific.xlsx"

        mock_ws = Mock()
        mock_ws.Name = "Data"
        mock_ws.Index = 1
        mock_ws.Visible = True
        mock_ws.UsedRange = None

        mock_wb.Worksheets = [mock_ws]

        manager = WorksheetManager(mock_excel_mgr)
        workbook_path = Path("C:/data/Specific.xlsx")

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            sheets = manager.list(workbook_path)

            assert len(sheets) == 1
            assert sheets[0].name == "Data"
            mock_resolve.assert_called_once_with(mock_app, workbook_path)

    def test_list_empty_workbook(self):
        """Test listing worksheets in empty workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook with no worksheets
        mock_wb = Mock()
        mock_wb.Name = "Empty.xlsx"
        mock_wb.Worksheets = []

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            sheets = manager.list()

            assert sheets == []

    def test_list_handles_read_error(self):
        """Test listing handles errors when reading worksheet info."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock worksheet that works
        mock_ws1 = Mock()
        mock_ws1.Name = "Sheet1"
        mock_ws1.Index = 1
        mock_ws1.Visible = True
        mock_ws1.UsedRange = None

        # Mock worksheet that raises error
        mock_ws2 = Mock()
        type(mock_ws2).Name = PropertyMock(side_effect=Exception("Read error"))

        mock_wb.Worksheets = [mock_ws1, mock_ws2]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            sheets = manager.list()

            # Should return only the readable worksheet
            assert len(sheets) == 1
            assert sheets[0].name == "Sheet1"

    def test_list_includes_visible_and_hidden(self):
        """Test that list includes both visible and hidden worksheets."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook with mixed visibility
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        mock_ws_visible = Mock()
        mock_ws_visible.Name = "Visible"
        mock_ws_visible.Index = 1
        mock_ws_visible.Visible = True
        mock_ws_visible.UsedRange = None

        mock_ws_hidden = Mock()
        mock_ws_hidden.Name = "Hidden"
        mock_ws_hidden.Index = 2
        mock_ws_hidden.Visible = False
        mock_ws_hidden.UsedRange = None

        mock_wb.Worksheets = [mock_ws_visible, mock_ws_hidden]

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            sheets = manager.list()

            assert len(sheets) == 2
            visible_sheets = [s for s in sheets if s.visible]
            hidden_sheets = [s for s in sheets if not s.visible]
            assert len(visible_sheets) == 1
            assert len(hidden_sheets) == 1


class TestWorksheetManagerCopy:
    """Tests for WorksheetManager.copy() method."""

    def test_copy_worksheet_success(self):
        """Test copying worksheet successfully."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock source worksheet
        mock_ws_source = Mock()
        mock_ws_source.Name = "Template"
        mock_ws_source.Copy = Mock()

        # Mock copied worksheet
        mock_ws_copy = Mock()
        mock_ws_copy.Name = "Copy"
        mock_ws_copy.Index = 2
        mock_ws_copy.Visible = True
        mock_ws_copy.UsedRange = None

        mock_wb.ActiveSheet = mock_ws_copy

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                # First call: find source (exists)
                # Second call: check destination (doesn't exist)
                mock_find.side_effect = [mock_ws_source, None]

                info = manager.copy("Template", "Copy")

                assert info.name == "Copy"
                assert info.index == 2
                assert info.visible is True
                mock_ws_source.Copy.assert_called_once_with(After=mock_ws_source)

    def test_copy_from_specific_workbook(self):
        """Test copying worksheet from specific workbook."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Specific.xlsx"

        # Mock source worksheet
        mock_ws_source = Mock()
        mock_ws_source.Name = "Original"
        mock_ws_source.Copy = Mock()

        # Mock copied worksheet
        mock_ws_copy = Mock()
        mock_ws_copy.Name = "Duplicate"
        mock_ws_copy.Index = 2
        mock_ws_copy.Visible = True
        mock_ws_copy.UsedRange = None

        mock_wb.ActiveSheet = mock_ws_copy

        manager = WorksheetManager(mock_excel_mgr)
        workbook_path = Path("C:/data/Specific.xlsx")

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.side_effect = [mock_ws_source, None]

                info = manager.copy("Original", "Duplicate", workbook_path)

                assert info.name == "Duplicate"
                mock_resolve.assert_called_once_with(mock_app, workbook_path)

    def test_copy_invalid_destination_name(self):
        """Test copying with invalid destination name."""
        mock_excel_mgr = Mock()
        manager = WorksheetManager(mock_excel_mgr)

        # Test invalid names
        with pytest.raises(WorksheetNameError):
            manager.copy("Source", "")

        with pytest.raises(WorksheetNameError):
            manager.copy("Source", "Sheet/Invalid")

    def test_copy_source_not_found(self):
        """Test copying non-existent source worksheet."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.return_value = None  # Source not found

                with pytest.raises(WorksheetNotFoundError) as exc_info:
                    manager.copy("Missing", "Copy")

                assert exc_info.value.name == "Missing"

    def test_copy_destination_already_exists(self):
        """Test copying when destination name already exists."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock source worksheet
        mock_ws_source = Mock()
        mock_ws_source.Name = "Source"

        # Mock existing destination worksheet
        mock_ws_existing = Mock()
        mock_ws_existing.Name = "Existing"

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                # First call: source exists
                # Second call: destination already exists
                mock_find.side_effect = [mock_ws_source, mock_ws_existing]

                with pytest.raises(WorksheetAlreadyExistsError) as exc_info:
                    manager.copy("Source", "Existing")

                assert exc_info.value.name == "Existing"

    def test_copy_com_error(self):
        """Test copying when COM error occurs."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock source worksheet that raises COM error
        class COMError(Exception):
            def __init__(self):
                self.hresult = 0x800A03EC

        mock_ws_source = Mock()
        mock_ws_source.Name = "Source"
        mock_ws_source.Copy = Mock(side_effect=COMError())

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.side_effect = [mock_ws_source, None]

                with pytest.raises(ExcelConnectionError) as exc_info:
                    manager.copy("Source", "Copy")

                assert exc_info.value.hresult == 0x800A03EC
                assert "Failed to copy worksheet 'Source'" in str(exc_info.value)

    def test_copy_placed_after_source(self):
        """Test that copy is placed immediately after source."""
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock workbook
        mock_wb = Mock()
        mock_wb.Name = "Test.xlsx"

        # Mock source worksheet at index 2
        mock_ws_source = Mock()
        mock_ws_source.Name = "Source"
        mock_ws_source.Index = 2
        mock_ws_source.Copy = Mock()

        # Mock copied worksheet at index 3 (after source)
        mock_ws_copy = Mock()
        mock_ws_copy.Name = "Copy"
        mock_ws_copy.Index = 3
        mock_ws_copy.Visible = True
        mock_ws_copy.UsedRange = None

        mock_wb.ActiveSheet = mock_ws_copy

        manager = WorksheetManager(mock_excel_mgr)

        with patch("xlmanage.worksheet_manager._resolve_workbook") as mock_resolve:
            mock_resolve.return_value = mock_wb

            with patch("xlmanage.worksheet_manager._find_worksheet") as mock_find:
                mock_find.side_effect = [mock_ws_source, None]

                info = manager.copy("Source", "Copy")

                # Verify Copy was called with After=source
                mock_ws_source.Copy.assert_called_once_with(After=mock_ws_source)
                # Verify copy is at index 3 (after source at index 2)
                assert info.index == 3
