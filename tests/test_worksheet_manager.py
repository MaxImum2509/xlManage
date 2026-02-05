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
        _validate_sheet_name,
        _resolve_workbook,
        _find_worksheet,
        SHEET_NAME_MAX_LENGTH,
        SHEET_NAME_FORBIDDEN_CHARS,
    )
except ImportError:
    from xlmanage.worksheet_manager import (
        WorksheetInfo,
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
)
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch


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
