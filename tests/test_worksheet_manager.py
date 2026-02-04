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
        SHEET_NAME_MAX_LENGTH,
        SHEET_NAME_FORBIDDEN_CHARS,
    )
except ImportError:
    from xlmanage.worksheet_manager import (
        WorksheetInfo,
        _validate_sheet_name,
        SHEET_NAME_MAX_LENGTH,
        SHEET_NAME_FORBIDDEN_CHARS,
    )

from xlmanage.exceptions import WorksheetNameError, ExcelManageError


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
