"""
Tests for WorkbookManager functionality.
"""

import pytest
from pathlib import Path
from unittest.mock import Mock

from xlmanage.workbook_manager import (
    WorkbookInfo,
    FILE_FORMAT_MAP,
    _detect_file_format,
)


class TestWorkbookInfo:
    """Tests for WorkbookInfo dataclass."""

    def test_workbook_info_creation(self):
        """Test creating WorkbookInfo instance."""
        info = WorkbookInfo(
            name="test.xlsx",
            full_path=Path("C:/data/test.xlsx"),
            read_only=False,
            saved=True,
            sheets_count=3,
        )

        assert info.name == "test.xlsx"
        assert info.full_path == Path("C:/data/test.xlsx")
        assert info.read_only is False
        assert info.saved is True
        assert info.sheets_count == 3

    def test_workbook_info_fields(self):
        """Test all fields are accessible."""
        info = WorkbookInfo(
            name="data.xlsm",
            full_path=Path("D:/projects/data.xlsm"),
            read_only=True,
            saved=False,
            sheets_count=5,
        )

        # Verify all fields
        assert isinstance(info.name, str)
        assert isinstance(info.full_path, Path)
        assert isinstance(info.read_only, bool)
        assert isinstance(info.saved, bool)
        assert isinstance(info.sheets_count, int)


class TestFileFormatMap:
    """Tests for FILE_FORMAT_MAP constant."""

    def test_file_format_map_keys(self):
        """Test FILE_FORMAT_MAP has all expected extensions."""
        expected_extensions = {".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx"}
        assert set(FILE_FORMAT_MAP.keys()) == expected_extensions

    def test_file_format_map_values(self):
        """Test FILE_FORMAT_MAP values are correct."""
        assert FILE_FORMAT_MAP[".xlsx"] == 51
        assert FILE_FORMAT_MAP[".xlsm"] == 52
        assert FILE_FORMAT_MAP[".xls"] == 56
        assert FILE_FORMAT_MAP[".xlsb"] == 50
        assert FILE_FORMAT_MAP[".xltx"] == 54


class TestDetectFileFormat:
    """Tests for _detect_file_format function."""

    def test_detect_xlsx_format(self):
        """Test detecting .xlsx format."""
        path = Path("C:/data/file.xlsx")
        assert _detect_file_format(path) == 51

    def test_detect_xlsm_format(self):
        """Test detecting .xlsm format."""
        path = Path("D:/projects/macro.xlsm")
        assert _detect_file_format(path) == 52

    def test_detect_xls_format(self):
        """Test detecting .xls format."""
        path = Path("E:/legacy/old.xls")
        assert _detect_file_format(path) == 56

    def test_detect_xlsb_format(self):
        """Test detecting .xlsb format."""
        path = Path("F:/binary/data.xlsb")
        assert _detect_file_format(path) == 50

    def test_detect_format_case_insensitive(self):
        """Test format detection is case-insensitive."""
        assert _detect_file_format(Path("test.XLSX")) == 51
        assert _detect_file_format(Path("test.XlSm")) == 52
        assert _detect_file_format(Path("test.XLS")) == 56
        assert _detect_file_format(Path("test.XLSB")) == 50

    def test_detect_format_unsupported_extension(self):
        """Test ValueError is raised for unsupported extensions."""
        with pytest.raises(ValueError) as exc_info:
            _detect_file_format(Path("document.docx"))

        assert "Unsupported file extension" in str(exc_info.value)
        assert ".docx" in str(exc_info.value)
        assert ".xlsx" in str(exc_info.value)  # Lists supported formats

    def test_detect_format_no_extension(self):
        """Test ValueError for files without extension."""
        with pytest.raises(ValueError) as exc_info:
            _detect_file_format(Path("file_without_extension"))

        assert "Unsupported file extension" in str(exc_info.value)

    def test_detect_format_wrong_extension(self):
        """Test ValueError for wrong extensions."""
        invalid_files = [
            Path("data.csv"),
            Path("data.txt"),
            Path("data.pdf"),
            Path("data.xml"),
        ]

        for path in invalid_files:
            with pytest.raises(ValueError):
                _detect_file_format(path)

    def test_detect_xltx_format(self):
        """Test detecting .xltx format."""
        path = Path("C:/templates/template.xltx")
        assert _detect_file_format(path) == 54
