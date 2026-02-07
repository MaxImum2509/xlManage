"""
Tests for WorkbookManager CRUD operations.

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


class TestFindOpenWorkbook:
    """Tests for _find_open_workbook function."""

    def test_find_workbook_by_full_path(self):
        """Test finding workbook by complete path."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\test.xlsx"
        mock_wb.Name = "test.xlsx"
        mock_app.Workbooks = [mock_wb]

        # Test
        path = Path("C:/data/test.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Verify
        assert result == mock_wb

    def test_find_workbook_by_name(self):
        """Test finding workbook by filename when paths differ."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock with different path representation
        mock_app = Mock()
        mock_wb = Mock()
        # FullName uses different path format (UNC path, for example)
        mock_wb.FullName = "\\\\server\\share\\test.xlsx"
        mock_wb.Name = "test.xlsx"
        mock_app.Workbooks = [mock_wb]

        # Test with local path
        path = Path("C:/local/test.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Should find by name since full paths don't match
        assert result == mock_wb

    def test_find_workbook_not_found(self):
        """Test returning None when workbook is not found."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\other.xlsx"
        mock_wb.Name = "other.xlsx"
        mock_app.Workbooks = [mock_wb]

        # Test
        path = Path("C:/data/missing.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Verify
        assert result is None

    def test_find_workbook_case_insensitive(self):
        """Test that search is case-insensitive."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.FullName = "C:\\DATA\\TEST.XLSX"
        mock_wb.Name = "TEST.XLSX"
        mock_app.Workbooks = [mock_wb]

        # Test with different case
        path = Path("c:/data/test.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Should find it
        assert result == mock_wb

    def test_find_workbook_multiple_workbooks(self):
        """Test finding correct workbook among multiple open workbooks."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mocks
        mock_app = Mock()

        mock_wb1 = Mock()
        mock_wb1.FullName = "C:\\data\\file1.xlsx"
        mock_wb1.Name = "file1.xlsx"

        mock_wb2 = Mock()
        mock_wb2.FullName = "C:\\data\\file2.xlsx"
        mock_wb2.Name = "file2.xlsx"

        mock_wb3 = Mock()
        mock_wb3.FullName = "C:\\data\\file3.xlsx"
        mock_wb3.Name = "file3.xlsx"

        mock_app.Workbooks = [mock_wb1, mock_wb2, mock_wb3]

        # Test finding the middle one
        path = Path("C:/data/file2.xlsx")
        result = _find_open_workbook(mock_app, path)

        assert result == mock_wb2

    def test_find_workbook_with_exception(self):
        """Test handling exception when accessing workbook properties."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock with one broken workbook
        mock_app = Mock()

        mock_wb1 = Mock()
        # This workbook raises exception when accessing FullName
        type(mock_wb1).FullName = Mock(side_effect=Exception("Access denied"))

        mock_wb2 = Mock()
        mock_wb2.FullName = "C:\\data\\good.xlsx"
        mock_wb2.Name = "good.xlsx"

        mock_app.Workbooks = [mock_wb1, mock_wb2]

        # Should skip the broken workbook and find the good one
        path = Path("C:/data/good.xlsx")
        result = _find_open_workbook(mock_app, path)

        assert result == mock_wb2

    def test_find_workbook_empty_collection(self):
        """Test with no open workbooks."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_app.Workbooks = []

        # Test
        path = Path("C:/data/any.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Should return None
        assert result is None

    def test_find_workbook_relative_path(self):
        """Test with relative path (should be resolved to absolute)."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_wb = Mock()
        # Absolute path in Excel
        mock_wb.FullName = str(Path("data/test.xlsx").resolve())
        mock_wb.Name = "test.xlsx"
        mock_app.Workbooks = [mock_wb]

        # Test with relative path
        path = Path("data/test.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Should resolve and find it
        assert result == mock_wb


class TestWorkbookManager:
    """Tests for WorkbookManager class."""

    def test_workbook_manager_initialization(self):
        """Test WorkbookManager initialization."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.excel_manager import ExcelManager

        mock_mgr = Mock(spec=ExcelManager)
        wb_mgr = WorkbookManager(mock_mgr)

        assert wb_mgr._mgr == mock_mgr


class TestWorkbookManagerOpen:
    """Tests for WorkbookManager.open() method."""

    def test_open_success(self, tmp_path):
        """Test successfully opening a workbook."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.excel_manager import ExcelManager

        # Create a temporary file (to pass exists() check)
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        # Setup mocks
        mock_excel_mgr = Mock(spec=ExcelManager)
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock Workbooks collection
        mock_workbooks = Mock()
        mock_app.Workbooks = mock_workbooks

        # No existing workbooks - make it iterable
        mock_workbooks.__iter__ = Mock(return_value=iter([]))

        # Mock the opened workbook
        mock_wb = Mock()
        mock_wb.Name = "test.xlsx"
        mock_wb.FullName = str(test_file)
        mock_wb.ReadOnly = False
        mock_wb.Saved = True
        mock_wb.Worksheets.Count = 3
        mock_workbooks.Open.return_value = mock_wb

        # Test
        wb_mgr = WorkbookManager(mock_excel_mgr)
        info = wb_mgr.open(test_file)

        # Verify
        assert info.name == "test.xlsx"
        assert info.full_path == test_file
        assert info.read_only is False
        assert info.saved is True
        assert info.sheets_count == 3

        # Verify COM call
        mock_workbooks.Open.assert_called_once()
        call_args = mock_workbooks.Open.call_args
        # Check that the first argument is the file path
        assert call_args.args[0] == str(test_file)

    def test_open_read_only(self, tmp_path):
        """Test opening workbook in read-only mode."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.excel_manager import ExcelManager

        test_file = tmp_path / "readonly.xlsx"
        test_file.touch()

        mock_excel_mgr = Mock(spec=ExcelManager)
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock Workbooks collection
        mock_workbooks = Mock()
        mock_app.Workbooks = mock_workbooks
        mock_workbooks.__iter__ = Mock(return_value=iter([]))

        mock_wb = Mock()
        mock_wb.Name = "readonly.xlsx"
        mock_wb.FullName = str(test_file)
        mock_wb.ReadOnly = True  # Excel confirmed read-only
        mock_wb.Saved = True
        mock_wb.Worksheets.Count = 1
        mock_workbooks.Open.return_value = mock_wb

        wb_mgr = WorkbookManager(mock_excel_mgr)
        info = wb_mgr.open(test_file, read_only=True)

        assert info.read_only is True

        # Verify ReadOnly parameter was passed
        call_args = mock_workbooks.Open.call_args
        assert call_args.kwargs.get("ReadOnly") is True

    def test_open_file_not_found(self):
        """Test opening non-existent file."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookNotFoundError

        mock_excel_mgr = Mock()
        wb_mgr = WorkbookManager(mock_excel_mgr)

        missing_file = Path("C:/nonexistent/missing.xlsx")

        with pytest.raises(WorkbookNotFoundError) as exc_info:
            wb_mgr.open(missing_file)

        assert exc_info.value.path == missing_file
        assert "not found" in str(exc_info.value).lower()

    def test_open_already_open(self, tmp_path):
        """Test opening a workbook that's already open."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookAlreadyOpenError

        test_file = tmp_path / "already_open.xlsx"
        test_file.touch()

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock already open workbook
        mock_existing_wb = Mock()
        mock_existing_wb.FullName = str(test_file)
        mock_existing_wb.Name = "already_open.xlsx"
        mock_app.Workbooks = [mock_existing_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(WorkbookAlreadyOpenError) as exc_info:
            wb_mgr.open(test_file)

        assert exc_info.value.path == test_file
        assert exc_info.value.name == "already_open.xlsx"
        assert "already open" in str(exc_info.value).lower()

    def test_open_com_error(self, tmp_path):
        """Test handling COM error during open."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import ExcelConnectionError

        test_file = tmp_path / "error.xlsx"
        test_file.touch()

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Mock Workbooks collection
        mock_workbooks = Mock()
        mock_app.Workbooks = mock_workbooks
        mock_workbooks.__iter__ = Mock(return_value=iter([]))

        # Mock COM error
        com_error = Exception("File is corrupted")
        com_error.hresult = 0x800A03EC
        mock_workbooks.Open.side_effect = com_error

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(ExcelConnectionError) as exc_info:
            wb_mgr.open(test_file)

        assert exc_info.value.hresult == 0x800A03EC
        assert "Failed to open workbook" in str(exc_info.value)

    def test_open_excel_not_started(self, tmp_path):
        """Test opening when Excel is not started."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import ExcelConnectionError

        # Create the file first to pass the exists() check
        test_file = tmp_path / "any.xlsx"
        test_file.touch()

        # Create a mock ExcelManager where the app property raises an exception
        class MockExcelManager:
            def __init__(self):
                pass

            @property
            def app(self):
                raise ExcelConnectionError(0x80080005, "Excel application not started")

        mock_excel_mgr = MockExcelManager()

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(ExcelConnectionError):
            wb_mgr.open(test_file)


class TestWorkbookManagerCreate:
    """Tests for WorkbookManager.create() method."""

    def test_create_blank_workbook(self, tmp_path):
        """Test creating a blank workbook."""
        from xlmanage.workbook_manager import WorkbookManager

        # Setup
        test_file = tmp_path / "new.xlsx"
        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.Name = "new.xlsx"
        mock_wb.FullName = str(test_file)
        mock_wb.ReadOnly = False
        mock_wb.Saved = True
        mock_wb.Worksheets.Count = 1
        mock_app.Workbooks.Add.return_value = mock_wb

        # Test
        wb_mgr = WorkbookManager(mock_excel_mgr)
        info = wb_mgr.create(test_file)

        # Verify
        assert info.name == "new.xlsx"
        assert info.sheets_count == 1

        # Verify COM calls
        mock_app.Workbooks.Add.assert_called_once_with()  # No template
        mock_wb.SaveAs.assert_called_once()

        # Verify SaveAs parameters
        save_call = mock_wb.SaveAs.call_args
        assert str(test_file) in str(save_call[0][0])  # First arg is path
        assert save_call.kwargs.get("FileFormat") == 51  # .xlsx format

    def test_create_with_template(self, tmp_path):
        """Test creating workbook from template."""
        from xlmanage.workbook_manager import WorkbookManager

        # Create template file
        template_file = tmp_path / "template.xltx"
        template_file.touch()

        test_file = tmp_path / "from_template.xlsx"

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.Name = "from_template.xlsx"
        mock_wb.FullName = str(test_file)
        mock_wb.ReadOnly = False
        mock_wb.Saved = True
        mock_wb.Worksheets.Count = 3
        mock_app.Workbooks.Add.return_value = mock_wb

        wb_mgr = WorkbookManager(mock_excel_mgr)
        info = wb_mgr.create(test_file, template=template_file)

        # Verify template was used
        mock_app.Workbooks.Add.assert_called_once()
        call_arg = mock_app.Workbooks.Add.call_args[0][0]
        assert str(template_file) in call_arg

    def test_create_xlsm_format(self, tmp_path):
        """Test creating macro-enabled workbook."""
        from xlmanage.workbook_manager import WorkbookManager

        test_file = tmp_path / "macros.xlsm"

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.Name = "macros.xlsm"
        mock_wb.FullName = str(test_file)
        mock_wb.ReadOnly = False
        mock_wb.Saved = True
        mock_wb.Worksheets.Count = 1
        mock_app.Workbooks.Add.return_value = mock_wb

        wb_mgr = WorkbookManager(mock_excel_mgr)
        info = wb_mgr.create(test_file)

        # Verify .xlsm format (52)
        save_call = mock_wb.SaveAs.call_args
        assert save_call.kwargs.get("FileFormat") == 52

    def test_create_xls_legacy_format(self, tmp_path):
        """Test creating legacy Excel 97-2003 workbook."""
        from xlmanage.workbook_manager import WorkbookManager

        test_file = tmp_path / "legacy.xls"

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.Name = "legacy.xls"
        mock_wb.FullName = str(test_file)
        mock_wb.ReadOnly = False
        mock_wb.Saved = True
        mock_wb.Worksheets.Count = 1
        mock_app.Workbooks.Add.return_value = mock_wb

        wb_mgr = WorkbookManager(mock_excel_mgr)
        info = wb_mgr.create(test_file)

        # Verify .xls format (56)
        save_call = mock_wb.SaveAs.call_args
        assert save_call.kwargs.get("FileFormat") == 56

    def test_create_template_not_found(self, tmp_path):
        """Test creating with non-existent template."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookNotFoundError

        missing_template = tmp_path / "missing_template.xltx"
        test_file = tmp_path / "new.xlsx"

        mock_excel_mgr = Mock()
        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(WorkbookNotFoundError) as exc_info:
            wb_mgr.create(test_file, template=missing_template)

        assert exc_info.value.path == missing_template
        assert "template" in str(exc_info.value).lower()

    def test_create_invalid_extension(self, tmp_path):
        """Test creating with invalid file extension."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookSaveError

        test_file = tmp_path / "invalid.txt"

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(WorkbookSaveError) as exc_info:
            wb_mgr.create(test_file)

        assert exc_info.value.path == test_file
        assert "extension" in str(exc_info.value).lower()

    def test_create_save_fails(self, tmp_path):
        """Test handling SaveAs failure."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookSaveError

        test_file = tmp_path / "save_fails.xlsx"

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        # SaveAs raises error
        save_error = Exception("Access denied")
        save_error.hresult = 0x80070005
        mock_wb.SaveAs.side_effect = save_error
        mock_app.Workbooks.Add.return_value = mock_wb

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(WorkbookSaveError) as exc_info:
            wb_mgr.create(test_file)

        assert exc_info.value.hresult == 0x80070005
        assert exc_info.value.path == test_file

        # Verify cleanup: workbook was closed without saving
        mock_wb.Close.assert_called_once_with(SaveChanges=False)

    def test_create_com_error(self, tmp_path):
        """Test handling COM error during creation."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import ExcelConnectionError

        test_file = tmp_path / "error.xlsx"

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        # Workbooks.Add raises COM error
        com_error = Exception("Excel is busy")
        com_error.hresult = 0x80080005
        mock_app.Workbooks.Add.side_effect = com_error

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(ExcelConnectionError) as exc_info:
            wb_mgr.create(test_file)

        assert exc_info.value.hresult == 0x80080005

    def test_create_cleanup_fails_silently(self, tmp_path):
        """Test that cleanup failure doesn't mask save error."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookSaveError

        test_file = tmp_path / "cleanup_fails.xlsx"

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        # SaveAs fails
        mock_wb.SaveAs.side_effect = Exception("Save failed")
        # Close also fails
        mock_wb.Close.side_effect = Exception("Close failed")
        mock_app.Workbooks.Add.return_value = mock_wb

        wb_mgr = WorkbookManager(mock_excel_mgr)

        # Should raise SaveError, not Close error
        with pytest.raises(WorkbookSaveError) as exc_info:
            wb_mgr.create(test_file)

        assert "save" in str(exc_info.value).lower()


class TestWorkbookManagerClose:
    """Tests for WorkbookManager.close() method."""

    def test_close_with_save(self):
        """Test closing workbook with save."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\test.xlsx"
        mock_wb.Name = "test.xlsx"
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        wb_mgr.close(Path("C:/data/test.xlsx"), save=True)

        mock_wb.Close.assert_called_once_with(SaveChanges=True)

    def test_close_without_save(self):
        """Test closing workbook without save."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\temp\\discard.xlsx"
        mock_wb.Name = "discard.xlsx"
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        wb_mgr.close(Path("C:/temp/discard.xlsx"), save=False)

        mock_wb.Close.assert_called_once_with(SaveChanges=False)

    def test_close_with_force(self):
        """Test force close suppresses alerts."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\force.xlsx"
        mock_wb.Name = "force.xlsx"
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        wb_mgr.close(Path("C:/data/force.xlsx"), force=True)

        # Verify DisplayAlerts was disabled then restored
        # Since DisplayAlerts is a property setter, we need to check the mock calls differently
        # The mock should have recorded the property assignments
        assert hasattr(mock_app, "DisplayAlerts")
        # The DisplayAlerts should have been set to True in the finally block
        assert mock_app.DisplayAlerts is True

    def test_close_workbook_not_open(self):
        """Test closing workbook that is not open."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookNotFoundError

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app
        mock_app.Workbooks = []  # No open workbooks

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(WorkbookNotFoundError):
            wb_mgr.close(Path("C:/data/notopen.xlsx"))

    def test_close_restores_alerts_on_error(self):
        """Test DisplayAlerts is restored even if Close fails."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\error.xlsx"
        mock_wb.Name = "error.xlsx"
        mock_wb.Close.side_effect = Exception("Close failed")
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(Exception):
            wb_mgr.close(Path("C:/data/error.xlsx"), force=True)

        # DisplayAlerts should be restored to True in finally
        assert mock_app.DisplayAlerts is True


class TestWorkbookManagerSave:
    """Tests for WorkbookManager.save() method."""

    def test_save_to_current_file(self):
        """Test saving to current file."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\work.xlsx"
        mock_wb.Name = "work.xlsx"
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        wb_mgr.save(Path("C:/data/work.xlsx"))

        # Should call Save(), not SaveAs()
        mock_wb.Save.assert_called_once()
        mock_wb.SaveAs.assert_not_called()

    def test_save_as_to_different_file(self):
        """Test SaveAs to different file."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\original.xlsx"
        mock_wb.Name = "original.xlsx"
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        output = Path("C:/backup/copy.xlsx")
        wb_mgr.save(Path("C:/data/original.xlsx"), output=output)

        # Should call SaveAs()
        mock_wb.SaveAs.assert_called_once()
        call_args = mock_wb.SaveAs.call_args
        assert str(output) in str(call_args[0][0])
        assert call_args.kwargs.get("FileFormat") == 51  # .xlsx

    def test_save_as_different_format(self):
        """Test SaveAs with format conversion."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\data.xlsx"
        mock_wb.Name = "data.xlsx"
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        output = Path("C:/archive/data.xlsb")  # Binary format
        wb_mgr.save(Path("C:/data/data.xlsx"), output=output)

        call_args = mock_wb.SaveAs.call_args
        assert call_args.kwargs.get("FileFormat") == 50  # .xlsb

    def test_save_workbook_not_open(self):
        """Test saving workbook that is not open."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookNotFoundError

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app
        mock_app.Workbooks = []

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(WorkbookNotFoundError):
            wb_mgr.save(Path("C:/data/notopen.xlsx"))

    def test_save_as_invalid_extension(self):
        """Test SaveAs with invalid extension."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookSaveError

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\work.xlsx"
        mock_wb.Name = "work.xlsx"
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(WorkbookSaveError) as exc_info:
            wb_mgr.save(Path("C:/data/work.xlsx"), output=Path("C:/data/work.txt"))

        assert "extension" in str(exc_info.value).lower()

    def test_save_com_error(self):
        """Test handling COM error during save."""
        from xlmanage.workbook_manager import WorkbookManager
        from xlmanage.exceptions import WorkbookSaveError

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\readonly.xlsx"
        mock_wb.Name = "readonly.xlsx"
        save_error = Exception("Access denied")
        save_error.hresult = 0x80070005
        mock_wb.Save.side_effect = save_error
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)

        with pytest.raises(WorkbookSaveError) as exc_info:
            wb_mgr.save(Path("C:/data/readonly.xlsx"))

        assert exc_info.value.hresult == 0x80070005


class TestWorkbookManagerList:
    """Tests for WorkbookManager.list() method."""

    def test_list_no_workbooks(self):
        """Test listing when no workbooks are open."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app
        mock_app.Workbooks = []

        wb_mgr = WorkbookManager(mock_excel_mgr)
        workbooks = wb_mgr.list()

        assert workbooks == []

    def test_list_single_workbook(self):
        """Test listing single workbook."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb = Mock()
        mock_wb.Name = "test.xlsx"
        mock_wb.FullName = "C:\\data\\test.xlsx"
        mock_wb.ReadOnly = False
        mock_wb.Saved = True
        mock_wb.Worksheets.Count = 3
        mock_app.Workbooks = [mock_wb]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        workbooks = wb_mgr.list()

        assert len(workbooks) == 1
        assert workbooks[0].name == "test.xlsx"
        assert workbooks[0].sheets_count == 3

    def test_list_multiple_workbooks(self):
        """Test listing multiple workbooks."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb1 = Mock()
        mock_wb1.Name = "file1.xlsx"
        mock_wb1.FullName = "C:\\data\\file1.xlsx"
        mock_wb1.ReadOnly = False
        mock_wb1.Saved = True
        mock_wb1.Worksheets.Count = 2

        mock_wb2 = Mock()
        mock_wb2.Name = "file2.xlsm"
        mock_wb2.FullName = "C:\\data\\file2.xlsm"
        mock_wb2.ReadOnly = True
        mock_wb2.Saved = False
        mock_wb2.Worksheets.Count = 5

        mock_app.Workbooks = [mock_wb1, mock_wb2]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        workbooks = wb_mgr.list()

        assert len(workbooks) == 2
        assert workbooks[0].name == "file1.xlsx"
        assert workbooks[1].name == "file2.xlsm"
        assert workbooks[1].read_only is True

    def test_list_with_error_workbook(self):
        """Test listing continues when one workbook raises error."""
        from xlmanage.workbook_manager import WorkbookManager

        mock_excel_mgr = Mock()
        mock_app = Mock()
        mock_excel_mgr.app = mock_app

        mock_wb1 = Mock()
        mock_wb1.Name = "good.xlsx"
        mock_wb1.FullName = "C:\\data\\good.xlsx"
        mock_wb1.ReadOnly = False
        mock_wb1.Saved = True
        mock_wb1.Worksheets.Count = 1

        mock_wb2 = Mock()
        type(mock_wb2).Name = Mock(side_effect=Exception("Corrupted"))

        mock_wb3 = Mock()
        mock_wb3.Name = "good2.xlsx"
        mock_wb3.FullName = "C:\\data\\good2.xlsx"
        mock_wb3.ReadOnly = False
        mock_wb3.Saved = True
        mock_wb3.Worksheets.Count = 2

        mock_app.Workbooks = [mock_wb1, mock_wb2, mock_wb3]

        wb_mgr = WorkbookManager(mock_excel_mgr)
        workbooks = wb_mgr.list()

        assert len(workbooks) == 2
        assert workbooks[0].name == "good.xlsx"
        assert workbooks[1].name == "good2.xlsx"
