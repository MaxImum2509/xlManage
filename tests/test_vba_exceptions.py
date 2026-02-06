"""Tests for VBA exceptions."""

import pytest

from xlmanage.exceptions import (
    VBAExportError,
    VBAImportError,
    VBAMacroError,
    VBAModuleAlreadyExistsError,
    VBAModuleNotFoundError,
    VBAProjectAccessError,
    VBAWorkbookFormatError,
)


def test_vba_project_access_error():
    """Test VBAProjectAccessError attributes and message."""
    error = VBAProjectAccessError("test.xlsm")
    assert error.workbook_name == "test.xlsm"
    assert "Trust access" in str(error)


def test_vba_module_not_found_error():
    """Test VBAModuleNotFoundError attributes and message."""
    error = VBAModuleNotFoundError("Module1", "test.xlsm")
    assert error.module_name == "Module1"
    assert error.workbook_name == "test.xlsm"
    assert "not found" in str(error)


def test_vba_module_already_exists_error():
    """Test VBAModuleAlreadyExistsError attributes and message."""
    error = VBAModuleAlreadyExistsError("Module1", "test.xlsm")
    assert error.module_name == "Module1"
    assert error.workbook_name == "test.xlsm"
    assert "already exists" in str(error)


def test_vba_import_error():
    """Test VBAImportError attributes and message."""
    error = VBAImportError("Module1.bas", "Invalid encoding")
    assert error.module_file == "Module1.bas"
    assert error.reason == "Invalid encoding"
    assert "Failed to import" in str(error)


def test_vba_export_error():
    """Test VBAExportError attributes and message."""
    error = VBAExportError("Module1", "C:\\output.bas", "Permission denied")
    assert error.module_name == "Module1"
    assert error.output_path == "C:\\output.bas"
    assert error.reason == "Permission denied"
    assert "Failed to export" in str(error)


def test_vba_macro_error():
    """Test VBAMacroError attributes and message."""
    error = VBAMacroError("MySub", "Runtime error '9': Subscript out of range")
    assert error.macro_name == "MySub"
    assert error.reason == "Runtime error '9': Subscript out of range"
    assert "failed" in str(error)


def test_vba_workbook_format_error():
    """Test VBAWorkbookFormatError attributes and message."""
    error = VBAWorkbookFormatError("data.xlsx")
    assert error.workbook_name == "data.xlsx"
    assert ".xlsx format" in str(error)
    assert ".xlsm" in str(error)
