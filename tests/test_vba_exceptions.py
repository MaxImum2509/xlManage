"""Tests for VBA-specific exceptions."""

import pytest

from xlmanage.exceptions import VBAModuleNotFoundError


def test_vba_module_not_found_error_with_reason():
    """Test VBAModuleNotFoundError with custom reason."""
    error = VBAModuleNotFoundError(
        "ThisWorkbook", "test.xlsm", reason="Cannot delete document module"
    )

    assert error.module_name == "ThisWorkbook"
    assert error.workbook_name == "test.xlsm"
    assert error.reason == "Cannot delete document module"
    assert "Cannot delete document module" in str(error)
    assert "ThisWorkbook" in str(error)
    assert "test.xlsm" in str(error)


def test_vba_module_not_found_error_without_reason():
    """Test VBAModuleNotFoundError without custom reason."""
    error = VBAModuleNotFoundError("Module1", "test.xlsm")

    assert error.module_name == "Module1"
    assert error.workbook_name == "test.xlsm"
    assert error.reason == ""
    assert "not found" in str(error)
    assert "Module1" in str(error)


def test_vba_module_not_found_error_empty_reason():
    """Test VBAModuleNotFoundError with empty string reason."""
    error = VBAModuleNotFoundError("Module1", "test.xlsm", reason="")

    assert error.reason == ""
    assert "not found" in str(error)
    # Vérifier que le format par défaut est utilisé
    assert "VBA module 'Module1' not found" in str(error)
