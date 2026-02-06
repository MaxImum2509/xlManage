"""Tests for VBAManager delete_module functionality."""

import pytest
from unittest.mock import Mock

import pywintypes

from xlmanage.excel_manager import ExcelManager
from xlmanage.exceptions import VBAModuleNotFoundError
from xlmanage.vba_manager import VBAManager


@pytest.fixture
def mock_excel_manager():
    """Create a mock ExcelManager."""
    mock_mgr = Mock(spec=ExcelManager)
    mock_app = Mock()
    mock_mgr.app = mock_app
    return mock_mgr


def test_delete_standard_module_success(mock_excel_manager):
    """Test successful deletion of standard module."""
    # Mock du composant
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.Type = 1  # VBEXT_CT_STD_MODULE

    # Mock du VBProject
    mock_vb_components = Mock()
    mock_vb_components.__iter__ = Mock(return_value=iter([mock_component]))
    mock_vb_components.Remove = Mock()

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mock_vb_components

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    vba_mgr.delete_module("Module1")

    # Vérifier que Remove a été appelé
    mock_vb_components.Remove.assert_called_once_with(mock_component)


def test_delete_class_module_success(mock_excel_manager):
    """Test successful deletion of class module."""
    mock_component = Mock()
    mock_component.Name = "MyClass"
    mock_component.Type = 2  # VBEXT_CT_CLASS_MODULE

    mock_vb_components = Mock()
    mock_vb_components.__iter__ = Mock(return_value=iter([mock_component]))
    mock_vb_components.Remove = Mock()

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mock_vb_components

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    vba_mgr.delete_module("MyClass")

    mock_vb_components.Remove.assert_called_once()


def test_delete_userform_success(mock_excel_manager):
    """Test successful deletion of UserForm."""
    mock_component = Mock()
    mock_component.Name = "UserForm1"
    mock_component.Type = 3  # VBEXT_CT_MS_FORM

    mock_vb_components = Mock()
    mock_vb_components.__iter__ = Mock(return_value=iter([mock_component]))
    mock_vb_components.Remove = Mock()

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mock_vb_components

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    vba_mgr.delete_module("UserForm1")

    mock_vb_components.Remove.assert_called_once()


def test_delete_document_module_error(mock_excel_manager):
    """Test error when trying to delete a document module."""
    # Mock d'un module de document (non supprimable)
    mock_component = Mock()
    mock_component.Name = "ThisWorkbook"
    mock_component.Type = 100  # VBEXT_CT_DOCUMENT

    mock_vb_components = Mock()
    mock_vb_components.__iter__ = Mock(return_value=iter([mock_component]))
    mock_vb_components.Remove = Mock()

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mock_vb_components

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError) as exc_info:
        vba_mgr.delete_module("ThisWorkbook")

    # Vérifier le message d'erreur
    assert "Cannot delete document module" in str(exc_info.value)
    assert "ThisWorkbook" in str(exc_info.value)
    # Vérifier que Remove n'a PAS été appelé
    mock_vb_components.Remove.assert_not_called()


def test_delete_sheet_module_error(mock_excel_manager):
    """Test error when trying to delete a Sheet document module."""
    mock_component = Mock()
    mock_component.Name = "Sheet1"
    mock_component.Type = 100  # VBEXT_CT_DOCUMENT

    mock_vb_components = Mock()
    mock_vb_components.__iter__ = Mock(return_value=iter([mock_component]))

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mock_vb_components

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError) as exc_info:
        vba_mgr.delete_module("Sheet1")

    assert "Cannot delete document module" in str(exc_info.value)
    assert "document" in str(exc_info.value).lower()


def test_delete_module_not_found(mock_excel_manager):
    """Test error when module doesn't exist."""
    mock_vb_components = Mock()
    mock_vb_components.__iter__ = Mock(return_value=iter([]))

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mock_vb_components

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError) as exc_info:
        vba_mgr.delete_module("NonExistent")

    assert "not found" in str(exc_info.value)
    assert "NonExistent" in str(exc_info.value)


def test_delete_module_com_error(mock_excel_manager):
    """Test error handling when COM fails during deletion."""
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.Type = 1

    # Simuler une erreur COM lors de Remove
    com_error = pywintypes.com_error(-2147352567, "Access denied", None, None)

    mock_vb_components = Mock()
    mock_vb_components.__iter__ = Mock(return_value=iter([mock_component]))
    mock_vb_components.Remove = Mock(side_effect=com_error)

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mock_vb_components

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError):
        vba_mgr.delete_module("Module1")


def test_delete_module_with_force_parameter(mock_excel_manager):
    """Test delete with force parameter (currently unused)."""
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.Type = 1

    mock_vb_components = Mock()
    mock_vb_components.__iter__ = Mock(return_value=iter([mock_component]))
    mock_vb_components.Remove = Mock()

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mock_vb_components

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    # Force parameter est accepté mais n'a pas d'effet actuellement
    vba_mgr.delete_module("Module1", force=True)

    mock_vb_components.Remove.assert_called_once()
