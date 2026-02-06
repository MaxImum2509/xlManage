"""Tests for VBAManager export and list functionality."""

import pytest
from pathlib import Path
from unittest.mock import Mock

import pywintypes

from xlmanage.excel_manager import ExcelManager
from xlmanage.exceptions import VBAExportError, VBAModuleNotFoundError
from xlmanage.vba_manager import VBAManager, VBAModuleInfo


@pytest.fixture
def mock_excel_manager():
    """Create a mock ExcelManager."""
    mock_mgr = Mock(spec=ExcelManager)
    mock_app = Mock()
    mock_mgr.app = mock_app
    return mock_mgr


def test_export_standard_module_success(mock_excel_manager, tmp_path):
    """Test successful export of standard module."""
    output_file = tmp_path / "Module1.bas"

    # Mock du composant
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.Type = 1  # VBEXT_CT_STD_MODULE
    mock_component.Export = Mock()

    # Mock du VBProject
    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.export_module("Module1", output_file)

    assert result == output_file
    mock_component.Export.assert_called_once()


def test_export_class_module_success(mock_excel_manager, tmp_path):
    """Test successful export of class module."""
    output_file = tmp_path / "MyClass.cls"

    # Mock du composant
    mock_component = Mock()
    mock_component.Name = "MyClass"
    mock_component.Type = 2  # VBEXT_CT_CLASS_MODULE
    mock_component.Export = Mock()

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.export_module("MyClass", output_file)

    assert result == output_file
    mock_component.Export.assert_called_once()


def test_export_userform_success(mock_excel_manager, tmp_path):
    """Test successful export of UserForm."""
    output_file = tmp_path / "UserForm1.frm"

    # Mock du composant
    mock_component = Mock()
    mock_component.Name = "UserForm1"
    mock_component.Type = 3  # VBEXT_CT_MS_FORM
    mock_component.Export = Mock()

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.export_module("UserForm1", output_file)

    assert result == output_file
    mock_component.Export.assert_called_once()


def test_export_document_module_success(mock_excel_manager, tmp_path):
    """Test export of document module with manual extraction."""
    output_file = tmp_path / "ThisWorkbook.cls"

    # Mock du code module
    mock_code_module = Mock()
    mock_code_module.CountOfLines = 5
    mock_code_module.Lines.return_value = "Option Explicit\n\nSub Test()\nEnd Sub"

    # Mock du composant document
    mock_component = Mock()
    mock_component.Name = "ThisWorkbook"
    mock_component.Type = 100  # VBEXT_CT_DOCUMENT
    mock_component.CodeModule = mock_code_module

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.export_module("ThisWorkbook", output_file)

    assert result == output_file
    assert output_file.exists()
    mock_code_module.Lines.assert_called_once_with(1, 5)

    # Vérifier le contenu du fichier
    content = output_file.read_text(encoding="windows-1252")
    assert "Option Explicit" in content
    assert "Sub Test()" in content


def test_export_document_module_empty(mock_excel_manager, tmp_path):
    """Test export of document module with no code."""
    output_file = tmp_path / "Sheet1.cls"

    # Mock du code module vide
    mock_code_module = Mock()
    mock_code_module.CountOfLines = 0

    mock_component = Mock()
    mock_component.Name = "Sheet1"
    mock_component.Type = 100  # VBEXT_CT_DOCUMENT
    mock_component.CodeModule = mock_code_module

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.export_module("Sheet1", output_file)

    assert result == output_file
    assert output_file.exists()
    # Vérifier que le fichier est vide
    assert output_file.read_text(encoding="windows-1252") == ""


def test_export_module_not_found(mock_excel_manager):
    """Test error when module doesn't exist."""
    mock_vb_project = Mock()
    mock_vb_project.VBComponents = []

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError):
        vba_mgr.export_module("NonExistent", Path("output.bas"))


def test_export_module_permission_error(mock_excel_manager, tmp_path):
    """Test error when permissions are insufficient."""
    output_file = tmp_path / "readonly" / "Module1.bas"

    # Mock du composant
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.Type = 1
    # Simuler une erreur de permission
    mock_component.Export.side_effect = PermissionError("Access denied")

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAExportError, match="Permission refusée"):
        vba_mgr.export_module("Module1", output_file)


def test_list_modules_success(mock_excel_manager):
    """Test listing all VBA modules."""
    # Mock de plusieurs composants
    mock_comp1 = Mock()
    mock_comp1.Name = "Module1"
    mock_comp1.Type = 1  # standard
    mock_comp1.CodeModule.CountOfLines = 42

    mock_comp2 = Mock()
    mock_comp2.Name = "MyClass"
    mock_comp2.Type = 2  # class
    mock_comp2.CodeModule.CountOfLines = 15
    mock_comp2.Properties.return_value.Value = True

    mock_comp3 = Mock()
    mock_comp3.Name = "ThisWorkbook"
    mock_comp3.Type = 100  # document
    mock_comp3.CodeModule.CountOfLines = 8

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_comp1, mock_comp2, mock_comp3]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    modules = vba_mgr.list_modules()

    assert len(modules) == 3
    assert modules[0].name == "Module1"
    assert modules[0].module_type == "standard"
    assert modules[0].lines_count == 42
    assert modules[0].has_predeclared_id is False

    assert modules[1].name == "MyClass"
    assert modules[1].module_type == "class"
    assert modules[1].has_predeclared_id is True

    assert modules[2].name == "ThisWorkbook"
    assert modules[2].module_type == "document"


def test_list_modules_class_without_predeclared_id(mock_excel_manager):
    """Test listing class module when PredeclaredId access fails."""
    mock_comp = Mock()
    mock_comp.Name = "MyClass"
    mock_comp.Type = 2  # class
    mock_comp.CodeModule.CountOfLines = 10
    # Simuler une erreur COM lors de l'accès à PredeclaredId
    mock_comp.Properties.side_effect = pywintypes.com_error(
        -2147352567, "Property not found", None, None
    )

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_comp]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    modules = vba_mgr.list_modules()

    assert len(modules) == 1
    assert modules[0].name == "MyClass"
    assert modules[0].has_predeclared_id is False  # Valeur par défaut


def test_list_modules_empty(mock_excel_manager):
    """Test listing modules when project is empty."""
    mock_vb_project = Mock()
    mock_vb_project.VBComponents = []

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    modules = vba_mgr.list_modules()

    assert modules == []


def test_list_modules_all_types(mock_excel_manager):
    """Test listing modules of all different types."""
    # Un de chaque type
    mocks = []
    types_info = [
        ("Module1", 1, "standard"),
        ("MyClass", 2, "class"),
        ("UserForm1", 3, "userform"),
        ("ThisWorkbook", 100, "document"),
    ]

    for name, type_code, _ in types_info:
        mock = Mock()
        mock.Name = name
        mock.Type = type_code
        mock.CodeModule.CountOfLines = 10
        if type_code == 2:  # class
            mock.Properties.return_value.Value = False
        mocks.append(mock)

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = mocks

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    modules = vba_mgr.list_modules()

    assert len(modules) == 4
    for i, (name, _, expected_type) in enumerate(types_info):
        assert modules[i].name == name
        assert modules[i].module_type == expected_type
