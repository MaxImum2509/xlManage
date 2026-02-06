"""Tests for VBAManager import_module functionality."""

import pytest
from pathlib import Path
from unittest.mock import Mock, PropertyMock

import pywintypes

from xlmanage.excel_manager import ExcelManager
from xlmanage.exceptions import (
    VBAImportError,
    VBAModuleAlreadyExistsError,
    VBAProjectAccessError,
)
from xlmanage.vba_manager import VBAManager, VBAModuleInfo


@pytest.fixture
def mock_excel_manager():
    """Create a mock ExcelManager."""
    mock_mgr = Mock(spec=ExcelManager)
    mock_app = Mock()
    mock_mgr.app = mock_app
    return mock_mgr


def test_import_module_file_not_found(mock_excel_manager):
    """Test error when module file doesn't exist."""
    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAImportError, match="introuvable"):
        vba_mgr.import_module(Path("nonexistent.bas"))


def test_import_standard_module_success(mock_excel_manager, tmp_path):
    """Test successful import of .bas module."""
    # Créer un fichier .bas
    bas_file = tmp_path / "Module1.bas"
    bas_content = '''Attribute VB_Name = "Module1"
Sub Hello()
    MsgBox "Hello"
End Sub
'''
    bas_file.write_text(bas_content, encoding="windows-1252")

    # Mock du VBProject
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.CodeModule.CountOfLines = 4

    mock_vb_project = Mock()
    mock_vb_project.Name = "test.xlsm"
    mock_vb_project.VBComponents.Import.return_value = mock_component
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.import_module(bas_file)

    assert result.name == "Module1"
    assert result.module_type == "standard"
    assert result.lines_count == 4
    assert result.has_predeclared_id is False
    mock_vb_project.VBComponents.Import.assert_called_once()


def test_import_class_module_success(mock_excel_manager, tmp_path):
    """Test successful import of .cls module."""
    cls_file = tmp_path / "MyClass.cls"
    cls_content = '''VERSION 1.0 CLASS
Attribute VB_Name = "MyClass"
Attribute VB_PredeclaredId = True
Option Explicit

Public Sub Test()
End Sub
'''
    cls_file.write_text(cls_content, encoding="windows-1252")

    # Mock du VBProject
    mock_component = Mock()
    mock_component.Name = "MyClass"
    mock_component.CodeModule.CountOfLines = 3
    mock_component.Properties = Mock(return_value=Mock())

    mock_vb_project = Mock()
    mock_vb_project.Name = "test.xlsm"
    mock_vb_project.VBComponents.Add.return_value = mock_component
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.import_module(cls_file)

    assert result.name == "MyClass"
    assert result.module_type == "class"
    assert result.has_predeclared_id is True
    mock_vb_project.VBComponents.Add.assert_called_once_with(2)  # VBEXT_CT_CLASS_MODULE


def test_import_class_module_without_predeclared_id(mock_excel_manager, tmp_path):
    """Test import of .cls module without PredeclaredId."""
    cls_file = tmp_path / "SimpleClass.cls"
    cls_content = '''VERSION 1.0 CLASS
Attribute VB_Name = "SimpleClass"
Option Explicit

Public Sub Test()
End Sub
'''
    cls_file.write_text(cls_content, encoding="windows-1252")

    # Mock du VBProject
    mock_component = Mock()
    mock_component.Name = "SimpleClass"
    mock_component.CodeModule.CountOfLines = 2

    mock_vb_project = Mock()
    mock_vb_project.Name = "test.xlsm"
    mock_vb_project.VBComponents.Add.return_value = mock_component
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.import_module(cls_file)

    assert result.name == "SimpleClass"
    assert result.module_type == "class"
    assert result.has_predeclared_id is False


def test_import_module_already_exists_no_overwrite(mock_excel_manager, tmp_path):
    """Test error when module exists and overwrite=False."""
    bas_file = tmp_path / "Module1.bas"
    bas_content = 'Attribute VB_Name = "Module1"\\nSub Test()\\nEnd Sub'
    bas_file.write_text(bas_content, encoding="windows-1252")

    # Mock d'un module existant
    existing_component = Mock()
    existing_component.Name = "Module1"

    # Mock du nouveau composant importé
    new_component = Mock()
    new_component.Name = "Module1"

    mock_vb_project = Mock()
    mock_vb_project.Name = "test.xlsm"
    mock_vb_project.VBComponents.Import.return_value = new_component
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([existing_component]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleAlreadyExistsError):
        vba_mgr.import_module(bas_file, overwrite=False)

    # Vérifier que le module importé a été retiré
    mock_vb_project.VBComponents.Remove.assert_called_once_with(new_component)


def test_import_module_with_overwrite(mock_excel_manager, tmp_path):
    """Test successful import with overwrite=True."""
    bas_file = tmp_path / "Module1.bas"
    bas_content = 'Attribute VB_Name = "Module1"\\nSub Test()\\nEnd Sub'
    bas_file.write_text(bas_content, encoding="windows-1252")

    # Mock du composant importé
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.CodeModule.CountOfLines = 2

    mock_vb_project = Mock()
    mock_vb_project.Name = "test.xlsm"
    mock_vb_project.VBComponents.Import.return_value = mock_component
    # Module existant ne devrait pas poser de problème avec overwrite=True
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([mock_component]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.import_module(bas_file, overwrite=True)

    assert result.name == "Module1"
    # Avec overwrite, pas de Remove appelé car l'import écrase automatiquement
    mock_vb_project.VBComponents.Remove.assert_not_called()


def test_import_class_module_with_overwrite(mock_excel_manager, tmp_path):
    """Test import of .cls module with overwrite=True."""
    cls_file = tmp_path / "MyClass.cls"
    cls_content = '''VERSION 1.0 CLASS
Attribute VB_Name = "MyClass"
Option Explicit

Public Sub Test()
End Sub
'''
    cls_file.write_text(cls_content, encoding="windows-1252")

    # Mock d'un composant existant
    existing_component = Mock()
    existing_component.Name = "MyClass"

    # Mock du nouveau composant
    new_component = Mock()
    new_component.Name = "MyClass"
    new_component.CodeModule.CountOfLines = 2

    mock_vb_project = Mock()
    mock_vb_project.Name = "test.xlsm"
    mock_vb_project.VBComponents.Add.return_value = new_component
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([existing_component]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.import_module(cls_file, overwrite=True)

    assert result.name == "MyClass"
    # Vérifier que l'ancien module a été supprimé
    mock_vb_project.VBComponents.Remove.assert_called_once_with(existing_component)


def test_import_userform_missing_frx(mock_excel_manager, tmp_path):
    """Test error when .frx file is missing for UserForm."""
    frm_file = tmp_path / "UserForm1.frm"
    frm_file.write_text("Dummy content", encoding="windows-1252")
    # Pas de fichier .frx

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAImportError, match=".frx manquant"):
        vba_mgr.import_module(frm_file)


def test_import_userform_success(mock_excel_manager, tmp_path):
    """Test successful import of UserForm with .frm and .frx."""
    frm_file = tmp_path / "UserForm1.frm"
    frx_file = tmp_path / "UserForm1.frx"
    frm_file.write_text("UserForm content", encoding="windows-1252")
    frx_file.write_bytes(b"Binary FRX data")  # Fichier binaire

    # Mock du composant UserForm
    mock_component = Mock()
    mock_component.Name = "UserForm1"
    mock_component.CodeModule.CountOfLines = 5

    mock_vb_project = Mock()
    mock_vb_project.Name = "test.xlsm"
    mock_vb_project.VBComponents.Import.return_value = mock_component
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.import_module(frm_file)

    assert result.name == "UserForm1"
    assert result.module_type == "userform"
    assert result.lines_count == 5
    assert result.has_predeclared_id is True  # UserForms ont toujours PredeclaredId=True
    mock_vb_project.VBComponents.Import.assert_called_once()


def test_import_module_invalid_type(mock_excel_manager, tmp_path):
    """Test error when module type is not supported."""
    txt_file = tmp_path / "invalid.txt"
    txt_file.write_text("Some content")

    mock_excel_manager.app.ActiveWorkbook.Name = "test.xlsm"

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAImportError, match="non reconnue"):
        vba_mgr.import_module(txt_file)


def test_import_module_com_error(mock_excel_manager, tmp_path):
    """Test handling of COM errors during import."""
    bas_file = tmp_path / "Module1.bas"
    bas_content = 'Attribute VB_Name = "Module1"\\nSub Test()\\nEnd Sub'
    bas_file.write_text(bas_content, encoding="windows-1252")

    # Mock d'une erreur COM
    mock_vb_project = Mock()
    mock_vb_project.Name = "test.xlsm"
    mock_vb_project.VBComponents.Import.side_effect = pywintypes.com_error(
        -2147352567, "Automation error", None, None
    )

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAImportError, match="Erreur COM"):
        vba_mgr.import_module(bas_file)
