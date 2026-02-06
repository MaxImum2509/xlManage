"""Tests for VBA utility functions."""

from pathlib import Path
from unittest.mock import Mock, PropertyMock

import pywintypes
import pytest

from xlmanage.exceptions import (
    VBAImportError,
    VBAProjectAccessError,
    VBAWorkbookFormatError,
)
from xlmanage.vba_manager import (
    _detect_module_type,
    _find_component,
    _get_vba_project,
    _parse_class_module,
)


def test_get_vba_project_success():
    """Test successful VBProject access."""
    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_vb_project = Mock()
    mock_wb.VBProject = mock_vb_project

    result = _get_vba_project(mock_wb)
    assert result == mock_vb_project


def test_get_vba_project_xlsx_format():
    """Test error when workbook is .xlsx format."""
    mock_wb = Mock()
    mock_wb.Name = "test.xlsx"

    with pytest.raises(VBAWorkbookFormatError) as exc_info:
        _get_vba_project(mock_wb)
    assert exc_info.value.workbook_name == "test.xlsx"


def test_get_vba_project_access_denied():
    """Test error when Trust Center blocks access."""
    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"

    # Simuler l'erreur COM 0x800A03EC
    com_error = pywintypes.com_error(-2146827284, "Access denied", None, None)
    type(mock_wb).VBProject = PropertyMock(side_effect=com_error)

    with pytest.raises(VBAProjectAccessError) as exc_info:
        _get_vba_project(mock_wb)
    assert exc_info.value.workbook_name == "test.xlsm"


def test_find_component_found():
    """Test finding an existing component."""
    mock_comp1 = Mock()
    mock_comp1.Name = "Module1"
    mock_comp2 = Mock()
    mock_comp2.Name = "Module2"

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_comp1, mock_comp2]

    result = _find_component(mock_vb_project, "Module2")
    assert result == mock_comp2


def test_find_component_not_found():
    """Test component not found returns None."""
    mock_comp = Mock()
    mock_comp.Name = "Module1"

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_comp]

    result = _find_component(mock_vb_project, "Module99")
    assert result is None


def test_detect_module_type_bas():
    """Test detection of .bas module."""
    assert _detect_module_type(Path("Module1.bas")) == "standard"
    assert _detect_module_type(Path("Module1.BAS")) == "standard"


def test_detect_module_type_cls():
    """Test detection of .cls module."""
    assert _detect_module_type(Path("Class1.cls")) == "class"


def test_detect_module_type_frm():
    """Test detection of .frm module."""
    assert _detect_module_type(Path("UserForm1.frm")) == "userform"


def test_detect_module_type_invalid():
    """Test error with invalid extension."""
    with pytest.raises(VBAImportError) as exc_info:
        _detect_module_type(Path("file.txt"))
    assert ".txt" in str(exc_info.value)


def test_parse_class_module_success(tmp_path):
    """Test parsing a valid .cls file."""
    cls_content = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Hello()
    MsgBox "Hello"
End Sub
'''
    cls_file = tmp_path / "MyClass.cls"
    cls_file.write_text(cls_content, encoding="windows-1252")

    name, predeclared, code = _parse_class_module(cls_file)

    assert name == "MyClass"
    assert predeclared is True
    assert "Option Explicit" in code
    assert "Public Sub Hello()" in code
    assert "Attribute" not in code


def test_parse_class_module_no_predeclared(tmp_path):
    """Test parsing .cls without PredeclaredId."""
    cls_content = '''Attribute VB_Name = "SimpleClass"
Option Explicit
'''
    cls_file = tmp_path / "SimpleClass.cls"
    cls_file.write_text(cls_content, encoding="windows-1252")

    name, predeclared, code = _parse_class_module(cls_file)

    assert name == "SimpleClass"
    assert predeclared is False


def test_parse_class_module_invalid_encoding(tmp_path):
    """Test error with wrong encoding."""
    cls_file = tmp_path / "bad.cls"
    # Use bytes that are invalid in windows-1252
    # 0x81, 0x8D, 0x8F, 0x90, 0x9D are undefined in windows-1252
    cls_file.write_bytes(b"Attribute VB_Name = \x81\x8D\x8F\x90\x9D")

    with pytest.raises(VBAImportError) as exc_info:
        _parse_class_module(cls_file)
    assert "windows-1252" in str(exc_info.value)


def test_parse_class_module_missing_vb_name(tmp_path):
    """Test error when VB_Name is missing."""
    cls_content = "Option Explicit\nPublic Sub Test()\nEnd Sub"
    cls_file = tmp_path / "bad.cls"
    cls_file.write_text(cls_content, encoding="windows-1252")

    with pytest.raises(VBAImportError) as exc_info:
        _parse_class_module(cls_file)
    assert "VB_Name manquant" in str(exc_info.value)
