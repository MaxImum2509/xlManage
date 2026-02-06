"""Tests for VBAManager initialization and dataclass."""

import pytest
from unittest.mock import Mock, PropertyMock

from xlmanage.excel_manager import ExcelManager
from xlmanage.vba_manager import VBAManager, VBAModuleInfo


def test_vba_module_info_creation():
    """Test VBAModuleInfo dataclass creation."""
    info = VBAModuleInfo(
        name="Module1",
        module_type="standard",
        lines_count=42,
        has_predeclared_id=False,
    )

    assert info.name == "Module1"
    assert info.module_type == "standard"
    assert info.lines_count == 42
    assert info.has_predeclared_id is False


def test_vba_module_info_defaults():
    """Test VBAModuleInfo default values."""
    info = VBAModuleInfo(name="MyClass", module_type="class", lines_count=10)

    # has_predeclared_id a une valeur par défaut False
    assert info.has_predeclared_id is False


def test_vba_manager_init():
    """Test VBAManager initialization with ExcelManager."""
    # Créer un mock ExcelManager
    mock_excel_mgr = Mock(spec=ExcelManager)
    mock_app = Mock()
    mock_excel_mgr.app = mock_app

    # Créer le VBAManager
    vba_mgr = VBAManager(mock_excel_mgr)

    # Vérifier que l'ExcelManager est stocké
    assert vba_mgr._mgr is mock_excel_mgr


def test_vba_manager_app_property():
    """Test VBAManager.app property delegates to ExcelManager."""
    mock_excel_mgr = Mock(spec=ExcelManager)
    mock_app = Mock()
    mock_excel_mgr.app = mock_app

    vba_mgr = VBAManager(mock_excel_mgr)

    # Vérifier que .app renvoie l'app de l'ExcelManager
    assert vba_mgr.app is mock_app


def test_vba_manager_app_property_not_started():
    """Test VBAManager.app raises when Excel not started."""
    mock_excel_mgr = Mock(spec=ExcelManager)
    # Simuler que Excel n'est pas démarré
    type(mock_excel_mgr).app = PropertyMock(side_effect=RuntimeError("Excel not started"))

    vba_mgr = VBAManager(mock_excel_mgr)

    with pytest.raises(RuntimeError, match="Excel not started"):
        _ = vba_mgr.app
