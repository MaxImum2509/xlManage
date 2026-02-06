"""Tests pour l'exécution de macros VBA."""

import pytest
from pathlib import Path
from unittest.mock import Mock, patch

import pywintypes

from xlmanage.macro_runner import (
    MacroRunner,
    MacroResult,
    _build_macro_reference,
    _format_return_value,
)
from xlmanage.exceptions import VBAMacroError, WorkbookNotFoundError
from xlmanage.excel_manager import ExcelManager


@pytest.fixture
def mock_excel_manager():
    """Fixture pour un ExcelManager mocké."""
    mgr = Mock(spec=ExcelManager)
    mgr.app = Mock()
    mgr.app.Workbooks = []
    return mgr


def test_macro_result_success():
    """Test MacroResult avec succès."""
    result = MacroResult(
        macro_name="Module1.Test",
        return_value=42,
        return_type="int",
        success=True,
        error_message=None,
    )

    assert result.success
    assert result.return_value == 42
    assert "42" in str(result)
    assert "✅" in str(result)


def test_macro_result_failure():
    """Test MacroResult avec erreur."""
    result = MacroResult(
        macro_name="Module1.Test",
        return_value=None,
        return_type="NoneType",
        success=False,
        error_message="Division by zero",
    )

    assert not result.success
    assert result.error_message == "Division by zero"
    assert "❌" in str(result)
    assert "Division by zero" in str(result)


def test_build_macro_reference_no_workbook(mock_excel_manager):
    """Test construction référence sans workbook."""
    ref = _build_macro_reference("MySub", None, mock_excel_manager.app)
    assert ref == "MySub"


def test_build_macro_reference_with_workbook(mock_excel_manager):
    """Test construction référence avec workbook."""
    # Mock du classeur ouvert
    mock_wb = Mock()
    mock_wb.Name = "data.xlsm"
    mock_excel_manager.app.Workbooks = [mock_wb]

    ref = _build_macro_reference(
        "Module1.Test", Path("data.xlsm"), mock_excel_manager.app
    )

    assert ref == "'data.xlsm'!Module1.Test"


def test_build_macro_reference_workbook_not_found(mock_excel_manager):
    """Test erreur si classeur non ouvert."""
    mock_excel_manager.app.Workbooks = []

    with pytest.raises(WorkbookNotFoundError) as exc_info:
        _build_macro_reference("Module1.Test", Path("missing.xlsm"), mock_excel_manager.app)

    assert "missing.xlsm" in str(exc_info.value)


def test_format_return_value_none():
    """Test formatage de None."""
    assert _format_return_value(None) == "(aucune valeur de retour)"


def test_format_return_value_simple():
    """Test formatage valeurs simples."""
    assert _format_return_value(42) == "42"
    assert _format_return_value("hello") == "hello"
    assert _format_return_value(3.14) == "3.14"


def test_format_return_value_array():
    """Test formatage tableau VBA."""
    array = ((1, 2, 3), (4, 5, 6))
    result = _format_return_value(array)

    assert "Tableau 2x3" in result
    assert "[[1, 2, 3], [4, 5, 6]]" in result


def test_format_return_value_datetime():
    """Test formatage date VBA (pywintypes.TimeType)."""
    from datetime import datetime

    # Créer un pywintypes.TimeType
    dt = pywintypes.Time(datetime(2024, 1, 15, 10, 30, 0))
    result = _format_return_value(dt)

    # Le résultat doit être au format ISO
    assert "2024-01-15" in result
    assert "10:30:00" in result


def test_macro_runner_init(mock_excel_manager):
    """Test initialisation MacroRunner."""
    runner = MacroRunner(mock_excel_manager)
    assert runner._mgr == mock_excel_manager


def test_macro_runner_run_sub_success(mock_excel_manager):
    """Test exécution d'un Sub VBA (pas de retour)."""
    mock_excel_manager.app.Run.return_value = None

    runner = MacroRunner(mock_excel_manager)
    result = runner.run("Module1.MySub")

    assert result.success
    assert result.return_value is None
    assert result.return_type == "NoneType"
    mock_excel_manager.app.Run.assert_called_once_with("Module1.MySub")


def test_macro_runner_run_function_success(mock_excel_manager):
    """Test exécution d'une Function VBA avec retour."""
    mock_excel_manager.app.Run.return_value = 42

    runner = MacroRunner(mock_excel_manager)
    result = runner.run("Module1.GetAnswer")

    assert result.success
    assert result.return_value == 42
    assert result.return_type == "int"


def test_macro_runner_run_with_args(mock_excel_manager):
    """Test exécution avec arguments."""
    mock_excel_manager.app.Run.return_value = "Hello, World"

    runner = MacroRunner(mock_excel_manager)
    result = runner.run("Module1.Greet", args='"World"')

    assert result.success
    assert result.return_value == "Hello, World"
    # Vérifier que app.Run a été appelé avec les bons arguments
    mock_excel_manager.app.Run.assert_called_once()
    call_args = mock_excel_manager.app.Run.call_args[0]
    assert call_args[0] == "Module1.Greet"
    assert call_args[1] == "World"


def test_macro_runner_run_vba_error(mock_excel_manager):
    """Test capture erreur VBA runtime."""
    # Simuler une com_error avec excepinfo
    com_error = pywintypes.com_error(
        0x800A03EC, "VBA error", (None, None, "Division by zero", None, None, 0), None
    )
    mock_excel_manager.app.Run.side_effect = com_error

    runner = MacroRunner(mock_excel_manager)
    result = runner.run("Module1.Divide", args="10,0")

    assert not result.success
    assert result.error_message == "Division by zero"
    assert result.return_value is None


def test_macro_runner_run_macro_not_found(mock_excel_manager):
    """Test erreur macro introuvable."""
    # Simuler erreur COM sans excepinfo (macro introuvable)
    com_error = pywintypes.com_error(0x80030000, "Macro not found", None, None)  # HRESULT différent
    mock_excel_manager.app.Run.side_effect = com_error

    runner = MacroRunner(mock_excel_manager)

    with pytest.raises(VBAMacroError) as exc_info:
        runner.run("Module1.Missing")

    assert "0x80030000" in str(exc_info.value)


def test_macro_runner_run_with_workbook(mock_excel_manager):
    """Test exécution avec workbook spécifié."""
    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_excel_manager.app.Workbooks = [mock_wb]
    mock_excel_manager.app.Run.return_value = 100

    runner = MacroRunner(mock_excel_manager)
    result = runner.run("Module1.Calc", workbook=Path("test.xlsm"))

    assert result.success
    assert result.return_value == 100
    # Vérifier que la référence complète a été utilisée
    mock_excel_manager.app.Run.assert_called_once()
    call_args = mock_excel_manager.app.Run.call_args[0]
    assert call_args[0] == "'test.xlsm'!Module1.Calc"
