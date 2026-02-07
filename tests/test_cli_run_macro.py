"""
Tests CLI pour la commande run-macro.

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
from typer.testing import CliRunner
from unittest.mock import Mock, patch

from xlmanage.cli import app
from xlmanage.macro_runner import MacroResult
from xlmanage.exceptions import VBAMacroError, WorkbookNotFoundError


runner = CliRunner()


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_success_sub(mock_runner_class, mock_mgr_class):
    """Test exécution réussie d'un Sub (pas de retour)."""
    # Mock du manager
    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None
    mock_mgr.start = Mock()

    # Mock du runner
    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner

    result_obj = MacroResult(
        macro_name="Module1.Test",
        return_value=None,
        return_type="NoneType",
        success=True,
        error_message=None,
    )
    mock_runner.run.return_value = result_obj

    # Exécuter la commande
    result = runner.invoke(app, ["run-macro", "Module1.Test"])

    assert result.exit_code == 0
    assert "✅" in result.stdout
    assert "Module1.Test" in result.stdout
    mock_runner.run.assert_called_once_with(
        macro_name="Module1.Test", workbook=None, args=None
    )


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_success_function(mock_runner_class, mock_mgr_class):
    """Test exécution réussie d'une Function avec retour."""
    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None

    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner

    result_obj = MacroResult(
        macro_name="Module1.GetSum",
        return_value=42,
        return_type="int",
        success=True,
        error_message=None,
    )
    mock_runner.run.return_value = result_obj

    result = runner.invoke(app, ["run-macro", "Module1.GetSum", "--args", "10,20"])

    assert result.exit_code == 0
    assert "✅" in result.stdout
    assert "42" in result.stdout
    assert "int" in result.stdout


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_vba_error(mock_runner_class, mock_mgr_class):
    """Test erreur VBA runtime."""
    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None

    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner

    result_obj = MacroResult(
        macro_name="Module1.Divide",
        return_value=None,
        return_type="NoneType",
        success=False,
        error_message="Division by zero",
    )
    mock_runner.run.return_value = result_obj

    result = runner.invoke(app, ["run-macro", "Module1.Divide", "--args", "10,0"])

    assert result.exit_code == 1
    assert "❌" in result.stdout
    assert "Division by zero" in result.stdout


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_with_workbook(mock_runner_class, mock_mgr_class, tmp_path):
    """Test avec workbook spécifié."""
    # Créer un fichier temporaire
    workbook_file = tmp_path / "test.xlsm"
    workbook_file.touch()

    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None

    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner

    result_obj = MacroResult(
        macro_name="'test.xlsm'!Module1.Test",
        return_value="OK",
        return_type="str",
        success=True,
        error_message=None,
    )
    mock_runner.run.return_value = result_obj

    result = runner.invoke(
        app, ["run-macro", "Module1.Test", "--workbook", str(workbook_file)]
    )

    assert result.exit_code == 0
    # Vérifier que run() a été appelé avec le bon Path
    call_args = mock_runner.run.call_args
    assert call_args.kwargs["workbook"] == workbook_file


def test_run_macro_workbook_not_found():
    """Test erreur si fichier workbook introuvable."""
    result = runner.invoke(
        app, ["run-macro", "Module1.Test", "--workbook", "C:/nonexistent.xlsm"]
    )

    assert result.exit_code == 1
    assert "introuvable" in result.stdout.lower()


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_macro_not_found(mock_runner_class, mock_mgr_class):
    """Test erreur macro introuvable."""
    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None

    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner
    mock_runner.run.side_effect = VBAMacroError(
        macro_name="Module1.Missing", reason="Macro introuvable"
    )

    result = runner.invoke(app, ["run-macro", "Module1.Missing"])

    assert result.exit_code == 1
    assert "❌" in result.stdout
    assert "Macro introuvable" in result.stdout


@patch("xlmanage.cli.ExcelManager")
def test_run_macro_help(mock_mgr_class):
    """Test affichage de l'aide."""
    result = runner.invoke(app, ["run-macro", "--help"])

    assert result.exit_code == 0
    assert "macro_name" in result.stdout.lower() or "macro" in result.stdout.lower()
    assert "--workbook" in result.stdout
    assert "--args" in result.stdout
    assert "--timeout" in result.stdout
    assert "Exemples:" in result.stdout
