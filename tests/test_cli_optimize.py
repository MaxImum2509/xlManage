"""
Tests for CLI optimize command.

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
from typer.testing import CliRunner
from unittest.mock import Mock, patch, MagicMock

from xlmanage.cli import app
from xlmanage.excel_optimizer import OptimizationState

runner = CliRunner()


def test_optimize_screen():
    """Test optimize --screen command."""
    mock_state = OptimizationState(
        screen={"ScreenUpdating": True, "DisplayStatusBar": True, "EnableAnimations": True},
        calculation={},
        full={},
        applied_at="2026-02-06T10:00:00",
        optimizer_type="screen",
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.screen_optimizer.ScreenOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.apply.return_value = mock_state
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--screen"])

        assert result.exit_code == 0
        assert "appliquées avec succès" in result.stdout
        mock_opt.apply.assert_called_once()


def test_optimize_calculation():
    """Test optimize --calculation command."""
    mock_state = OptimizationState(
        screen={},
        calculation={"Calculation": -4135, "Iteration": False},
        full={},
        applied_at="2026-02-06T10:00:00",
        optimizer_type="calculation",
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.calculation_optimizer.CalculationOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.apply.return_value = mock_state
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--calculation"])

        assert result.exit_code == 0
        assert "appliquées avec succès" in result.stdout
        mock_opt.apply.assert_called_once()


def test_optimize_all():
    """Test optimize --all command."""
    mock_state = OptimizationState(
        screen={},
        calculation={},
        full={
            "ScreenUpdating": True,
            "DisplayAlerts": True,
            "Calculation": -4105,
        },
        applied_at="2026-02-06T10:00:00",
        optimizer_type="all",
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.excel_optimizer.ExcelOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.apply.return_value = mock_state
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--all"])

        assert result.exit_code == 0
        assert "appliquées avec succès" in result.stdout
        mock_opt.apply.assert_called_once()


def test_optimize_default_is_all():
    """Test that optimize without options defaults to --all."""
    mock_state = OptimizationState(
        screen={},
        calculation={},
        full={"ScreenUpdating": True},
        applied_at="2026-02-06T10:00:00",
        optimizer_type="all",
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.excel_optimizer.ExcelOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.apply.return_value = mock_state
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize"])

        assert result.exit_code == 0
        mock_opt.apply.assert_called_once()


def test_optimize_status():
    """Test optimize --status command."""
    mock_settings = {
        "ScreenUpdating": True,
        "DisplayAlerts": False,
        "Calculation": -4135,
        "EnableEvents": True,
        "DisplayStatusBar": True,
        "EnableAnimations": True,
        "AskToUpdateLinks": False,
        "Iteration": False,
    }

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.excel_optimizer.ExcelOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.get_current_settings.return_value = mock_settings
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--status"])

        assert result.exit_code == 0
        assert "État actuel" in result.stdout
        mock_opt.get_current_settings.assert_called_once()


def test_optimize_status_empty_settings():
    """Test optimize --status with empty settings."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.excel_optimizer.ExcelOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.get_current_settings.return_value = {}
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--status"])

        assert result.exit_code == 0
        assert "Impossible de lire" in result.stdout


def test_optimize_restore():
    """Test optimize --restore command."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.excel_optimizer.ExcelOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.restore.return_value = None
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--restore"])

        assert result.exit_code == 0
        assert "restaurés" in result.stdout
        mock_opt.restore.assert_called_once()


def test_optimize_restore_without_apply():
    """Test optimize --restore when no settings were saved."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.excel_optimizer.ExcelOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.restore.side_effect = RuntimeError("no settings were saved")
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--restore"])

        assert result.exit_code == 0
        assert "no settings were saved" in result.stdout


def test_optimize_force_calculate():
    """Test optimize --force-calculate command."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:
        mock_wb = Mock()
        mock_wb.Name = "test.xlsx"

        mock_app = Mock()
        mock_app.ActiveWorkbook = mock_wb
        mock_app.CalculateFullRebuild = Mock()

        mock_mgr = Mock()
        mock_mgr.app = mock_app
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        result = runner.invoke(app, ["optimize", "--force-calculate"])

        assert result.exit_code == 0
        assert "Recalcul complet terminé" in result.stdout
        mock_app.CalculateFullRebuild.assert_called_once()


def test_optimize_force_calculate_no_workbook():
    """Test optimize --force-calculate with no active workbook."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:
        mock_app = Mock()
        mock_app.ActiveWorkbook = None

        mock_mgr = Mock()
        mock_mgr.app = mock_app
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        result = runner.invoke(app, ["optimize", "--force-calculate"])

        assert result.exit_code == 0
        assert "Aucun classeur actif" in result.stdout


def test_optimize_multiple_options_error():
    """Test error when multiple options are specified."""
    result = runner.invoke(app, ["optimize", "--screen", "--calculation"])

    assert result.exit_code == 1
    assert "une seule option" in result.stdout


def test_optimize_with_visible_flag():
    """Test optimize with --visible flag."""
    mock_state = OptimizationState(
        screen={},
        calculation={},
        full={},
        applied_at="2026-02-06T10:00:00",
        optimizer_type="all",
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
        "xlmanage.excel_optimizer.ExcelOptimizer"
    ) as mock_opt_class:
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.apply.return_value = mock_state
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--all", "--visible"])

        assert result.exit_code == 0
        # Vérifier que ExcelManager a été appelé avec visible=True
        mock_mgr_class.assert_called_once_with(visible=True)
