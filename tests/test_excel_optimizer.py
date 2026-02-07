"""
Tests for ExcelOptimizer class.

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
from unittest.mock import Mock

from xlmanage.excel_optimizer import ExcelOptimizer, OptimizationState


@pytest.fixture
def mock_app():
    """Fixture providing a mock Excel Application."""
    app = Mock()
    # Valeurs par défaut (non optimisées)
    app.ScreenUpdating = True
    app.DisplayStatusBar = True
    app.EnableAnimations = True
    app.Calculation = -4105  # xlCalculationAutomatic
    app.EnableEvents = True
    app.DisplayAlerts = True
    app.AskToUpdateLinks = True
    app.Iteration = False
    app.MaxIterations = 100
    app.MaxChange = 0.001
    return app


@pytest.fixture
def mock_excel_mgr(mock_app):
    """Fixture providing a mock ExcelManager."""
    mgr = Mock()
    mgr.app = mock_app
    return mgr


def test_excel_optimizer_init(mock_excel_mgr, mock_app):
    """Test ExcelOptimizer initialization."""
    optimizer = ExcelOptimizer(mock_excel_mgr)
    assert optimizer._app is mock_app
    assert optimizer._original_settings == {}


def test_excel_optimizer_apply_restore(mock_excel_mgr, mock_app):
    """Test apply/restore workflow without context manager."""
    optimizer = ExcelOptimizer(mock_excel_mgr)

    # État initial
    mock_app.ScreenUpdating = True
    mock_app.DisplayAlerts = True
    mock_app.Calculation = -4105

    # Appliquer les optimisations
    state = optimizer.apply()

    # Vérifier que les optimisations sont appliquées
    assert mock_app.ScreenUpdating is False
    assert mock_app.DisplayAlerts is False
    assert mock_app.DisplayStatusBar is False
    assert mock_app.EnableAnimations is False
    assert mock_app.Calculation == -4135  # xlCalculationManual
    assert mock_app.EnableEvents is False
    assert mock_app.AskToUpdateLinks is False
    assert mock_app.Iteration is False

    # Vérifier l'état retourné
    assert isinstance(state, OptimizationState)
    assert state.optimizer_type == "all"
    assert state.applied_at  # Timestamp présent
    assert len(state.full) == 10  # 10 propriétés sauvegardées
    assert len(state.screen) == 3  # 3 propriétés d'écran
    assert len(state.calculation) == 4  # 4 propriétés de calcul

    # Restaurer
    optimizer.restore()

    # Vérifier que les valeurs originales sont restaurées
    assert mock_app.ScreenUpdating is True
    assert mock_app.DisplayAlerts is True
    assert mock_app.Calculation == -4105


def test_excel_optimizer_restore_without_apply(mock_excel_mgr, mock_app):
    """Test error when calling restore() before apply()."""
    optimizer = ExcelOptimizer(mock_excel_mgr)

    with pytest.raises(RuntimeError, match="no settings were saved"):
        optimizer.restore()


def test_excel_optimizer_get_current_settings(mock_excel_mgr, mock_app):
    """Test get_current_settings() returns all properties."""
    optimizer = ExcelOptimizer(mock_excel_mgr)

    settings = optimizer.get_current_settings()

    assert "ScreenUpdating" in settings
    assert "DisplayAlerts" in settings
    assert "Calculation" in settings
    assert "EnableEvents" in settings
    assert "EnableAnimations" in settings
    assert "DisplayStatusBar" in settings
    assert "AskToUpdateLinks" in settings
    assert "Iteration" in settings
    assert "MaxIterations" in settings
    assert "MaxChange" in settings
    assert len(settings) == 10


def test_excel_optimizer_context_manager_still_works(mock_excel_mgr, mock_app):
    """Test that existing context manager usage still works."""
    optimizer = ExcelOptimizer(mock_excel_mgr)

    mock_app.ScreenUpdating = True
    mock_app.DisplayAlerts = True

    with optimizer:
        # Optimisations appliquées
        assert mock_app.ScreenUpdating is False
        assert mock_app.DisplayAlerts is False

    # Restaurées après le with
    assert mock_app.ScreenUpdating is True
    assert mock_app.DisplayAlerts is True


def test_excel_optimizer_apply_exception_handling(mock_excel_mgr, mock_app):
    """Test that exceptions during apply are handled gracefully."""
    optimizer = ExcelOptimizer(mock_excel_mgr)

    # Simuler une erreur sur une propriété
    def raise_error():
        raise Exception("COM error")

    type(mock_app).ScreenUpdating = property(lambda self: True, lambda self, v: raise_error())

    # L'apply ne doit pas lever d'exception
    state = optimizer.apply()
    assert isinstance(state, OptimizationState)


def test_excel_optimizer_get_current_settings_error(mock_excel_mgr, mock_app):
    """Test get_current_settings with COM errors."""
    optimizer = ExcelOptimizer(mock_excel_mgr)

    # Simuler des erreurs sur toutes les propriétés
    for prop in ["ScreenUpdating", "DisplayAlerts", "Calculation"]:
        delattr(mock_app, prop)

    settings = optimizer.get_current_settings()
    assert settings == {}


def test_excel_optimizer_multiple_apply_calls(mock_excel_mgr, mock_app):
    """Test calling apply() multiple times."""
    optimizer = ExcelOptimizer(mock_excel_mgr)

    # Premier apply
    state1 = optimizer.apply()
    assert mock_app.ScreenUpdating is False

    # Deuxième apply (doit écraser le premier)
    mock_app.ScreenUpdating = False  # Déjà optimisé
    state2 = optimizer.apply()

    # Le restore doit restaurer au dernier état sauvegardé
    optimizer.restore()
    # Les paramètres sont restaurés (peu importe si c'était déjà False)


def test_excel_optimizer_context_manager_with_exception(mock_excel_mgr, mock_app):
    """Test context manager restores settings even with exception."""
    optimizer = ExcelOptimizer(mock_excel_mgr)

    mock_app.ScreenUpdating = True

    try:
        with optimizer:
            assert mock_app.ScreenUpdating is False
            raise ValueError("Test exception")
    except ValueError:
        pass

    # Les paramètres doivent être restaurés malgré l'exception
    assert mock_app.ScreenUpdating is True
