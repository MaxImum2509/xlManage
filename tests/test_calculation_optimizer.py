"""
Tests for CalculationOptimizer class.

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

from xlmanage.calculation_optimizer import CalculationOptimizer
from xlmanage.excel_optimizer import OptimizationState


@pytest.fixture
def mock_app():
    """Fixture providing a mock Excel Application."""
    app = Mock()
    app.Calculation = -4105  # xlCalculationAutomatic
    app.Iteration = False
    app.MaxIterations = 100
    app.MaxChange = 0.001
    return app


def test_calculation_optimizer_init(mock_app):
    """Test CalculationOptimizer initialization."""
    optimizer = CalculationOptimizer(mock_app)
    assert optimizer._app is mock_app
    assert optimizer._original_settings == {}


def test_calculation_optimizer_apply_restore(mock_app):
    """Test apply/restore workflow without context manager."""
    optimizer = CalculationOptimizer(mock_app)

    # État initial
    mock_app.Calculation = -4105  # xlCalculationAutomatic
    mock_app.Iteration = False

    # Appliquer les optimisations
    state = optimizer.apply()

    # Vérifier que les optimisations sont appliquées
    assert mock_app.Calculation == -4135  # xlCalculationManual
    assert mock_app.Iteration is False

    # Vérifier l'état retourné
    assert isinstance(state, OptimizationState)
    assert state.optimizer_type == "calculation"
    assert state.applied_at  # Timestamp présent
    assert len(state.calculation) == 4  # 4 propriétés sauvegardées

    # Restaurer
    optimizer.restore()

    # Vérifier que les valeurs originales sont restaurées
    assert mock_app.Calculation == -4105


def test_calculation_optimizer_restore_without_apply(mock_app):
    """Test error when calling restore() before apply()."""
    optimizer = CalculationOptimizer(mock_app)

    with pytest.raises(RuntimeError, match="no settings were saved"):
        optimizer.restore()


def test_calculation_optimizer_get_current_settings(mock_app):
    """Test get_current_settings() returns calculation properties."""
    optimizer = CalculationOptimizer(mock_app)

    settings = optimizer.get_current_settings()

    assert "Calculation" in settings
    assert "Iteration" in settings
    assert "MaxIterations" in settings
    assert "MaxChange" in settings
    assert len(settings) == 4


def test_calculation_optimizer_context_manager_still_works(mock_app):
    """Test that existing context manager usage still works."""
    optimizer = CalculationOptimizer(mock_app)

    mock_app.Calculation = -4105  # xlCalculationAutomatic

    with optimizer:
        # Optimisations appliquées
        assert mock_app.Calculation == -4135  # xlCalculationManual

    # Restaurées après le with
    assert mock_app.Calculation == -4105


def test_calculation_optimizer_apply_exception_handling(mock_app):
    """Test that exceptions during apply are handled gracefully."""
    optimizer = CalculationOptimizer(mock_app)

    # Simuler une erreur sur une propriété
    def raise_error():
        raise Exception("COM error")

    type(mock_app).Calculation = property(lambda self: -4105, lambda self, v: raise_error())

    # L'apply ne doit pas lever d'exception
    state = optimizer.apply()
    assert isinstance(state, OptimizationState)


def test_calculation_optimizer_get_current_settings_error(mock_app):
    """Test get_current_settings with COM errors."""
    optimizer = CalculationOptimizer(mock_app)

    # Simuler des erreurs sur toutes les propriétés
    for prop in ["Calculation", "Iteration", "MaxIterations", "MaxChange"]:
        delattr(mock_app, prop)

    settings = optimizer.get_current_settings()
    assert settings == {}


def test_calculation_optimizer_context_manager_with_exception(mock_app):
    """Test context manager restores settings even with exception."""
    optimizer = CalculationOptimizer(mock_app)

    mock_app.Calculation = -4105

    try:
        with optimizer:
            assert mock_app.Calculation == -4135
            raise ValueError("Test exception")
    except ValueError:
        pass

    # Les paramètres doivent être restaurés malgré l'exception
    assert mock_app.Calculation == -4105


def test_calculation_optimizer_optimization_state_structure(mock_app):
    """Test that OptimizationState has correct structure for calculation optimizer."""
    optimizer = CalculationOptimizer(mock_app)

    state = optimizer.apply()

    # Vérifier la structure de l'état
    assert state.optimizer_type == "calculation"
    assert len(state.calculation) == 4
    assert state.screen == {}
    assert state.full == {}
    assert state.applied_at

    optimizer.restore()
