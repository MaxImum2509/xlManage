"""
Tests for ScreenOptimizer class.

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

from xlmanage.screen_optimizer import ScreenOptimizer
from xlmanage.excel_optimizer import OptimizationState


@pytest.fixture
def mock_app():
    """Fixture providing a mock Excel Application."""
    app = Mock()
    app.ScreenUpdating = True
    app.DisplayStatusBar = True
    app.EnableAnimations = True
    return app


def test_screen_optimizer_init(mock_app):
    """Test ScreenOptimizer initialization."""
    optimizer = ScreenOptimizer(mock_app)
    assert optimizer._app is mock_app
    assert optimizer._original_settings == {}


def test_screen_optimizer_apply_restore(mock_app):
    """Test apply/restore workflow without context manager."""
    optimizer = ScreenOptimizer(mock_app)

    # État initial
    mock_app.ScreenUpdating = True
    mock_app.DisplayStatusBar = True
    mock_app.EnableAnimations = True

    # Appliquer les optimisations
    state = optimizer.apply()

    # Vérifier que les optimisations sont appliquées
    assert mock_app.ScreenUpdating is False
    assert mock_app.DisplayStatusBar is False
    assert mock_app.EnableAnimations is False

    # Vérifier l'état retourné
    assert isinstance(state, OptimizationState)
    assert state.optimizer_type == "screen"
    assert state.applied_at  # Timestamp présent
    assert len(state.screen) == 3  # 3 propriétés sauvegardées

    # Restaurer
    optimizer.restore()

    # Vérifier que les valeurs originales sont restaurées
    assert mock_app.ScreenUpdating is True
    assert mock_app.DisplayStatusBar is True
    assert mock_app.EnableAnimations is True


def test_screen_optimizer_restore_without_apply(mock_app):
    """Test error when calling restore() before apply()."""
    optimizer = ScreenOptimizer(mock_app)

    with pytest.raises(RuntimeError, match="no settings were saved"):
        optimizer.restore()


def test_screen_optimizer_get_current_settings(mock_app):
    """Test get_current_settings() returns screen properties."""
    optimizer = ScreenOptimizer(mock_app)

    settings = optimizer.get_current_settings()

    assert "ScreenUpdating" in settings
    assert "DisplayStatusBar" in settings
    assert "EnableAnimations" in settings
    assert len(settings) == 3


def test_screen_optimizer_context_manager_still_works(mock_app):
    """Test that existing context manager usage still works."""
    optimizer = ScreenOptimizer(mock_app)

    mock_app.ScreenUpdating = True
    mock_app.DisplayStatusBar = True

    with optimizer:
        # Optimisations appliquées
        assert mock_app.ScreenUpdating is False
        assert mock_app.DisplayStatusBar is False

    # Restaurées après le with
    assert mock_app.ScreenUpdating is True
    assert mock_app.DisplayStatusBar is True


def test_screen_optimizer_apply_exception_handling(mock_app):
    """Test that exceptions during apply are handled gracefully."""
    optimizer = ScreenOptimizer(mock_app)

    # Simuler une erreur sur une propriété
    def raise_error():
        raise Exception("COM error")

    type(mock_app).ScreenUpdating = property(lambda self: True, lambda self, v: raise_error())

    # L'apply ne doit pas lever d'exception
    state = optimizer.apply()
    assert isinstance(state, OptimizationState)


def test_screen_optimizer_get_current_settings_error(mock_app):
    """Test get_current_settings with COM errors."""
    optimizer = ScreenOptimizer(mock_app)

    # Simuler des erreurs sur toutes les propriétés
    for prop in ["ScreenUpdating", "DisplayStatusBar", "EnableAnimations"]:
        delattr(mock_app, prop)

    settings = optimizer.get_current_settings()
    assert settings == {}


def test_screen_optimizer_context_manager_with_exception(mock_app):
    """Test context manager restores settings even with exception."""
    optimizer = ScreenOptimizer(mock_app)

    mock_app.ScreenUpdating = True

    try:
        with optimizer:
            assert mock_app.ScreenUpdating is False
            raise ValueError("Test exception")
    except ValueError:
        pass

    # Les paramètres doivent être restaurés malgré l'exception
    assert mock_app.ScreenUpdating is True


def test_screen_optimizer_optimization_state_structure(mock_app):
    """Test that OptimizationState has correct structure for screen optimizer."""
    optimizer = ScreenOptimizer(mock_app)

    state = optimizer.apply()

    # Vérifier la structure de l'état
    assert state.optimizer_type == "screen"
    assert len(state.screen) == 3
    assert state.calculation == {}
    assert state.full == {}
    assert state.applied_at

    optimizer.restore()
