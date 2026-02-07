"""
Screen optimizer for Excel display settings optimization.

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

from datetime import datetime
from typing import TYPE_CHECKING, Any

from .excel_optimizer import OptimizationState

if TYPE_CHECKING:
    from .excel_manager import ExcelManager


class ScreenOptimizer:
    """Optimiseur des propriétés d'affichage Excel.

    Gère les 3 propriétés d'écran pour améliorer les performances visuelles :
    - ScreenUpdating : Mise à jour de l'écran
    - DisplayStatusBar : Affichage de la barre de statut
    - EnableAnimations : Animations de l'interface

    Usage avec context manager :
        >>> with ScreenOptimizer(app):
        ...     # Affichage désactivé
        ...     for i in range(1000):
        ...         sheet.Cells(i, 1).Value = i

    Usage avec apply/restore :
        >>> optimizer = ScreenOptimizer(app)
        >>> optimizer.apply()
        >>> # Affichage désactivé
        >>> optimizer.restore()
    """

    def __init__(self, excel_manager: "ExcelManager") -> None:
        """Initialise l'optimiseur d'écran.

        Args:
            excel_manager: Instance ExcelManager (doit être démarrée)
        """
        self._mgr = excel_manager
        self._app = excel_manager.app
        self._original_settings: dict[str, Any] = {}

    def __enter__(self) -> "ScreenOptimizer":
        """Entre dans le context manager et applique les optimisations."""
        self._save_current_settings()
        self._apply_optimizations()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Sort du context manager et restaure les paramètres originaux."""
        self._restore_original_settings()

    def apply(self) -> OptimizationState:
        """Applique les optimisations d'écran SANS context manager.

        Returns:
            OptimizationState: État sauvegardé avant l'application
        """
        self._save_current_settings()
        self._apply_optimizations()

        return OptimizationState(
            screen=self._original_settings.copy() if self._original_settings else {},
            calculation={},
            full={},
            applied_at=datetime.now().isoformat(),
            optimizer_type="screen",
        )

    def restore(self) -> None:
        """Restaure les paramètres d'écran sauvegardés.

        Raises:
            RuntimeError: Si apply() n'a pas été appelé avant
        """
        if not self._original_settings:
            raise RuntimeError(
                "Cannot restore: no settings were saved. Call apply() first."
            )

        self._restore_original_settings()

    def get_current_settings(self) -> dict[str, object]:
        """Retourne l'état actuel des propriétés d'écran.

        Returns:
            dict[str, object]: {ScreenUpdating, DisplayStatusBar, EnableAnimations}
        """
        try:
            return {
                "ScreenUpdating": self._app.ScreenUpdating,
                "DisplayStatusBar": self._app.DisplayStatusBar,
                "EnableAnimations": self._app.EnableAnimations,
            }
        except Exception:
            return {}

    def _save_current_settings(self) -> None:
        """Sauvegarde les paramètres d'écran actuels."""
        try:
            self._original_settings = {
                "ScreenUpdating": self._app.ScreenUpdating,
                "DisplayStatusBar": self._app.DisplayStatusBar,
                "EnableAnimations": self._app.EnableAnimations,
            }
        except Exception:
            self._original_settings = {}

    def _apply_optimizations(self) -> None:
        """Applique les optimisations d'écran."""
        try:
            self._app.ScreenUpdating = False
            self._app.DisplayStatusBar = False
            self._app.EnableAnimations = False
        except Exception:
            pass

    def _restore_original_settings(self) -> None:
        """Restaure les paramètres d'écran originaux."""
        if not self._original_settings:
            return

        try:
            for prop, value in self._original_settings.items():
                setattr(self._app, prop, value)
        except Exception:
            pass

        self._original_settings = {}
