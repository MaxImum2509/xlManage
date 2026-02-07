"""
Calculation optimizer for Excel calculation settings optimization.

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


class CalculationOptimizer:
    """Optimiseur des propriétés de calcul Excel.

    Gère les 4 propriétés de calcul pour améliorer les performances :
    - Calculation : Mode de calcul (Manuel/Automatique)
    - Iteration : Activer/Désactiver les itérations
    - MaxIterations : Nombre maximum d'itérations
    - MaxChange : Changement maximum entre itérations

    Usage avec context manager :
        >>> with CalculationOptimizer(app):
        ...     # Calcul manuel activé
        ...     wb = app.Workbooks.Open("complex.xlsx")
        ...     # Modifications sans recalcul
        ...     app.Calculate()  # Recalcul manuel

    Usage avec apply/restore :
        >>> optimizer = CalculationOptimizer(app)
        >>> optimizer.apply()
        >>> # Calcul manuel
        >>> optimizer.restore()
    """

    def __init__(self, excel_manager: "ExcelManager") -> None:
        """Initialise l'optimiseur de calcul.

        Args:
            excel_manager: Instance ExcelManager (doit être démarrée)
        """
        self._mgr = excel_manager
        self._app = excel_manager.app
        self._original_settings: dict[str, Any] = {}

    def __enter__(self) -> "CalculationOptimizer":
        """Entre dans le context manager et applique les optimisations."""
        self._save_current_settings()
        self._apply_optimizations()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Sort du context manager et restaure les paramètres originaux."""
        self._restore_original_settings()

    def apply(self) -> OptimizationState:
        """Applique les optimisations de calcul SANS context manager.

        Returns:
            OptimizationState: État sauvegardé avant l'application
        """
        self._save_current_settings()
        self._apply_optimizations()

        return OptimizationState(
            screen={},
            calculation=self._original_settings.copy()
            if self._original_settings
            else {},
            full={},
            applied_at=datetime.now().isoformat(),
            optimizer_type="calculation",
        )

    def restore(self) -> None:
        """Restaure les paramètres de calcul sauvegardés.

        Raises:
            RuntimeError: Si apply() n'a pas été appelé avant
        """
        if not self._original_settings:
            raise RuntimeError(
                "Cannot restore: no settings were saved. Call apply() first."
            )

        self._restore_original_settings()

    def get_current_settings(self) -> dict[str, object]:
        """Retourne l'état actuel des propriétés de calcul.

        Returns:
            dict[str, object]: {Calculation, Iteration, MaxIterations, MaxChange}
        """
        try:
            return {
                "Calculation": self._app.Calculation,
                "Iteration": self._app.Iteration,
                "MaxIterations": self._app.MaxIterations,
                "MaxChange": self._app.MaxChange,
            }
        except Exception:
            return {}

    def _save_current_settings(self) -> None:
        """Sauvegarde les paramètres de calcul actuels."""
        try:
            self._original_settings = {
                "Calculation": self._app.Calculation,
                "Iteration": self._app.Iteration,
                "MaxIterations": self._app.MaxIterations,
                "MaxChange": self._app.MaxChange,
            }
        except Exception:
            self._original_settings = {}

    def _apply_optimizations(self) -> None:
        """Applique les optimisations de calcul."""
        try:
            # Passer en calcul manuel
            self._app.Calculation = -4135  # xlCalculationManual

            # Désactiver l'itération
            self._app.Iteration = False
        except Exception:
            pass

    def _restore_original_settings(self) -> None:
        """Restaure les paramètres de calcul originaux."""
        if not self._original_settings:
            return

        try:
            for prop, value in self._original_settings.items():
                setattr(self._app, prop, value)
        except Exception:
            pass

        self._original_settings = {}
