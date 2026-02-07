"""
Excel optimizer for performance tuning via COM automation.

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

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .excel_manager import ExcelManager


@dataclass
class OptimizationState:
    """État des optimisations Excel pour tracking et restauration.

    Attributes:
        screen: État des propriétés d'écran sauvegardées
            (ScreenUpdating, DisplayStatusBar, EnableAnimations)
        calculation: État des propriétés de calcul sauvegardées
            (Calculation, Iteration, MaxIterations, MaxChange)
        full: État complet des 8 propriétés (pour ExcelOptimizer)
        applied_at: Timestamp ISO de l'application des optimisations
        optimizer_type: Type d'optimizer ("screen", "calculation", "all")
    """

    screen: dict[str, object]
    calculation: dict[str, object]
    full: dict[str, object]
    applied_at: str
    optimizer_type: str


class ExcelOptimizer:
    """Optimiseur complet des performances Excel.

    Gère l'ensemble des 8 propriétés Excel critiques pour les performances :
    - ScreenUpdating, DisplayStatusBar, EnableAnimations (écran)
    - Calculation, Iteration (calcul)
    - EnableEvents, DisplayAlerts, AskToUpdateLinks (événements)

    Usage avec context manager (restauration automatique) :
        >>> with ExcelOptimizer(app):
        ...     # Excel est optimisé ici
        ...     wb = app.Workbooks.Open("large.xlsx")
        ... # Paramètres restaurés automatiquement

    Usage avec apply/restore (contrôle manuel) :
        >>> optimizer = ExcelOptimizer(app)
        >>> state = optimizer.apply()
        >>> # Excel reste optimisé
        >>> optimizer.restore()  # Restauration manuelle
    """

    def __init__(self, excel_manager: "ExcelManager") -> None:
        """Initialise l'optimiseur avec une instance Excel.

        Args:
            excel_manager: Instance ExcelManager (doit être démarrée)
        """
        self._mgr = excel_manager
        self._app = excel_manager.app
        self._original_settings: dict[str, Any] = {}

    def __enter__(self) -> "ExcelOptimizer":
        """Entre dans le context manager et applique les optimisations."""
        self._save_current_settings()
        self._apply_optimizations()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Sort du context manager et restaure les paramètres originaux."""
        self._restore_original_settings()

    def apply(self) -> OptimizationState:
        """Applique les optimisations SANS context manager.

        Les optimisations persistent jusqu'à un appel à restore().
        Cette méthode sauvegarde d'abord l'état actuel, puis applique
        les optimisations.

        Returns:
            OptimizationState: État sauvegardé avant l'application

        Example:
            >>> optimizer = ExcelOptimizer(app)
            >>> state = optimizer.apply()
            >>> # ... travail avec Excel optimisé ...
            >>> optimizer.restore()  # Restaurer l'état original
        """
        # Sauvegarder l'état actuel
        self._save_current_settings()

        # Appliquer les optimisations
        self._apply_optimizations()

        # Extraire les sous-ensembles screen et calculation
        screen_keys = {"ScreenUpdating", "DisplayStatusBar", "EnableAnimations"}
        calc_keys = {"Calculation", "Iteration", "MaxIterations", "MaxChange"}

        # Créer et retourner l'état
        return OptimizationState(
            screen={
                k: v for k, v in self._original_settings.items() if k in screen_keys
            },
            calculation={
                k: v for k, v in self._original_settings.items() if k in calc_keys
            },
            full=self._original_settings.copy() if self._original_settings else {},
            applied_at=datetime.now().isoformat(),
            optimizer_type="all",
        )

    def restore(self) -> None:
        """Restaure les paramètres sauvegardés par apply().

        Raises:
            RuntimeError: Si apply() n'a pas été appelé avant
        """
        if not self._original_settings:
            raise RuntimeError(
                "Cannot restore: no settings were saved. Call apply() first."
            )

        self._restore_original_settings()

    def get_current_settings(self) -> dict[str, object]:
        """Retourne l'état actuel des propriétés Excel.

        Returns:
            dict[str, object]: Dictionnaire {nom_propriété: valeur_actuelle}

        Example:
            >>> optimizer = ExcelOptimizer(app)
            >>> settings = optimizer.get_current_settings()
            >>> print(settings['ScreenUpdating'])
            True
        """
        try:
            return {
                "ScreenUpdating": self._app.ScreenUpdating,
                "DisplayStatusBar": self._app.DisplayStatusBar,
                "EnableAnimations": self._app.EnableAnimations,
                "Calculation": self._app.Calculation,
                "EnableEvents": self._app.EnableEvents,
                "DisplayAlerts": self._app.DisplayAlerts,
                "AskToUpdateLinks": self._app.AskToUpdateLinks,
                "Iteration": self._app.Iteration,
                "MaxIterations": self._app.MaxIterations,
                "MaxChange": self._app.MaxChange,
            }
        except Exception:
            # Si une propriété n'est pas accessible, retourner un dict vide
            return {}

    def _save_current_settings(self) -> None:
        """Sauvegarde les paramètres actuels d'Excel."""
        try:
            self._original_settings = {
                "ScreenUpdating": self._app.ScreenUpdating,
                "DisplayStatusBar": self._app.DisplayStatusBar,
                "EnableAnimations": self._app.EnableAnimations,
                "Calculation": self._app.Calculation,
                "EnableEvents": self._app.EnableEvents,
                "DisplayAlerts": self._app.DisplayAlerts,
                "AskToUpdateLinks": self._app.AskToUpdateLinks,
                "Iteration": self._app.Iteration,
                "MaxIterations": self._app.MaxIterations,
                "MaxChange": self._app.MaxChange,
            }
        except Exception:
            # En cas d'erreur, initialiser un dict vide
            self._original_settings = {}

    def _apply_optimizations(self) -> None:
        """Applique les optimisations de performance."""
        try:
            # Désactiver l'affichage
            self._app.ScreenUpdating = False
            self._app.DisplayStatusBar = False
            self._app.EnableAnimations = False

            # Passer en calcul manuel
            self._app.Calculation = -4135  # xlCalculationManual

            # Désactiver les événements
            self._app.EnableEvents = False
            self._app.DisplayAlerts = False
            self._app.AskToUpdateLinks = False

            # Désactiver l'itération
            self._app.Iteration = False
        except Exception:
            # Ignorer les erreurs d'application
            pass

    def _restore_original_settings(self) -> None:
        """Restaure les paramètres originaux d'Excel."""
        if not self._original_settings:
            return

        try:
            for prop, value in self._original_settings.items():
                setattr(self._app, prop, value)
        except Exception:
            # Ignorer les erreurs de restauration
            pass

        # Vider le cache des paramètres
        self._original_settings = {}
