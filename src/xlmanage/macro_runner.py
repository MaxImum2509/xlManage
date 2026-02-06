"""
Execution de macros VBA avec parsing des arguments.

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

import re
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

import pywintypes
from win32com.client import CDispatch

from xlmanage.exceptions import VBAMacroError, WorkbookNotFoundError

if TYPE_CHECKING:
    from xlmanage.excel_manager import ExcelManager

# Limite COM pour le nombre d'arguments
MAX_MACRO_ARGS = 30


def _parse_macro_args(args_str: str) -> list[str | int | float | bool]:
    """Parse une chaîne CSV en liste d'arguments typés pour VBA.

    Les arguments sont convertis selon ces règles (dans l'ordre de priorité) :
    1. Chaînes entre guillemets ("..." ou '...') → str (sans les guillemets)
    2. "true" ou "false" (case-insensitive) → bool
    3. Nombre avec point décimal → float
    4. Nombre entier (avec signe optionnel) → int
    5. Tout le reste → str

    Exemples de parsing :
        '"hello, world",42,3.14,true' → ["hello, world", 42, 3.14, True]
        "'test',false,-100' → ["test", False, -100]
        '123,"abc",45.6' → [123, "abc", 45.6]

    Args:
        args_str: Chaîne CSV des arguments (ex: '"hello",42,3.14,true')

    Returns:
        list[Union[str, int, float, bool]]: Arguments parsés et typés

    Raises:
        VBAMacroError: Si > 30 arguments ou syntaxe CSV invalide

    Note:
        Les virgules dans les chaînes entre guillemets sont préservées.
        Les guillemets échappés dans les chaînes ne sont pas supportés.
    """
    if not args_str or not args_str.strip():
        return []

    # Splitter en respectant les guillemets
    # Regex pour découper : virgules hors guillemets
    # Pattern: correspond aux éléments entre virgules, en gérant les guillemets
    pattern = r"""
        (?:^|,)                    # Début de chaîne ou virgule
        \s*                        # Espaces optionnels
        (?:
            "([^"]*)"              # Chaîne entre guillemets doubles (group 1)
            |'([^']*)'             # OU chaîne entre guillemets simples (group 2)
            |([^,]+)               # OU valeur sans guillemets (group 3)
        )
        \s*                        # Espaces optionnels
    """

    matches = re.finditer(pattern, args_str, re.VERBOSE)
    raw_values: list[str] = []

    for match in matches:
        # Prendre le groupe non-None (double quote, single quote, ou sans quote)
        value = match.group(1) or match.group(2) or match.group(3)
        if value is not None:
            raw_values.append(value.strip())

    # Vérifier la limite COM
    if len(raw_values) > MAX_MACRO_ARGS:
        raise VBAMacroError(
            reason=(
                f"Trop d'arguments ({len(raw_values)}), "
                f"maximum autorisé : {MAX_MACRO_ARGS}"
            )
        )

    # Convertir chaque valeur selon son type
    typed_args: list[str | int | float | bool] = []

    for raw in raw_values:
        # 1. Bool (true/false case-insensitive)
        if raw.lower() == "true":
            typed_args.append(True)
            continue
        if raw.lower() == "false":
            typed_args.append(False)
            continue

        # 2. Float (contient un point décimal)
        if "." in raw:
            try:
                typed_args.append(float(raw))
                continue
            except ValueError:
                pass  # Pas un float valide, passer au suivant

        # 3. Int (nombre entier avec signe optionnel)
        if re.match(r"^[+-]?\d+$", raw):
            try:
                typed_args.append(int(raw))
                continue
            except ValueError:
                pass  # Pas un int valide, passer au suivant

        # 4. Default: str
        typed_args.append(raw)

    return typed_args


@dataclass
class MacroResult:
    """Résultat d'exécution d'une macro VBA.

    Attributes:
        macro_name: Nom complet de la macro exécutée (ex: "Module1.MySub")
        return_value: Valeur brute retournée par app.Run (None pour les Sub)
        return_type: Type Python du retour ("str", "int", "float", "NoneType", etc.)
        success: True si exécution sans erreur VBA
        error_message: Message d'erreur VBA si échec (None si succès)

    Example:
        >>> result = MacroResult(
        ...     macro_name="Module1.GetTotal",
        ...     return_value=42.5,
        ...     return_type="float",
        ...     success=True,
        ...     error_message=None
        ... )
    """

    macro_name: str
    return_value: Any | None
    return_type: str
    success: bool
    error_message: str | None

    def __str__(self) -> str:
        """Représentation textuelle du résultat."""
        if not self.success:
            return f"❌ {self.macro_name} - Erreur: {self.error_message}"

        if self.return_value is None:
            return f"✅ {self.macro_name} - Exécutée (pas de retour)"

        formatted_value = _format_return_value(self.return_value)
        return f"✅ {self.macro_name} - Retour ({self.return_type}): {formatted_value}"


def _build_macro_reference(
    macro_name: str, workbook: Path | None, app: CDispatch
) -> str:
    """Construit la référence complète d'une macro VBA.

    Sans workbook : retourne macro_name tel quel
    (macro dans classeur actif ou PERSONAL.XLSB)
    Avec workbook : retourne "'WorkbookName.xlsm'!macro_name"

    Les guillemets simples sont nécessaires si le nom contient
    des espaces ou des points.

    Args:
        macro_name: Nom de la macro (ex: "Module1.MySub" ou "MySub")
        workbook: Chemin du classeur contenant la macro (optionnel)
        app: Objet COM Excel.Application

    Returns:
        str: Référence complète de la macro (ex: "'data.xlsm'!Module1.MySub")

    Raises:
        WorkbookNotFoundError: Si le classeur spécifié n'est pas ouvert

    Example:
        >>> _build_macro_reference("MySub", None, app)
        "MySub"
        >>> _build_macro_reference("Module1.MySub", Path("data.xlsm"), app)
        "'data.xlsm'!Module1.MySub"
    """
    if workbook is None:
        # Pas de classeur spécifié, utiliser la macro telle quelle
        # Elle sera cherchée dans le classeur actif puis PERSONAL.XLSB
        return macro_name

    # Trouver le classeur ouvert
    workbook_name = workbook.name
    found = False

    for wb in app.Workbooks:
        if wb.Name.lower() == workbook_name.lower():
            found = True
            workbook_name = wb.Name  # Utiliser le nom exact (casse)
            break

    if not found:
        raise WorkbookNotFoundError(
            path=workbook,
            message=(
                f"Classeur '{workbook_name}' non ouvert - "
                "impossible d'exécuter la macro"
            ),
        )

    # Construire la référence avec guillemets simples si nécessaire
    # Format: 'WorkbookName.xlsm'!MacroName
    # Les guillemets sont obligatoires si le nom contient espaces, points, etc.
    return f"'{workbook_name}'!{macro_name}"


def _format_return_value(value: Any) -> str:
    """Formate une valeur de retour VBA pour affichage.

    Gère les cas spéciaux :
    - None → "(aucune valeur de retour)"
    - pywintypes.datetime → format ISO 8601
    - tuple de tuple (tableau VBA) → représentation tabulaire simplifiée
    - Autres → str(value)

    Args:
        value: Valeur retournée par app.Run()

    Returns:
        str: Représentation formatée pour affichage

    Example:
        >>> _format_return_value(None)
        "(aucune valeur de retour)"
        >>> _format_return_value(42)
        "42"
        >>> _format_return_value(((1, 2), (3, 4)))
        "Tableau 2x2: [[1, 2], [3, 4]]"
    """
    if value is None:
        return "(aucune valeur de retour)"

    # Dates VBA (pywintypes.datetime)
    if isinstance(value, pywintypes.TimeType):
        # Convertir en datetime Python puis formater ISO
        dt: pywintypes.TimeType = value
        return str(dt.isoformat())

    # Tableaux VBA (tuple de tuples)
    if isinstance(value, tuple) and value and isinstance(value[0], tuple):
        rows = len(value)
        cols = len(value[0]) if value else 0
        # Convertir en liste de listes pour affichage
        array_repr = [list(row) for row in value]
        return f"Tableau {rows}x{cols}: {array_repr}"

    # Default: conversion str
    return str(value)


class MacroRunner:
    """Exécuteur de macros VBA.

    Permet d'exécuter des Sub et Function VBA avec passage d'arguments
    et récupération de valeurs de retour.

    Attributes:
        _mgr: Instance ExcelManager pour l'accès à l'application Excel

    Example:
        >>> with ExcelManager() as mgr:
        ...     mgr.start()
        ...     runner = MacroRunner(mgr)
        ...     result = runner.run("Module1.MySub", args='"hello",42')
        ...     print(result.success)
        True
    """

    def __init__(self, excel_manager: "ExcelManager") -> None:
        """Initialise le runner avec un ExcelManager.

        Args:
            excel_manager: Instance ExcelManager démarrée
        """
        self._mgr = excel_manager

    def run(
        self, macro_name: str, workbook: Path | None = None, args: str | None = None
    ) -> MacroResult:
        """Exécute une macro VBA avec arguments optionnels.

        Args:
            macro_name: Nom de la macro (ex: "Module1.MySub" ou "MySub")
            workbook: Classeur contenant la macro
                (None = classeur actif ou PERSONAL.XLSB)
            args: Arguments CSV (ex: '"hello",42,3.14,true')

        Returns:
            MacroResult: Résultat d'exécution avec valeur de retour et statut

        Raises:
            VBAMacroError: Si parsing des arguments échoue ou macro introuvable
            WorkbookNotFoundError: Si le classeur n'est pas ouvert

        Example:
            >>> runner.run("Module1.GetSum", args="10,20")
            MacroResult(macro_name="Module1.GetSum", return_value=30, ...)

            >>> runner.run("Module1.SayHello", args='"World"')
            MacroResult(macro_name="Module1.SayHello", return_value=None, ...)
        """
        # 1. Construire la référence complète
        full_ref = _build_macro_reference(macro_name, workbook, self._mgr.app)

        # 2. Parser les arguments
        parsed_args: list[Any] = []
        if args:
            parsed_args = _parse_macro_args(args)

        # 3. Exécuter la macro
        try:
            return_value = self._mgr.app.Run(full_ref, *parsed_args)

            # Succès : construire le résultat
            return MacroResult(
                macro_name=full_ref,
                return_value=return_value,
                return_type=type(return_value).__name__,
                success=True,
                error_message=None,
            )

        except pywintypes.com_error as e:
            # Erreur COM : extraire le message VBA
            hresult = e.hresult
            error_msg = "Erreur VBA inconnue"

            # Le message d'erreur VBA est dans excepinfo[2]
            if e.excepinfo and len(e.excepinfo) > 2 and e.excepinfo[2]:
                error_msg = e.excepinfo[2]

            # HRESULT courants :
            # 0x800A03EC : Erreur générique Excel/VBA runtime
            # 0x80020009 : Exception avec excepinfo
            if hresult in (0x800A03EC, 0x80020009):
                # Erreur VBA runtime
                return MacroResult(
                    macro_name=full_ref,
                    return_value=None,
                    return_type="NoneType",
                    success=False,
                    error_message=error_msg,
                )
            else:
                # Autre erreur COM (macro introuvable, etc.)
                raise VBAMacroError(
                    macro_name=full_ref,
                    reason=f"Erreur COM (0x{hresult:08X}): {error_msg}",
                )
