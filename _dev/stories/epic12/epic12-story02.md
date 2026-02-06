# Epic 12 - Story 2: Implémenter MacroRunner pour l'exécution de macros VBA

**Statut** : ✅ Terminé

**Date de réalisation** : 2026-02-06

**En tant que** développeur
**Je veux** une classe MacroRunner capable d'exécuter des macros VBA
**Afin de** permettre aux utilisateurs de lancer des Sub et Function VBA avec arguments et récupération de retour

## Contexte

L'exécution de macros VBA se fait via `app.Run(macro_name, *args)`. Cette méthode COM peut :

1. Exécuter des Sub VBA (pas de valeur de retour)
2. Exécuter des Function VBA (avec valeur de retour)
3. Retourner des valeurs de différents types (str, int, float, date, tableau, etc.)
4. Lever des erreurs VBA runtime (capturées via `pywintypes.com_error`)

Le MacroRunner doit :

- Construire la référence complète de la macro (avec nom de classeur si nécessaire)
- Parser les arguments via `_parse_macro_args()` (Story 1)
- Exécuter la macro via `app.Run()`
- Capturer les erreurs COM et les traduire en `VBAMacroError`
- Encapsuler le résultat dans une dataclass `MacroResult`

## Critères d'acceptation

1. ✅ La dataclass `MacroResult` contient tous les champs nécessaires
2. ✅ La classe `MacroRunner` s'initialise avec un `ExcelManager`
3. ✅ La méthode `run()` exécute une macro et retourne un `MacroResult`
4. ✅ Les erreurs VBA runtime sont capturées et traduites en `VBAMacroError`
5. ✅ Les valeurs de retour VBA sont correctement typées
6. ✅ Les fonctions utilitaires `_build_macro_reference()` et `_format_return_value()` sont implémentées
7. ✅ Les tests couvrent l'exécution de Sub, Function, erreurs VBA

## Tâches techniques

### Tâche 2.1 : Créer la dataclass MacroResult

**Fichier** : `src/xlmanage/macro_runner.py`

Ajouter après les imports et avant les fonctions :

```python
from dataclasses import dataclass
from typing import Any


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
```

**Points d'attention** :
- `return_value` peut être `None` pour les Sub VBA (pas de return)
- `return_type` est le nom du type Python (ex: "str", "int", "tuple")
- `error_message` contient le message VBA extrait de `excepinfo[2]`

### Tâche 2.2 : Implémenter _build_macro_reference()

**Fichier** : `src/xlmanage/macro_runner.py`

```python
from pathlib import Path
from win32com.client import CDispatch

from xlmanage.exceptions import WorkbookNotFoundError


def _build_macro_reference(
    macro_name: str,
    workbook: Path | None,
    app: CDispatch
) -> str:
    """Construit la référence complète d'une macro VBA.

    Sans workbook : retourne macro_name tel quel (macro dans classeur actif ou PERSONAL.XLSB)
    Avec workbook : retourne "'WorkbookName.xlsm'!macro_name"

    Les guillemets simples sont nécessaires si le nom contient des espaces ou des points.

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
            message=f"Classeur '{workbook_name}' non ouvert - impossible d'exécuter la macro"
        )

    # Construire la référence avec guillemets simples si nécessaire
    # Format: 'WorkbookName.xlsm'!MacroName
    # Les guillemets sont obligatoires si le nom contient espaces, points, etc.
    return f"'{workbook_name}'!{macro_name}"
```

**Points d'attention** :
- Si `workbook=None`, la macro est cherchée dans le classeur actif ou PERSONAL.XLSB
- Les guillemets simples sont TOUJOURS utilisés autour du nom de classeur (simplifie)
- La recherche du classeur est case-insensitive mais on utilise le nom exact ensuite
- Si le classeur n'est pas ouvert, on lève `WorkbookNotFoundError`

### Tâche 2.3 : Implémenter _format_return_value()

**Fichier** : `src/xlmanage/macro_runner.py`

```python
import pywintypes


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
        dt = value  # pywintypes.datetime est compatible datetime
        return dt.isoformat()

    # Tableaux VBA (tuple de tuples)
    if isinstance(value, tuple) and value and isinstance(value[0], tuple):
        rows = len(value)
        cols = len(value[0]) if value else 0
        # Convertir en liste de listes pour affichage
        array_repr = [list(row) for row in value]
        return f"Tableau {rows}x{cols}: {array_repr}"

    # Default: conversion str
    return str(value)
```

**Points d'attention** :
- `pywintypes.TimeType` est le type retourné par Excel pour les dates
- Les tableaux VBA sont retournés comme `tuple[tuple[Any, ...], ...]`
- Pour les gros tableaux, cette fonction peut générer beaucoup de texte (acceptable)

### Tâche 2.4 : Implémenter la classe MacroRunner

**Fichier** : `src/xlmanage/macro_runner.py`

```python
from xlmanage.excel_manager import ExcelManager


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

    def __init__(self, excel_manager: ExcelManager) -> None:
        """Initialise le runner avec un ExcelManager.

        Args:
            excel_manager: Instance ExcelManager démarrée
        """
        self._mgr = excel_manager

    def run(
        self,
        macro_name: str,
        workbook: Path | None = None,
        args: str | None = None
    ) -> MacroResult:
        """Exécute une macro VBA avec arguments optionnels.

        Args:
            macro_name: Nom de la macro (ex: "Module1.MySub" ou "MySub")
            workbook: Classeur contenant la macro (None = classeur actif ou PERSONAL.XLSB)
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
                error_message=None
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
                    error_message=error_msg
                )
            else:
                # Autre erreur COM (macro introuvable, etc.)
                raise VBAMacroError(
                    macro_name=full_ref,
                    reason=f"Erreur COM (0x{hresult:08X}): {error_msg}"
                )
```

**Points d'attention** :
- `app.Run()` retourne `None` pour les Sub VBA (pas de valeur de retour)
- Les erreurs VBA runtime ont HRESULT 0x800A03EC ou 0x80020009
- Le message d'erreur VBA est dans `e.excepinfo[2]`
- Les autres HRESULT (macro introuvable, etc.) lèvent `VBAMacroError`

### Tâche 2.5 : Tests unitaires pour MacroRunner

**Fichier** : `tests/test_macro_runner.py`

```python
"""Tests pour l'exécution de macros VBA."""

import pytest
from unittest.mock import Mock, patch
from pathlib import Path
import pywintypes

from xlmanage.macro_runner import MacroRunner, MacroResult, _build_macro_reference, _format_return_value
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
        error_message=None
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
        error_message="Division by zero"
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
        "Module1.Test",
        Path("data.xlsm"),
        mock_excel_manager.app
    )

    assert ref == "'data.xlsm'!Module1.Test"


def test_build_macro_reference_workbook_not_found(mock_excel_manager):
    """Test erreur si classeur non ouvert."""
    mock_excel_manager.app.Workbooks = []

    with pytest.raises(WorkbookNotFoundError) as exc_info:
        _build_macro_reference(
            "Module1.Test",
            Path("missing.xlsm"),
            mock_excel_manager.app
        )

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
        0x800A03EC,
        "VBA error",
        (None, None, "Division by zero", None, None, 0),
        None
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
    com_error = pywintypes.com_error(
        0x80030000,  # HRESULT différent
        "Macro not found",
        None,
        None
    )
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
```

**Points d'attention** :
- Tester les Sub (retour None) et Function (avec retour)
- Tester les erreurs VBA runtime (HRESULT 0x800A03EC)
- Tester les erreurs COM génériques (macro introuvable)
- Vérifier que les arguments sont correctement parsés et passés

## Tests à implémenter

Tous les tests sont dans `tests/test_macro_runner.py` (17 tests).

**Coverage attendue** : > 90% pour MacroRunner et fonctions utilitaires

**Commande de test** :
```bash
pytest tests/test_macro_runner.py -v --cov=src/xlmanage/macro_runner --cov-report=term
```

## Dépendances

- Epic 5, Story 2 (ExcelManager)
- Epic 12, Story 1 (_parse_macro_args)
- Epic 6 (WorkbookManager pour WorkbookNotFoundError)

## Définition of Done

- [x] Dataclass `MacroResult` créée avec `__str__` implémenté
- [x] Fonction `_build_macro_reference()` implémentée
- [x] Fonction `_format_return_value()` implémentée
- [x] Classe `MacroRunner` implémentée avec méthode `run()`
- [x] Tous les tests passent (17 tests)
- [x] Couverture > 90% pour macro_runner.py
- [x] Les docstrings sont complètes avec exemples
- [x] mypy passe sans erreur (types annotés correctement)

## Notes pour le développeur junior

**Concepts clés à comprendre** :

1. **app.Run() retourne des types variés** :
   - `None` pour les Sub VBA
   - Types simples : int, float, str, bool
   - `pywintypes.TimeType` pour les dates
   - `tuple[tuple[...], ...]` pour les tableaux VBA

2. **HRESULT et excepinfo** :
   - HRESULT = code d'erreur Windows (hexadécimal)
   - `excepinfo` = tuple avec (source, description, helpfile, helpcontext, helpfile, scode)
   - Le message VBA est dans `excepinfo[2]`

3. **Référence de macro VBA** :
   - Sans workbook : `"MacroName"` → cherche dans actif + PERSONAL.XLSB
   - Avec workbook : `"'WorkbookName.xlsm'!MacroName"`
   - Guillemets simples obligatoires autour du nom de fichier

4. **Différence Sub vs Function** :
   - `Sub` : procédure, pas de return (retourne `None`)
   - `Function` : fonction, retourne une valeur

**Pièges à éviter** :

- ❌ Ne pas confondre erreur VBA runtime (0x800A03EC) et macro introuvable (autres HRESULT)
- ❌ Ne pas oublier de déréférencer `excepinfo[2]` pour le message d'erreur
- ❌ Ne pas utiliser `eval()` pour convertir les retours VBA
- ❌ Ne pas oublier que `app.Run(*args)` dépack les arguments

**Ressources** :

- [Application.Run Method](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run)
- [HRESULT Error Codes](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/0642cb2f-2075-4469-918c-4441e69c548a)
- [pywintypes.com_error](http://timgolden.me.uk/pywin32-docs/pywintypes__com_error_meth.html)
