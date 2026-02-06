# Epic 12 - Story 1: Parser des arguments de macros VBA

**Statut** : ✅ Terminé

**Date de réalisation** : 2026-02-06

**En tant que** développeur
**Je veux** un parser CSV pour les arguments de macros VBA
**Afin de** convertir une chaîne d'arguments en types Python compatibles COM

## Contexte

L'exécution d'une macro VBA via `app.Run()` nécessite de passer les arguments en types Python natifs (str, int, float, bool). Les utilisateurs fourniront les arguments en format CSV (ex: `"hello",42,3.14,true`).

Le parser doit :

1. Gérer les chaînes entre guillemets (simples ou doubles) avec des virgules internes
2. Convertir les types automatiquement (bool, int, float, str)
3. Respecter la limite COM de 30 arguments maximum
4. Valider la syntaxe et lever des erreurs claires

## Critères d'acceptation

1. ✅ La fonction `_parse_macro_args()` parse une chaîne CSV en liste d'arguments typés
2. ✅ Les chaînes entre guillemets sont détectées et les guillemets sont supprimés
3. ✅ Les types sont convertis dans l'ordre : bool → float → int → str
4. ✅ Les virgules dans les chaînes entre guillemets sont préservées
5. ✅ Une erreur est levée si > 30 arguments (limite COM)
6. ✅ Les tests couvrent tous les cas de conversion et les cas limites

## Tâches techniques

### Tâche 1.1 : Créer l'exception VBAMacroError

**Fichier** : `src/xlmanage/exceptions.py`

Ajouter cette exception au fichier (elle sera utilisée par le parser et par MacroRunner) :

```python
class VBAMacroError(ExcelManageError):
    """Échec d'exécution ou de parsing de macro VBA.

    Attributes:
        macro_name: Nom de la macro concernée (optionnel pour le parsing)
        reason: Description détaillée de l'erreur
    """

    def __init__(self, macro_name: str = "", reason: str = "") -> None:
        """Initialise l'exception.

        Args:
            macro_name: Nom de la macro (peut être vide pour les erreurs de parsing)
            reason: Raison de l'échec
        """
        self.macro_name = macro_name
        self.reason = reason
        message = f"Macro error"
        if macro_name:
            message += f" '{macro_name}'"
        if reason:
            message += f": {reason}"
        super().__init__(message)
```

**Points d'attention** :
- Ajouter `VBAMacroError` à `__all__` dans `exceptions.py`
- Cette exception sera utilisée pour les erreurs de parsing ET d'exécution
- Le `macro_name` est optionnel car le parser ne connaît pas encore le nom de la macro

### Tâche 1.2 : Implémenter _parse_macro_args()

**Fichier** : `src/xlmanage/macro_runner.py` (fonction au niveau module)

```python
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
from typing import Union

from xlmanage.exceptions import VBAMacroError


# Limite COM pour le nombre d'arguments
MAX_MACRO_ARGS = 30


def _parse_macro_args(args_str: str) -> list[Union[str, int, float, bool]]:
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
    pattern = r'''
        (?:^|,)                    # Début de chaîne ou virgule
        \s*                        # Espaces optionnels
        (?:
            "([^"]*)"              # Chaîne entre guillemets doubles (group 1)
            |'([^']*)'             # OU chaîne entre guillemets simples (group 2)
            |([^,]+)               # OU valeur sans guillemets (group 3)
        )
        \s*                        # Espaces optionnels
    '''

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
            reason=f"Trop d'arguments ({len(raw_values)}), maximum autorisé : {MAX_MACRO_ARGS}"
        )

    # Convertir chaque valeur selon son type
    typed_args: list[Union[str, int, float, bool]] = []

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
        if re.match(r'^[+-]?\d+$', raw):
            try:
                typed_args.append(int(raw))
                continue
            except ValueError:
                pass  # Pas un int valide, passer au suivant

        # 4. Default: str
        typed_args.append(raw)

    return typed_args
```

**Points d'attention** :
- La regex `(?:^|,)\s*(?:"([^"]*)"|'([^']*)'|([^,]+))\s*` gère les 3 cas :
  - Chaînes entre guillemets doubles : `"hello, world"`
  - Chaînes entre guillemets simples : `'test'`
  - Valeurs sans guillemets : `42`, `3.14`, `true`
- L'ordre de conversion est critique : bool avant float avant int avant str
- Les espaces autour des valeurs sont supprimés automatiquement
- La limite de 30 arguments est une contrainte COM (IDispatch::Invoke)

### Tâche 1.3 : Tests unitaires pour _parse_macro_args()

**Fichier** : `tests/test_macro_parser.py`

```python
"""Tests pour le parser d'arguments de macros VBA."""

import pytest
from xlmanage.macro_runner import _parse_macro_args
from xlmanage.exceptions import VBAMacroError


def test_parse_empty_args():
    """Test avec une chaîne vide."""
    assert _parse_macro_args("") == []
    assert _parse_macro_args("   ") == []


def test_parse_single_string():
    """Test avec une seule chaîne."""
    assert _parse_macro_args('"hello"') == ["hello"]
    assert _parse_macro_args("'world'") == ["world"]


def test_parse_string_with_comma():
    """Test chaîne avec virgule interne."""
    assert _parse_macro_args('"hello, world"') == ["hello, world"]
    assert _parse_macro_args("'a,b,c'") == ["a,b,c"]


def test_parse_integers():
    """Test conversion en int."""
    assert _parse_macro_args("42") == [42]
    assert _parse_macro_args("-100") == [-100]
    assert _parse_macro_args("+99") == [99]


def test_parse_floats():
    """Test conversion en float."""
    assert _parse_macro_args("3.14") == [3.14]
    assert _parse_macro_args("-0.5") == [-0.5]
    assert _parse_macro_args("123.456") == [123.456]


def test_parse_booleans():
    """Test conversion en bool."""
    assert _parse_macro_args("true") == [True]
    assert _parse_macro_args("false") == [False]
    assert _parse_macro_args("True") == [True]
    assert _parse_macro_args("FALSE") == [False]
    assert _parse_macro_args("TrUe") == [True]


def test_parse_mixed_types():
    """Test avec plusieurs types mélangés."""
    result = _parse_macro_args('"hello",42,3.14,true,"world",false,-10')
    assert result == ["hello", 42, 3.14, True, "world", False, -10]


def test_parse_complex_string():
    """Test chaîne complexe avec virgules et guillemets."""
    result = _parse_macro_args('"hello, world",42,"foo,bar,baz",3.14')
    assert result == ["hello, world", 42, "foo,bar,baz", 3.14]


def test_parse_with_spaces():
    """Test avec espaces autour des valeurs."""
    result = _parse_macro_args('  "hello"  ,  42  ,  3.14  ')
    assert result == ["hello", 42, 3.14]


def test_parse_unquoted_string():
    """Test chaîne sans guillemets (fallback)."""
    result = _parse_macro_args("hello,world")
    assert result == ["hello", "world"]


def test_parse_number_as_string():
    """Test nombre qui n'est pas valide → str."""
    result = _parse_macro_args("123abc")
    assert result == ["123abc"]


def test_parse_too_many_args():
    """Test erreur si > 30 arguments."""
    # Créer 31 arguments
    args_str = ",".join([str(i) for i in range(31)])

    with pytest.raises(VBAMacroError) as exc_info:
        _parse_macro_args(args_str)

    assert "31" in str(exc_info.value)
    assert "30" in str(exc_info.value)


def test_parse_edge_case_single_quote_in_double():
    """Test guillemet simple dans une chaîne entre guillemets doubles."""
    result = _parse_macro_args('''"it's working"''')
    assert result == ["it's working"]


def test_parse_edge_case_double_quote_in_single():
    """Test guillemet double dans une chaîne entre guillemets simples."""
    result = _parse_macro_args("""'he said "hello"'""")
    assert result == ['he said "hello"']


def test_parse_realistic_scenario():
    """Test scénario réaliste avec plusieurs types."""
    args = '''"Report_2024",100,"Sheet1",true,3.5,"C:\\Users\\test.xlsx"'''
    result = _parse_macro_args(args)

    assert result == [
        "Report_2024",
        100,
        "Sheet1",
        True,
        3.5,
        "C:\\Users\\test.xlsx"
    ]
    assert isinstance(result[0], str)
    assert isinstance(result[1], int)
    assert isinstance(result[2], str)
    assert isinstance(result[3], bool)
    assert isinstance(result[4], float)
    assert isinstance(result[5], str)
```

**Points d'attention** :
- Tester tous les types (str, int, float, bool)
- Tester les cas limites (chaînes vides, espaces, guillemets imbriqués)
- Vérifier la limite de 30 arguments
- S'assurer que les types retournés sont corrects (`isinstance`)

## Tests à implémenter

Tous les tests sont dans le fichier `tests/test_macro_parser.py` ci-dessus (17 tests).

**Coverage attendue** : > 95% pour `_parse_macro_args()`

**Commande de test** :
```bash
pytest tests/test_macro_parser.py -v --cov=src/xlmanage/macro_runner --cov-report=term
```

## Dépendances

- Epic 5 (ExcelManager) - non bloquant pour cette story car le parser est indépendant
- `exceptions.py` avec VBAMacroError

## Définition of Done

- [x] Exception `VBAMacroError` créée dans `exceptions.py`
- [x] Fonction `_parse_macro_args()` implémentée dans `macro_runner.py`
- [x] Tous les tests passent (17 tests)
- [x] Couverture > 95% pour le parser
- [x] Les docstrings sont complètes avec exemples
- [x] L'entête de licence GPL est présente dans `macro_runner.py`
- [x] Les types sont correctement annotés (mypy passe sans erreur)

## Notes pour le développeur junior

**Concepts clés à comprendre** :

1. **Regex avec VERBOSE mode** : Le flag `re.VERBOSE` permet d'écrire des regex multi-lignes avec commentaires pour la lisibilité.

2. **Match groups** : `match.group(1)`, `match.group(2)`, `match.group(3)` correspondent aux 3 cas (double quote, single quote, sans quote). Un seul sera non-None.

3. **Ordre de conversion** : Important ! bool avant float avant int avant str. Si on testait int avant float, "3.14" serait détecté comme "3" + ".14".

4. **Limite COM** : COM (Component Object Model) limite `IDispatch::Invoke` à 30 arguments maximum. C'est une contrainte Windows, pas Python.

5. **Type Union** : `Union[str, int, float, bool]` signifie "peut être l'un de ces 4 types". En Python 3.10+, on peut écrire `str | int | float | bool`.

**Pièges à éviter** :

- ❌ Ne pas utiliser `eval()` ou `ast.literal_eval()` : c'est dangereux et incorrect pour les booléens VBA
- ❌ Ne pas faire `split(",")` naïvement : ça casse les chaînes avec virgules
- ❌ Ne pas oublier le strip() des espaces autour des valeurs
- ❌ Ne pas inverser l'ordre de conversion des types

**Ressources** :

- [Regex Python documentation](https://docs.python.org/3/library/re.html)
- [IDispatch::Invoke limits](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke)
