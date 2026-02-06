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

from xlmanage.exceptions import VBAMacroError

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
