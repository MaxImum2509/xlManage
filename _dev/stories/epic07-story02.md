# Epic 7 - Story 2: Implémenter la dataclass WorksheetInfo et les constantes de validation

**Statut** : ✅ Terminé

**En tant que** développeur
**Je veux** une structure de données pour représenter les informations d'une feuille
**Afin de** retourner des informations typées aux utilisateurs de l'API

## Critères d'acceptation

1. ✅ WorksheetInfo dataclass créée avec 5 champs
2. ✅ Constantes de validation définies (longueur max, caractères interdits)
3. ✅ Fonction `_validate_sheet_name()` implémentée
4. ✅ Tests unitaires couvrent toutes les validations et les cas limites

## Tâches techniques

### Tâche 2.1 : Créer le fichier worksheet_manager.py

**Fichier** : `src/xlmanage/worksheet_manager.py`

Commencer par les imports et la structure de base avec WorksheetInfo dataclass :

```python
from dataclasses import dataclass

@dataclass
class WorksheetInfo:
    """Information about an Excel worksheet.

    Attributes:
        name: Name of the worksheet (e.g., "Sheet1")
        index: Position in the workbook (1-based as in Excel)
        visible: Whether the worksheet is visible to the user
        rows_used: Number of rows containing data
        columns_used: Number of columns containing data
    """

    name: str
    index: int
    visible: bool
    rows_used: int
    columns_used: int
```

### Tâche 2.2 : Définir les constantes de validation

```python
# Excel worksheet name constraints
SHEET_NAME_MAX_LENGTH: int = 31
SHEET_NAME_FORBIDDEN_CHARS: str = r'\\/\*\?:\[\]'  # \ / * ? : [ ]
```

**Points d'attention** :
- Excel impose une limite stricte de 31 caractères pour les noms de feuilles
- Les caractères interdits sont : `\ / * ? : [ ]`
- La constante `SHEET_NAME_FORBIDDEN_CHARS` est une regex character class
- Ces règles sont imposées par Excel, pas par notre code

### Tâche 2.3 : Implémenter _validate_sheet_name()

```python
def _validate_sheet_name(name: str) -> None:
    """Validate an Excel worksheet name.

    Checks that the name follows Excel's naming rules:
    - Not empty
    - Maximum 31 characters
    - No forbidden characters: \\ / * ? : [ ]

    Args:
        name: The worksheet name to validate

    Raises:
        WorksheetNameError: If the name violates any rule

    Examples:
        >>> _validate_sheet_name("Sheet1")  # OK
        >>> _validate_sheet_name("Data-2024_Q1")  # OK
        >>> _validate_sheet_name("A" * 32)  # Raises: too long
        >>> _validate_sheet_name("Sheet/1")  # Raises: forbidden char
    """
    # Rule 1: Name cannot be empty
    if not name or not name.strip():
        raise WorksheetNameError(name, "name cannot be empty")

    # Rule 2: Maximum 31 characters
    if len(name) > SHEET_NAME_MAX_LENGTH:
        raise WorksheetNameError(
            name,
            f"name exceeds {SHEET_NAME_MAX_LENGTH} characters (length: {len(name)})"
        )

    # Rule 3: No forbidden characters
    forbidden_pattern = f"[{SHEET_NAME_FORBIDDEN_CHARS}]"
    match = re.search(forbidden_pattern, name)
    if match:
        forbidden_char = match.group(0)
        raise WorksheetNameError(
            name,
            f"contains forbidden character '{forbidden_char}'"
        )
```

## Dépendances

- Story 1 (Exceptions) - ✅ À créer avant

## Définition of Done

- [x] WorksheetInfo dataclass créée avec 5 champs
- [x] Constantes SHEET_NAME_MAX_LENGTH et SHEET_NAME_FORBIDDEN_CHARS définies
- [x] `_validate_sheet_name()` implémentée avec toutes les règles
- [x] Tous les tests passent (210 tests au total)
- [x] Couverture de code 100% pour la dataclass et la fonction
- [x] Documentation complète avec exemples
