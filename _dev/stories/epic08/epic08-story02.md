# Epic 8 - Story 2: Implémenter la dataclass TableInfo et les constantes de validation

**Statut** : ✅ Implémentée

**En tant que** développeur
**Je veux** une structure de données pour représenter les informations d'une table
**Afin de** retourner des informations typées aux utilisateurs de l'API

## Critères d'acceptation

1. ✅ TableInfo dataclass créée avec 5 champs
2. ✅ Constantes de validation définies (longueur max, règles de nommage)
3. ✅ Fonction `_validate_table_name()` implémentée
4. ✅ Tests unitaires couvrent toutes les validations et les cas limites

## Tâches techniques

### Tâche 2.1 : Créer le fichier table_manager.py

**Fichier** : `src/xlmanage/table_manager.py`

```python
from dataclasses import dataclass

@dataclass
class TableInfo:
    """Information about an Excel table (ListObject).

    Attributes:
        name: Name of the table (e.g., "tbl_Sales")
        worksheet_name: Name of the worksheet containing the table
        range_ref: Range reference (e.g., "A1:D100")
        header_row_range: Range of the header row
        rows_count: Number of data rows (excluding header)
    """

    name: str
    worksheet_name: str
    range_ref: str
    header_row_range: str
    rows_count: int
```

### Tâche 2.2 : Définir les constantes de validation

```python
# Excel table name constraints
TABLE_NAME_MAX_LENGTH: int = 255
# Must start with letter or underscore, contains only alphanumeric and underscores
TABLE_NAME_PATTERN: str = r'^[a-zA-Z_][a-zA-Z0-9_]*$'
```

**Points d'attention** :
- Max 255 caractères pour les noms de tables
- Doit commencer par une lettre ou underscore
- Peut contenir lettres, chiffres, underscores uniquement
- Pas d'espaces ni caractères spéciaux

### Tâche 2.3 : Implémenter _validate_table_name()

```python
def _validate_table_name(name: str) -> None:
    """Validate an Excel table name.

    Checks that the name follows Excel's naming rules.

    Args:
        name: The table name to validate

    Raises:
        TableNameError: If the name violates any rule

    Examples:
        >>> _validate_table_name("tbl_Sales")  # OK
        >>> _validate_table_name("Data_2024")  # OK
        >>> _validate_table_name("A" * 256)  # Raises: too long
        >>> _validate_table_name("1Data")  # Raises: starts with digit
    """
    # Rule 1: Name cannot be empty
    if not name or not name.strip():
        raise TableNameError(name, "name cannot be empty")

    # Rule 2: Maximum 255 characters
    if len(name) > TABLE_NAME_MAX_LENGTH:
        raise TableNameError(
            name,
            f"name exceeds {TABLE_NAME_MAX_LENGTH} characters (length: {len(name)})"
        )

    # Rule 3: Must match pattern (start with letter or _, only alphanumeric and _)
    if not re.match(TABLE_NAME_PATTERN, name):
        raise TableNameError(
            name,
            "must start with letter or underscore, contain only alphanumeric characters and underscores"
        )

    # Rule 4: Cannot be a cell reference
    if re.match(r'^[A-Z]+\d+$|^[rR]\d+[cC]\d+$', name):
        raise TableNameError(name, "cannot be a cell reference")
```

## Dépendances

- Story 1 (Exceptions) - ✅ À créer avant

## Définition of Done

- [x] TableInfo dataclass créée avec 5 champs
- [x] Constantes TABLE_NAME_MAX_LENGTH et TABLE_NAME_PATTERN définies
- [x] `_validate_table_name()` implémentée avec toutes les règles
- [x] Tous les tests passent (20 tests)
- [x] Couverture de code 100% pour la dataclass et la fonction
- [x] Documentation complète avec exemples
