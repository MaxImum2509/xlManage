# Epic 8 - Story 1: Créer les exceptions pour la gestion des tables

**Statut** : ✅ Implémentée

**En tant que** développeur
**Je veux** avoir des exceptions spécifiques pour les erreurs de gestion des tables
**Afin de** fournir des messages d'erreur clairs et exploitables aux utilisateurs

## Critères d'acceptation

1. ✅ Quatre nouvelles exceptions sont créées dans `src/xlmanage/exceptions.py`
2. ✅ Toutes héritent de `ExcelManageError`
3. ✅ Chaque exception a des attributs métier appropriés
4. ✅ Les exceptions sont exportées dans `__init__.py`
5. ✅ Les tests couvrent tous les cas d'usage

## Tâches techniques

### Tâche 1.1 : Créer TableNotFoundError

**Fichier** : `src/xlmanage/exceptions.py`

```python
class TableNotFoundError(ExcelManageError):
    """Table introuvable dans la feuille.

    Raised when attempting to access a table that doesn't exist.
    """

    def __init__(self, name: str, worksheet_name: str):
        """Initialize table not found error.

        Args:
            name: Name of the table that was not found
            worksheet_name: Name of the worksheet that was searched
        """
        self.name = name
        self.worksheet_name = worksheet_name
        super().__init__(f"Table '{name}' not found in worksheet '{worksheet_name}'")
```

**Points d'attention** :
- Les tables sont identifiées par leur nom (unique dans le classeur)
- On stocke le nom de la feuille pour le contexte
- Les noms de tables sont case-sensitive dans Excel

### Tâche 1.2 : Créer TableAlreadyExistsError

```python
class TableAlreadyExistsError(ExcelManageError):
    """Nom de table déjà utilisé.

    Raised when attempting to create a table with a name that already exists.
    """

    def __init__(self, name: str, workbook_name: str):
        """Initialize table already exists error.

        Args:
            name: Name of the table that already exists
            workbook_name: Name of the workbook (tables are unique per workbook)
        """
        self.name = name
        self.workbook_name = workbook_name
        super().__init__(f"Table '{name}' already exists in workbook '{workbook_name}'")
```

**Points d'attention** :
- Les noms de tables sont uniques **au niveau du classeur** (pas de la feuille)
- Contrairement aux feuilles, on ne peut pas avoir deux tables avec le même nom même dans des feuilles différentes

### Tâche 1.3 : Créer TableRangeError

```python
class TableRangeError(ExcelManageError):
    """Plage de table invalide.

    Raised when a table range is invalid (syntax error, empty, overlaps, etc.).
    """

    def __init__(self, range_ref: str, reason: str):
        """Initialize table range error.

        Args:
            range_ref: The invalid range reference (e.g., "A1:D10")
            reason: Explanation of why the range is invalid
        """
        self.range_ref = range_ref
        self.reason = reason
        super().__init__(f"Invalid table range '{range_ref}': {reason}")
```

### Tâche 1.4 : Créer TableNameError

```python
class TableNameError(ExcelManageError):
    """Nom de table invalide.

    Raised when a table name violates Excel naming rules.
    """

    def __init__(self, name: str, reason: str):
        """Initialize table name error.

        Args:
            name: The invalid table name
            reason: Explanation of why the name is invalid
        """
        self.name = name
        self.reason = reason
        super().__init__(f"Invalid table name '{name}': {reason}")
```

**Points d'attention** :
- Excel a des règles strictes pour les noms de tables :
  - Max 255 caractères
  - Doit commencer par une lettre ou underscore
  - Peut contenir lettres, chiffres, underscores
  - Pas d'espaces ni caractères spéciaux
  - Pas de nom de cellule (ex: "A1", "R1C1")

## Dépendances

- Aucune dépendance (exceptions de base)

## Définition of Done

- [x] Les 4 exceptions sont créées avec docstrings complètes
- [x] Les exceptions sont exportées dans `__init__.py`
- [x] Tous les tests passent (16 tests)
- [x] Couverture de code 100% pour les nouvelles exceptions
- [x] Le code suit les conventions du projet
