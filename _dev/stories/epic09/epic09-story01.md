# Epic 9 - Story 1: Créer les exceptions pour la gestion des modules VBA

**Statut** : ✅ Terminé

**En tant que** développeur
**Je veux** avoir des exceptions spécifiques pour les erreurs de gestion VBA
**Afin de** fournir des messages d'erreur clairs quand l'accès VBA échoue ou qu'un module est invalide

## Critères d'acceptation

1. ✅ Sept nouvelles exceptions VBA sont créées dans `src/xlmanage/exceptions.py`
2. ✅ Toutes héritent de `ExcelManageError`
3. ✅ Chaque exception a des attributs métier appropriés
4. ✅ Les exceptions sont exportées dans `__init__.py`
5. ✅ Les tests couvrent tous les cas d'usage

## Tâches techniques

### Tâche 1.1 : Créer VBAProjectAccessError

**Fichier** : `src/xlmanage/exceptions.py`

```python
class VBAProjectAccessError(ExcelManageError):
    """Accès au projet VBA refusé par le Trust Center.

    Raised when Excel's Trust Center blocks programmatic access to VBA.
    """

    def __init__(self, workbook_name: str):
        """Initialize VBA project access error.

        Args:
            workbook_name: Name of the workbook with blocked VBA access
        """
        self.workbook_name = workbook_name
        super().__init__(
            f"Access to VBA project in '{workbook_name}' denied. "
            "Enable 'Trust access to the VBA project object model' in Excel Trust Center."
        )
```

**Points d'attention** :
- Cette erreur se produit quand on tente d'accéder à `wb.VBProject`
- Le HRESULT COM typique est 0x800A03EC
- L'utilisateur doit activer l'option manuellement : File > Options > Trust Center > Trust Center Settings > Macro Settings

### Tâche 1.2 : Créer VBAModuleNotFoundError

```python
class VBAModuleNotFoundError(ExcelManageError):
    """Module VBA introuvable dans le projet.

    Raised when attempting to access a VBA module that doesn't exist.
    """

    def __init__(self, module_name: str, workbook_name: str):
        """Initialize VBA module not found error.

        Args:
            module_name: Name of the missing module
            workbook_name: Name of the workbook that was searched
        """
        self.module_name = module_name
        self.workbook_name = workbook_name
        super().__init__(
            f"VBA module '{module_name}' not found in workbook '{workbook_name}'"
        )
```

### Tâche 1.3 : Créer VBAModuleAlreadyExistsError

```python
class VBAModuleAlreadyExistsError(ExcelManageError):
    """Module VBA avec ce nom existe déjà.

    Raised when attempting to import a module with a duplicate name.
    """

    def __init__(self, module_name: str, workbook_name: str):
        """Initialize VBA module already exists error.

        Args:
            module_name: Name of the duplicate module
            workbook_name: Name of the workbook containing the duplicate
        """
        self.module_name = module_name
        self.workbook_name = workbook_name
        super().__init__(
            f"VBA module '{module_name}' already exists in workbook '{workbook_name}'"
        )
```

### Tâche 1.4 : Créer VBAImportError

```python
class VBAImportError(ExcelManageError):
    """Échec d'import de module VBA.

    Raised when importing a VBA module fails (invalid file, wrong encoding, etc.).
    """

    def __init__(self, module_file: str, reason: str):
        """Initialize VBA import error.

        Args:
            module_file: Path to the module file that failed to import
            reason: Explanation of why the import failed
        """
        self.module_file = module_file
        self.reason = reason
        super().__init__(f"Failed to import VBA module from '{module_file}': {reason}")
```

**Points d'attention** :
- Les fichiers VBA doivent être encodés en `windows-1252` (pas UTF-8)
- Les modules de classe (.cls) nécessitent un parsing spécial des attributs
- Les UserForms (.frm) doivent avoir leur fichier .frx associé dans le même dossier

### Tâche 1.5 : Créer VBAExportError

```python
class VBAExportError(ExcelManageError):
    """Échec d'export de module VBA.

    Raised when exporting a VBA module fails (permissions, invalid path, etc.).
    """

    def __init__(self, module_name: str, output_path: str, reason: str):
        """Initialize VBA export error.

        Args:
            module_name: Name of the module that failed to export
            output_path: Destination path where export was attempted
            reason: Explanation of why the export failed
        """
        self.module_name = module_name
        self.output_path = output_path
        self.reason = reason
        super().__init__(
            f"Failed to export VBA module '{module_name}' to '{output_path}': {reason}"
        )
```

### Tâche 1.6 : Créer VBAMacroError

```python
class VBAMacroError(ExcelManageError):
    """Échec d'exécution de macro VBA.

    Raised when a VBA macro execution fails or the macro is not found.
    """

    def __init__(self, macro_name: str, reason: str):
        """Initialize VBA macro error.

        Args:
            macro_name: Name of the macro that failed
            reason: Explanation of the failure (from COM excepinfo[2])
        """
        self.macro_name = macro_name
        self.reason = reason
        super().__init__(f"Macro '{macro_name}' failed: {reason}")
```

### Tâche 1.7 : Créer VBAWorkbookFormatError

```python
class VBAWorkbookFormatError(ExcelManageError):
    """Classeur au format .xlsx ne supportant pas les macros.

    Raised when attempting VBA operations on a macro-free workbook format.
    """

    def __init__(self, workbook_name: str):
        """Initialize VBA workbook format error.

        Args:
            workbook_name: Name of the .xlsx workbook
        """
        self.workbook_name = workbook_name
        super().__init__(
            f"Workbook '{workbook_name}' is in .xlsx format which doesn't support VBA. "
            "Convert to .xlsm format to use macros."
        )
```

**Points d'attention** :
- Seuls les formats .xlsm, .xlsb et .xls supportent les macros
- Le format .xlsx ne peut pas contenir de VBProject
- Il faut vérifier le format avant toute opération VBA

### Tâche 1.8 : Exporter les exceptions dans __init__.py

**Fichier** : `src/xlmanage/__init__.py`

Ajouter à la liste `__all__` :
```python
"VBAProjectAccessError",
"VBAModuleNotFoundError",
"VBAModuleAlreadyExistsError",
"VBAImportError",
"VBAExportError",
"VBAMacroError",
"VBAWorkbookFormatError",
```

Et importer depuis exceptions :
```python
from .exceptions import (
    # ... existantes
    VBAProjectAccessError,
    VBAModuleNotFoundError,
    VBAModuleAlreadyExistsError,
    VBAImportError,
    VBAExportError,
    VBAMacroError,
    VBAWorkbookFormatError,
)
```

## Tests à implémenter

Créer `tests/test_vba_exceptions.py` :

```python
import pytest
from xlmanage.exceptions import (
    VBAProjectAccessError,
    VBAModuleNotFoundError,
    VBAModuleAlreadyExistsError,
    VBAImportError,
    VBAExportError,
    VBAMacroError,
    VBAWorkbookFormatError,
)


def test_vba_project_access_error():
    """Test VBAProjectAccessError attributes and message."""
    error = VBAProjectAccessError("test.xlsm")
    assert error.workbook_name == "test.xlsm"
    assert "Trust access" in str(error)


def test_vba_module_not_found_error():
    """Test VBAModuleNotFoundError attributes and message."""
    error = VBAModuleNotFoundError("Module1", "test.xlsm")
    assert error.module_name == "Module1"
    assert error.workbook_name == "test.xlsm"
    assert "not found" in str(error)


def test_vba_module_already_exists_error():
    """Test VBAModuleAlreadyExistsError attributes and message."""
    error = VBAModuleAlreadyExistsError("Module1", "test.xlsm")
    assert error.module_name == "Module1"
    assert error.workbook_name == "test.xlsm"
    assert "already exists" in str(error)


def test_vba_import_error():
    """Test VBAImportError attributes and message."""
    error = VBAImportError("Module1.bas", "Invalid encoding")
    assert error.module_file == "Module1.bas"
    assert error.reason == "Invalid encoding"
    assert "Failed to import" in str(error)


def test_vba_export_error():
    """Test VBAExportError attributes and message."""
    error = VBAExportError("Module1", "C:\\output.bas", "Permission denied")
    assert error.module_name == "Module1"
    assert error.output_path == "C:\\output.bas"
    assert error.reason == "Permission denied"
    assert "Failed to export" in str(error)


def test_vba_macro_error():
    """Test VBAMacroError attributes and message."""
    error = VBAMacroError("MySub", "Runtime error '9': Subscript out of range")
    assert error.macro_name == "MySub"
    assert error.reason == "Runtime error '9': Subscript out of range"
    assert "failed" in str(error)


def test_vba_workbook_format_error():
    """Test VBAWorkbookFormatError attributes and message."""
    error = VBAWorkbookFormatError("data.xlsx")
    assert error.workbook_name == "data.xlsx"
    assert ".xlsx format" in str(error)
    assert ".xlsm" in str(error)
```

## Dépendances

- Aucune dépendance (exceptions de base)

## Définition of Done

- [x] Les 7 exceptions sont créées avec docstrings complètes
- [x] Les exceptions sont exportées dans `__init__.py`
- [x] Tous les tests passent (7 tests)
- [x] Couverture de code 100% pour les nouvelles exceptions
- [x] Le code suit les conventions du projet (type hints, docstrings Google style)
