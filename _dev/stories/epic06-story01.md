# Epic 6 - Story 1 : Créer les exceptions pour la gestion des classeurs

## Vue d'ensemble

**En tant que** développeur
**Je veux** avoir des exceptions spécifiques pour les erreurs de gestion des classeurs
**Afin de** fournir des messages d'erreur clairs et exploitables aux utilisateurs

## Critères d'acceptation

1. ✅ Trois nouvelles exceptions sont créées dans `src/xlmanage/exceptions.py`
2. ✅ Toutes héritent de `ExcelManageError`
3. ✅ Chaque exception a des attributs métier appropriés
4. ✅ Les exceptions sont exportées dans `__init__.py`
5. ✅ Les tests couvrent tous les cas d'usage

## Tâches techniques

### Tâche 1.1 : Créer WorkbookNotFoundError

**Fichier** : `src/xlmanage/exceptions.py`

```python
class WorkbookNotFoundError(ExcelManageError):
    """Classeur introuvable sur le disque.

    Raised when attempting to open a workbook file that doesn't exist.
    """

    def __init__(self, path: Path, message: str = "Workbook not found"):
        """Initialize workbook not found error.

        Args:
            path: Path to the workbook file that was not found
            message: Human-readable error message
        """
        self.path = path
        self.message = message
        super().__init__(f"{message}: {path}")
```

**Points d'attention** :
- Importer `Path` depuis `pathlib` en haut du fichier
- L'attribut `path` est de type `Path`, pas `str`
- Le message parent inclut le chemin complet pour faciliter le débogage

### Tâche 1.2 : Créer WorkbookAlreadyOpenError

```python
class WorkbookAlreadyOpenError(ExcelManageError):
    """Classeur déjà ouvert dans l'instance Excel.

    Raised when attempting to open a workbook that is already open.
    """

    def __init__(self, path: Path, name: str, message: str = "Workbook already open"):
        """Initialize workbook already open error.

        Args:
            path: Path to the workbook file
            name: Name of the workbook (e.g., "data.xlsx")
            message: Human-readable error message
        """
        self.path = path
        self.name = name
        self.message = message
        super().__init__(f"{message}: {name} at {path}")
```

**Points d'attention** :
- On stocke à la fois `path` ET `name` car Excel peut avoir des classeurs avec le même nom dans des dossiers différents
- Le message doit être explicite pour que l'utilisateur sache quel fichier est déjà ouvert

### Tâche 1.3 : Créer WorkbookSaveError

```python
class WorkbookSaveError(ExcelManageError):
    """Échec de sauvegarde du classeur.

    Raised when save operation fails due to permissions, invalid path, or format issues.
    """

    def __init__(self, path: Path, hresult: int = 0, message: str = "Save failed"):
        """Initialize workbook save error.

        Args:
            path: Path where the save was attempted
            hresult: COM HRESULT error code (0 if not a COM error)
            message: Human-readable error message
        """
        self.path = path
        self.hresult = hresult
        self.message = message

        if hresult != 0:
            super().__init__(f"{message}: {path} (HRESULT: {hresult:#010x})")
        else:
            super().__init__(f"{message}: {path}")
```

**Points d'attention** :
- Le HRESULT est optionnel (défaut 0) car l'erreur peut venir du système de fichiers, pas du COM
- Si `hresult` est présent, on l'affiche en hexadécimal comme pour les exceptions COM

### Tâche 1.4 : Ajouter les imports nécessaires

En haut de `src/xlmanage/exceptions.py`, ajouter :

```python
from pathlib import Path
```

### Tâche 1.5 : Exporter les nouvelles exceptions

Dans `src/xlmanage/__init__.py`, ajouter à `__all__` :

```python
__all__ = [
    # ... existant ...
    "WorkbookNotFoundError",
    "WorkbookAlreadyOpenError",
    "WorkbookSaveError",
]
```

Et importer :

```python
from .exceptions import (
    # ... existant ...
    WorkbookNotFoundError,
    WorkbookAlreadyOpenError,
    WorkbookSaveError,
)
```

### Tâche 1.6 : Écrire les tests

**Fichier** : `tests/test_exceptions.py`

Ajouter 3 nouvelles classes de tests :

```python
class TestWorkbookNotFoundError:
    """Tests for WorkbookNotFoundError."""

    def test_workbook_not_found_default_message(self):
        """Test WorkbookNotFoundError with default message."""
        path = Path("C:/test/missing.xlsx")
        error = WorkbookNotFoundError(path)

        assert error.path == path
        assert error.message == "Workbook not found"
        assert "Workbook not found" in str(error)
        assert "missing.xlsx" in str(error)

    def test_workbook_not_found_custom_message(self):
        """Test WorkbookNotFoundError with custom message."""
        path = Path("D:/data/file.xlsm")
        error = WorkbookNotFoundError(path, "File does not exist")

        assert error.path == path
        assert error.message == "File does not exist"
        assert "File does not exist" in str(error)

    def test_workbook_not_found_inheritance(self):
        """Test WorkbookNotFoundError inherits from ExcelManageError."""
        error = WorkbookNotFoundError(Path("test.xlsx"))
        assert isinstance(error, ExcelManageError)
        assert isinstance(error, Exception)


class TestWorkbookAlreadyOpenError:
    """Tests for WorkbookAlreadyOpenError."""

    def test_workbook_already_open_default_message(self):
        """Test WorkbookAlreadyOpenError with default message."""
        path = Path("C:/test/open.xlsx")
        name = "open.xlsx"
        error = WorkbookAlreadyOpenError(path, name)

        assert error.path == path
        assert error.name == name
        assert error.message == "Workbook already open"
        assert name in str(error)

    def test_workbook_already_open_custom_message(self):
        """Test WorkbookAlreadyOpenError with custom message."""
        path = Path("D:/data/file.xlsm")
        name = "file.xlsm"
        error = WorkbookAlreadyOpenError(path, name, "Already loaded")

        assert error.message == "Already loaded"
        assert "Already loaded" in str(error)

    def test_workbook_already_open_inheritance(self):
        """Test WorkbookAlreadyOpenError inherits from ExcelManageError."""
        error = WorkbookAlreadyOpenError(Path("test.xlsx"), "test.xlsx")
        assert isinstance(error, ExcelManageError)


class TestWorkbookSaveError:
    """Tests for WorkbookSaveError."""

    def test_workbook_save_error_without_hresult(self):
        """Test WorkbookSaveError without COM error."""
        path = Path("C:/test/readonly.xlsx")
        error = WorkbookSaveError(path)

        assert error.path == path
        assert error.hresult == 0
        assert error.message == "Save failed"
        assert "readonly.xlsx" in str(error)
        assert "HRESULT" not in str(error)  # No HRESULT when 0

    def test_workbook_save_error_with_hresult(self):
        """Test WorkbookSaveError with COM error."""
        path = Path("D:/test/file.xlsx")
        error = WorkbookSaveError(path, hresult=0x80070005)

        assert error.hresult == 0x80070005
        assert "0x80070005" in str(error)
        assert "HRESULT" in str(error)

    def test_workbook_save_error_custom_message(self):
        """Test WorkbookSaveError with custom message."""
        path = Path("E:/data/protected.xlsx")
        error = WorkbookSaveError(path, message="Access denied")

        assert error.message == "Access denied"
        assert "Access denied" in str(error)

    def test_workbook_save_error_inheritance(self):
        """Test WorkbookSaveError inherits from ExcelManageError."""
        error = WorkbookSaveError(Path("test.xlsx"))
        assert isinstance(error, ExcelManageError)
```

**Commande de test** :
```bash
poetry run pytest tests/test_exceptions.py::TestWorkbookNotFoundError -v
poetry run pytest tests/test_exceptions.py::TestWorkbookAlreadyOpenError -v
poetry run pytest tests/test_exceptions.py::TestWorkbookSaveError -v
```

## Définition of Done

- [ ] Les 3 exceptions sont créées avec docstrings complètes
- [ ] Les exceptions sont exportées dans `__init__.py`
- [ ] Tous les tests passent (minimum 9 tests)
- [ ] Couverture de code 100% pour les nouvelles exceptions
- [ ] Le code suit les conventions du projet (pas d'emojis, docstrings en anglais)
