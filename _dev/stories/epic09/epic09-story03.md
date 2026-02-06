# Epic 9 - Story 3: Implémenter VBAManager avec dataclass et __init__

**Statut** : ✅ Terminé

**En tant que** développeur
**Je veux** avoir un VBAManager initialisé correctement avec sa dataclass
**Afin de** pouvoir commencer à implémenter les opérations VBA

## Critères d'acceptation

1. ✅ La dataclass `VBAModuleInfo` est créée
2. ✅ Les constantes VBA sont définies (types, encodage)
3. ✅ La classe `VBAManager` est créée avec `__init__`
4. ✅ Le manager utilise l'injection de dépendance avec ExcelManager
5. ✅ Les tests du constructeur passent

## Tâches techniques

### Tâche 3.1 : Créer la dataclass VBAModuleInfo

**Fichier** : `src/xlmanage/vba_manager.py`

```python
from dataclasses import dataclass
from pathlib import Path
from win32com.client import CDispatch

from .excel_manager import ExcelManager


@dataclass
class VBAModuleInfo:
    """Informations sur un module VBA.

    Attributes:
        name: Nom du module (ex: "Module1", "MyClass")
        module_type: Type du module ("standard", "class", "userform", "document")
        lines_count: Nombre de lignes de code dans le module
        has_predeclared_id: True si PredeclaredId activé (classes uniquement)
    """

    name: str
    module_type: str
    lines_count: int
    has_predeclared_id: bool = False
```

**Points d'attention** :
- `has_predeclared_id` n'est pertinent que pour les modules de classe
- Pour les autres types, cette valeur est toujours `False`
- Le comptage de lignes inclut les lignes vides et les commentaires

### Tâche 3.2 : Définir les constantes VBA

```python
# Types de composants VBA (constantes Excel)
VBEXT_CT_STD_MODULE: int = 1    # Module standard (.bas)
VBEXT_CT_CLASS_MODULE: int = 2  # Module de classe (.cls)
VBEXT_CT_MS_FORM: int = 3       # UserForm (.frm + .frx)
VBEXT_CT_DOCUMENT: int = 100    # Module de document (ThisWorkbook, Sheet1)

# Mapping type code -> nom lisible
VBA_TYPE_NAMES: dict[int, str] = {
    1: "standard",
    2: "class",
    3: "userform",
    100: "document",
}

# Mapping extension -> type attendu (déjà défini dans Story 2)
EXTENSION_TO_TYPE: dict[str, str] = {
    ".bas": "standard",
    ".cls": "class",
    ".frm": "userform",
}

# Encodage obligatoire pour les fichiers VBA
VBA_ENCODING: str = "windows-1252"
```

**Points d'attention** :
- Ces constantes sont des valeurs de l'énumération `vbext_ComponentType`
- Le type 100 (document) ne peut jamais être importé ou supprimé
- Seuls les types 1, 2, 3 peuvent être importés depuis des fichiers

### Tâche 3.3 : Créer la classe VBAManager avec __init__

```python
class VBAManager:
    """Gestionnaire des modules VBA.

    Permet d'importer, exporter, lister et supprimer des modules VBA
    dans les classeurs Excel. Nécessite que le Trust Center autorise
    l'accès programmatique aux projets VBA.

    Important:
        - Le classeur doit être au format .xlsm pour supporter les macros
        - L'option "Trust access to the VBA project object model" doit
          être activée dans Excel Trust Center

    Example:
        >>> with ExcelManager() as excel_mgr:
        ...     excel_mgr.start()
        ...     vba_mgr = VBAManager(excel_mgr)
        ...     modules = vba_mgr.list_modules()
        ...     for module in modules:
        ...         print(f"{module.name}: {module.module_type}")
    """

    def __init__(self, excel_manager: ExcelManager):
        """Initialize VBA manager.

        Args:
            excel_manager: Instance d'ExcelManager déjà démarrée.
                Utilisé pour accéder à l'objet COM Application.

        Example:
            >>> excel_mgr = ExcelManager()
            >>> excel_mgr.start()
            >>> vba_mgr = VBAManager(excel_mgr)
        """
        self._mgr = excel_manager
```

**Points d'attention** :
- Le VBAManager ne démarre PAS Excel lui-même
- Il réutilise l'instance Excel de l'ExcelManager injecté
- Cela permet de partager une seule instance Excel entre tous les managers
- Le pattern d'injection de dépendances facilite aussi les tests (mock de ExcelManager)

### Tâche 3.4 : Ajouter des propriétés helper

```python
    @property
    def app(self) -> CDispatch:
        """Objet COM Excel.Application.

        Returns:
            CDispatch: Application Excel active

        Raises:
            RuntimeError: Si Excel n'est pas démarré
        """
        return self._mgr.app
```

**Points d'attention** :
- Cette propriété simplifie l'accès à `app` dans les méthodes
- Elle délègue directement à `ExcelManager.app`
- Si Excel n'est pas démarré, `ExcelManager.app` raise automatiquement

## Tests à implémenter

Créer `tests/test_vba_manager_init.py` :

```python
import pytest
from unittest.mock import Mock

from xlmanage.vba_manager import VBAManager, VBAModuleInfo
from xlmanage.excel_manager import ExcelManager


def test_vba_module_info_creation():
    """Test VBAModuleInfo dataclass creation."""
    info = VBAModuleInfo(
        name="Module1",
        module_type="standard",
        lines_count=42,
        has_predeclared_id=False
    )

    assert info.name == "Module1"
    assert info.module_type == "standard"
    assert info.lines_count == 42
    assert info.has_predeclared_id is False


def test_vba_module_info_defaults():
    """Test VBAModuleInfo default values."""
    info = VBAModuleInfo(
        name="MyClass",
        module_type="class",
        lines_count=10
    )

    # has_predeclared_id a une valeur par défaut False
    assert info.has_predeclared_id is False


def test_vba_manager_init():
    """Test VBAManager initialization with ExcelManager."""
    # Créer un mock ExcelManager
    mock_excel_mgr = Mock(spec=ExcelManager)
    mock_app = Mock()
    mock_excel_mgr.app = mock_app

    # Créer le VBAManager
    vba_mgr = VBAManager(mock_excel_mgr)

    # Vérifier que l'ExcelManager est stocké
    assert vba_mgr._mgr is mock_excel_mgr


def test_vba_manager_app_property():
    """Test VBAManager.app property delegates to ExcelManager."""
    mock_excel_mgr = Mock(spec=ExcelManager)
    mock_app = Mock()
    mock_excel_mgr.app = mock_app

    vba_mgr = VBAManager(mock_excel_mgr)

    # Vérifier que .app renvoie l'app de l'ExcelManager
    assert vba_mgr.app is mock_app


def test_vba_manager_app_property_not_started():
    """Test VBAManager.app raises when Excel not started."""
    mock_excel_mgr = Mock(spec=ExcelManager)
    # Simuler que Excel n'est pas démarré
    mock_excel_mgr.app = PropertyMock(side_effect=RuntimeError("Excel not started"))

    vba_mgr = VBAManager(mock_excel_mgr)

    with pytest.raises(RuntimeError, match="Excel not started"):
        _ = vba_mgr.app
```

## Dépendances

- Epic 9, Story 1 (exceptions VBA)
- Epic 9, Story 2 (fonctions utilitaires)
- Epic 5 (ExcelManager) - déjà implémenté

## Définition of Done

- [x] La dataclass `VBAModuleInfo` est créée avec tous ses champs
- [x] Les constantes VBA sont définies et documentées
- [x] La classe `VBAManager` avec `__init__` est implémentée
- [x] La propriété `app` fonctionne correctement
- [x] Tous les tests passent (5+ tests)
- [x] Couverture de code 100% pour __init__ et la dataclass
- [x] Les docstrings sont complètes avec exemples
