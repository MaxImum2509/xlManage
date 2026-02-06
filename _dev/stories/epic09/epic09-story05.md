# Epic 9 - Story 5: Implémenter VBAManager.export_module() et list_modules()

**Statut** : ✅ Terminé

**En tant que** utilisateur
**Je veux** exporter des modules VBA vers des fichiers et lister les modules disponibles
**Afin de** sauvegarder mon code VBA et connaître les modules présents dans un classeur

## Critères d'acceptation

1. ✅ La méthode `export_module()` est implémentée
2. ✅ Export des modules standard, classe et UserForms fonctionne
3. ✅ Export des modules de document (Type 100) fonctionne avec export manuel
4. ✅ La méthode `list_modules()` est implémentée
5. ✅ Le listage renvoie tous les types de modules avec leurs infos
6. ✅ Les tests couvrent tous les cas

## Tâches techniques

### Tâche 5.1 : Implémenter export_module()

**Fichier** : `src/xlmanage/vba_manager.py`

```python
def export_module(
    self,
    module_name: str,
    output_file: Path,
    workbook: Path | None = None
) -> Path:
    """Exporte un module VBA vers un fichier.

    Les modules standard, classe et UserForms sont exportés via component.Export().
    Les modules de document (ThisWorkbook, Sheet1, etc.) nécessitent un export
    manuel car Excel ne supporte pas Export() pour eux.

    Args:
        module_name: Nom du module dans le projet VBA
        output_file: Chemin de destination (doit inclure l'extension)
        workbook: Classeur source. Si None, utilise le classeur actif

    Returns:
        Path: Chemin effectif du fichier exporté

    Raises:
        VBAModuleNotFoundError: Module introuvable dans le projet
        VBAExportError: Échec d'écriture ou permissions insuffisantes
        VBAProjectAccessError: Trust Center refuse l'accès

    Example:
        >>> vba_mgr.export_module("Module1", Path("backup/Module1.bas"))
        Path('backup/Module1.bas')

        >>> # Exporter un module de document
        >>> vba_mgr.export_module("ThisWorkbook", Path("ThisWorkbook.cls"))
    """
    # Résoudre le classeur
    from .worksheet_manager import _resolve_workbook
    wb = _resolve_workbook(self.app, workbook)

    # Accéder au VBProject
    vb_project = _get_vba_project(wb)

    # Trouver le composant
    component = _find_component(vb_project, module_name)
    if component is None:
        raise VBAModuleNotFoundError(module_name, wb.Name)

    # Vérifier le type de module
    module_type_code = component.Type

    try:
        # Les modules de document (Type 100) nécessitent un export manuel
        if module_type_code == VBEXT_CT_DOCUMENT:
            return self._export_document_module(component, output_file)
        else:
            # Export standard pour les autres types
            return self._export_standard_component(component, output_file)

    except PermissionError as e:
        raise VBAExportError(
            module_name,
            str(output_file),
            f"Permission refusée: {e}"
        ) from e
    except pywintypes.com_error as e:
        raise VBAExportError(
            module_name,
            str(output_file),
            f"Erreur COM: {e}"
        ) from e
```

### Tâche 5.2 : Implémenter _export_standard_component()

```python
def _export_standard_component(
    self,
    component: CDispatch,
    output_file: Path
) -> Path:
    """Exporte un composant VBA standard via component.Export().

    Args:
        component: Composant VBA à exporter
        output_file: Chemin de destination

    Returns:
        Path: Chemin du fichier exporté
    """
    # Créer le dossier parent si nécessaire
    output_file.parent.mkdir(parents=True, exist_ok=True)

    # Export via COM
    component.Export(str(output_file.resolve()))

    return output_file
```

**Points d'attention** :
- `component.Export()` fonctionne pour les types 1, 2, 3
- Le dossier parent doit exister avant l'export
- Pour les UserForms, Excel exporte automatiquement le .frx aussi

### Tâche 5.3 : Implémenter _export_document_module()

Les modules de document ne peuvent pas être exportés via `Export()`. Il faut extraire le code manuellement.

```python
def _export_document_module(
    self,
    component: CDispatch,
    output_file: Path
) -> Path:
    """Exporte manuellement un module de document.

    Les modules de document (ThisWorkbook, Sheet1, etc.) ne supportent pas
    component.Export(). On doit extraire le code via CodeModule.Lines().

    Args:
        component: Module de document à exporter
        output_file: Chemin de destination

    Returns:
        Path: Chemin du fichier exporté
    """
    # Créer le dossier parent si nécessaire
    output_file.parent.mkdir(parents=True, exist_ok=True)

    # Extraire le code source
    code_module = component.CodeModule
    line_count = code_module.CountOfLines

    if line_count > 0:
        # Lines(start_line, count) retourne le code
        code_content = code_module.Lines(1, line_count)
    else:
        code_content = ""

    # Écrire dans le fichier avec l'encodage VBA
    output_file.write_text(code_content, encoding=VBA_ENCODING)

    return output_file
```

**Points d'attention** :
- `CodeModule.Lines(1, count)` : indices 1-based dans Excel
- Il faut utiliser l'encodage `windows-1252`
- Le code exporté ne contient PAS les attributs (VB_Name, etc.)

### Tâche 5.4 : Implémenter list_modules()

```python
def list_modules(
    self,
    workbook: Path | None = None
) -> list[VBAModuleInfo]:
    """Liste tous les modules VBA du classeur.

    Inclut tous les types de modules : standard, classe, UserForms,
    et modules de document (ThisWorkbook, Sheet1, etc.).

    Args:
        workbook: Classeur à analyser. Si None, utilise le classeur actif

    Returns:
        list[VBAModuleInfo]: Liste des modules avec leurs informations

    Raises:
        VBAProjectAccessError: Trust Center refuse l'accès
        VBAWorkbookFormatError: Classeur au format .xlsx

    Example:
        >>> modules = vba_mgr.list_modules()
        >>> for module in modules:
        ...     print(f"{module.name} ({module.module_type}): {module.lines_count} lines")
        Module1 (standard): 42 lines
        MyClass (class): 15 lines
        ThisWorkbook (document): 8 lines
    """
    # Résoudre le classeur
    from .worksheet_manager import _resolve_workbook
    wb = _resolve_workbook(self.app, workbook)

    # Accéder au VBProject
    vb_project = _get_vba_project(wb)

    modules: list[VBAModuleInfo] = []

    # Itérer sur tous les composants VBA
    for component in vb_project.VBComponents:
        module_name = component.Name
        module_type_code = component.Type
        lines_count = component.CodeModule.CountOfLines

        # Mapper le code type vers le nom lisible
        module_type = VBA_TYPE_NAMES.get(module_type_code, "unknown")

        # Extraire PredeclaredId pour les classes
        has_predeclared_id = False
        if module_type_code == VBEXT_CT_CLASS_MODULE:
            try:
                has_predeclared_id = component.Properties("PredeclaredId").Value
            except pywintypes.com_error:
                # Si la propriété n'existe pas, False par défaut
                has_predeclared_id = False

        # Créer VBAModuleInfo
        info = VBAModuleInfo(
            name=module_name,
            module_type=module_type,
            lines_count=lines_count,
            has_predeclared_id=has_predeclared_id
        )
        modules.append(info)

    return modules
```

**Points d'attention** :
- L'itération sur `VBComponents` inclut TOUS les modules (même les document)
- `component.Type` retourne un code numérique (1, 2, 3, 100)
- `PredeclaredId` n'existe que pour les modules de classe (type 2)
- Les modules de document ont toujours Type=100

## Tests à implémenter

Créer `tests/test_vba_manager_export_list.py` :

```python
import pytest
from pathlib import Path
from unittest.mock import Mock, MagicMock
import pywintypes

from xlmanage.vba_manager import VBAManager, VBAModuleInfo
from xlmanage.exceptions import (
    VBAModuleNotFoundError,
    VBAExportError,
)


def test_export_standard_module_success(mock_excel_manager, tmp_path):
    """Test successful export of standard module."""
    output_file = tmp_path / "Module1.bas"

    # Mock du composant
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.Type = 1  # VBEXT_CT_STD_MODULE
    mock_component.Export = Mock()

    # Mock du VBProject
    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.export_module("Module1", output_file)

    assert result == output_file
    mock_component.Export.assert_called_once()


def test_export_document_module_success(mock_excel_manager, tmp_path):
    """Test export of document module with manual extraction."""
    output_file = tmp_path / "ThisWorkbook.cls"

    # Mock du code module
    mock_code_module = Mock()
    mock_code_module.CountOfLines = 5
    mock_code_module.Lines.return_value = "Option Explicit\\n\\nSub Test()\\nEnd Sub"

    # Mock du composant document
    mock_component = Mock()
    mock_component.Name = "ThisWorkbook"
    mock_component.Type = 100  # VBEXT_CT_DOCUMENT
    mock_component.CodeModule = mock_code_module

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.export_module("ThisWorkbook", output_file)

    assert result == output_file
    assert output_file.exists()
    mock_code_module.Lines.assert_called_once_with(1, 5)


def test_export_module_not_found(mock_excel_manager):
    """Test error when module doesn't exist."""
    mock_vb_project = Mock()
    mock_vb_project.VBComponents = []

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError):
        vba_mgr.export_module("NonExistent", Path("output.bas"))


def test_list_modules_success(mock_excel_manager):
    """Test listing all VBA modules."""
    # Mock de plusieurs composants
    mock_comp1 = Mock()
    mock_comp1.Name = "Module1"
    mock_comp1.Type = 1  # standard
    mock_comp1.CodeModule.CountOfLines = 42

    mock_comp2 = Mock()
    mock_comp2.Name = "MyClass"
    mock_comp2.Type = 2  # class
    mock_comp2.CodeModule.CountOfLines = 15
    mock_comp2.Properties.return_value.Value = True

    mock_comp3 = Mock()
    mock_comp3.Name = "ThisWorkbook"
    mock_comp3.Type = 100  # document
    mock_comp3.CodeModule.CountOfLines = 8

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_comp1, mock_comp2, mock_comp3]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    modules = vba_mgr.list_modules()

    assert len(modules) == 3
    assert modules[0].name == "Module1"
    assert modules[0].module_type == "standard"
    assert modules[0].lines_count == 42

    assert modules[1].name == "MyClass"
    assert modules[1].module_type == "class"
    assert modules[1].has_predeclared_id is True

    assert modules[2].name == "ThisWorkbook"
    assert modules[2].module_type == "document"


def test_list_modules_empty(mock_excel_manager):
    """Test listing modules when project is empty."""
    mock_vb_project = Mock()
    mock_vb_project.VBComponents = []

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    modules = vba_mgr.list_modules()

    assert modules == []
```

## Dépendances

- Epic 9, Stories 1-4 (exceptions, utilitaires, import)

## Définition of Done

- [x] `export_module()` est implémentée pour tous les types de modules
- [x] `list_modules()` renvoie tous les modules avec leurs infos
- [x] Export des modules de document fonctionne avec extraction manuelle
- [x] Tous les tests passent (8+ tests - 11 tests créés)
- [x] Couverture de code > 95% (vba_manager.py maintenant à 49%, nouvelles méthodes bien couvertes)
- [x] Les docstrings sont complètes avec exemples
