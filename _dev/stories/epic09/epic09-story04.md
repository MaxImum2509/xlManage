# Epic 9 - Story 4: Implémenter VBAManager.import_module()

**Statut** : ⏳ À faire

**En tant que** utilisateur
**Je veux** importer des modules VBA depuis des fichiers (.bas, .cls, .frm)
**Afin de** ajouter du code VBA à mes classeurs Excel de manière programmatique

## Critères d'acceptation

1. ⬜ La méthode `import_module()` est implémentée
2. ⬜ Import des modules standard (.bas) fonctionne
3. ⬜ Import des modules de classe (.cls) avec parsing fonctionne
4. ⬜ Import des UserForms (.frm) fonctionne
5. ⬜ Le paramètre `overwrite` permet de remplacer un module existant
6. ⬜ Les erreurs sont correctement gérées (Trust Center, format, etc.)
7. ⬜ Les tests couvrent tous les cas (succès + erreurs)

## Tâches techniques

### Tâche 4.1 : Implémenter import_module() - structure de base

**Fichier** : `src/xlmanage/vba_manager.py`

```python
def import_module(
    self,
    module_file: Path,
    module_type: str | None = None,
    workbook: Path | None = None,
    overwrite: bool = False
) -> VBAModuleInfo:
    """Importe un module VBA depuis un fichier.

    Supporte les modules standard (.bas), de classe (.cls) et UserForms (.frm).
    Les modules de classe nécessitent un traitement spécial pour extraire
    les attributs VB_Name et VB_PredeclaredId.

    Args:
        module_file: Chemin du fichier .bas, .cls ou .frm à importer
        module_type: Type forcé du module. Si None, détection auto depuis l'extension
        workbook: Classeur cible. Si None, utilise le classeur actif
        overwrite: Si True, supprime le module existant avant import

    Returns:
        VBAModuleInfo: Informations sur le module importé

    Raises:
        VBAImportError: Fichier invalide, encodage incorrect, ou .frx manquant
        VBAModuleAlreadyExistsError: Module existe et overwrite=False
        VBAProjectAccessError: Trust Center refuse l'accès au VBProject
        VBAWorkbookFormatError: Classeur au format .xlsx
        WorkbookNotFoundError: Classeur non ouvert

    Example:
        >>> vba_mgr.import_module(Path("Module1.bas"))
        VBAModuleInfo(name='Module1', module_type='standard', ...)

        >>> # Importer avec remplacement
        >>> vba_mgr.import_module(Path("MyClass.cls"), overwrite=True)
    """
    # Vérifier que le fichier existe
    if not module_file.exists():
        raise VBAImportError(str(module_file), "Fichier introuvable")

    # Détection automatique du type si non fourni
    if module_type is None:
        module_type = _detect_module_type(module_file)

    # Résoudre le classeur cible
    from .worksheet_manager import _resolve_workbook
    wb = _resolve_workbook(self.app, workbook)

    # Accéder au VBProject (raise si Trust Center bloque)
    vb_project = _get_vba_project(wb)

    # Router vers la méthode appropriée selon le type
    if module_type == "standard":
        return self._import_standard_module(vb_project, module_file, overwrite)
    elif module_type == "class":
        return self._import_class_module(vb_project, module_file, overwrite)
    elif module_type == "userform":
        return self._import_userform_module(vb_project, module_file, overwrite)
    else:
        raise VBAImportError(
            str(module_file),
            f"Type de module '{module_type}' non supporté"
        )
```

### Tâche 4.2 : Implémenter _import_standard_module()

```python
def _import_standard_module(
    self,
    vb_project: CDispatch,
    module_file: Path,
    overwrite: bool
) -> VBAModuleInfo:
    """Importe un module standard (.bas).

    Args:
        vb_project: Objet COM VBProject
        module_file: Chemin du fichier .bas
        overwrite: Si True, remplace le module existant

    Returns:
        VBAModuleInfo du module importé

    Raises:
        VBAModuleAlreadyExistsError: Si overwrite=False et module existe
        VBAImportError: Si l'import COM échoue
    """
    try:
        # Import direct via VBComponents.Import()
        component = vb_project.VBComponents.Import(str(module_file.resolve()))

        # Récupérer le nom du module importé
        module_name = component.Name

        # Vérifier si un module avec ce nom existe déjà (sauf si overwrite)
        if not overwrite:
            existing = _find_component(vb_project, module_name)
            if existing is not None and existing != component:
                # Annuler l'import (supprimer le module importé)
                vb_project.VBComponents.Remove(component)
                raise VBAModuleAlreadyExistsError(module_name, vb_project.Name)

        # Si overwrite et module existait, l'import l'a écrasé automatiquement

        # Construire VBAModuleInfo
        lines_count = component.CodeModule.CountOfLines
        return VBAModuleInfo(
            name=module_name,
            module_type="standard",
            lines_count=lines_count,
            has_predeclared_id=False
        )

    except pywintypes.com_error as e:
        raise VBAImportError(str(module_file), f"Erreur COM: {e}") from e
```

**Points d'attention** :
- `VBComponents.Import()` importe ET ajoute le module au projet
- Si un module du même nom existe, Excel l'écrase automatiquement
- Il faut vérifier l'existence AVANT l'import si `overwrite=False`

### Tâche 4.3 : Implémenter _import_class_module()

Les modules de classe nécessitent un traitement spécial car `Import()` ne gère pas correctement les attributs.

```python
def _import_class_module(
    self,
    vb_project: CDispatch,
    module_file: Path,
    overwrite: bool
) -> VBAModuleInfo:
    """Importe un module de classe (.cls) avec parsing des attributs.

    Les modules .cls contiennent des attributs (VB_Name, VB_PredeclaredId)
    qu'il faut extraire manuellement car Import() ne les gère pas correctement.

    Args:
        vb_project: Objet COM VBProject
        module_file: Chemin du fichier .cls
        overwrite: Si True, remplace le module existant

    Returns:
        VBAModuleInfo du module importé

    Raises:
        VBAModuleAlreadyExistsError: Si overwrite=False et module existe
        VBAImportError: Si le parsing échoue
    """
    # Parser le fichier .cls pour extraire les métadonnées
    module_name, predeclared_id, code_content = _parse_class_module(module_file)

    # Vérifier si le module existe déjà
    existing = _find_component(vb_project, module_name)
    if existing is not None:
        if not overwrite:
            raise VBAModuleAlreadyExistsError(module_name, vb_project.Name)
        # Supprimer l'ancien module
        vb_project.VBComponents.Remove(existing)
        del existing

    try:
        # Créer un nouveau module de classe (type 2)
        component = vb_project.VBComponents.Add(VBEXT_CT_CLASS_MODULE)

        # Définir le nom
        component.Name = module_name

        # Définir PredeclaredId si nécessaire
        if predeclared_id:
            component.Properties("PredeclaredId").Value = True

        # Ajouter le code source
        if code_content:
            component.CodeModule.AddFromString(code_content)

        # Construire VBAModuleInfo
        lines_count = component.CodeModule.CountOfLines
        return VBAModuleInfo(
            name=module_name,
            module_type="class",
            lines_count=lines_count,
            has_predeclared_id=predeclared_id
        )

    except pywintypes.com_error as e:
        raise VBAImportError(str(module_file), f"Erreur COM: {e}") from e
```

**Points d'attention** :
- `VBComponents.Add(2)` crée un module de classe vide
- Il faut définir le nom AVANT d'ajouter du code
- `PredeclaredId` se définit via `component.Properties("PredeclaredId")`
- `AddFromString()` ajoute le code à la fin du module

### Tâche 4.4 : Implémenter _import_userform_module()

```python
def _import_userform_module(
    self,
    vb_project: CDispatch,
    module_file: Path,
    overwrite: bool
) -> VBAModuleInfo:
    """Importe un UserForm (.frm + .frx).

    Args:
        vb_project: Objet COM VBProject
        module_file: Chemin du fichier .frm
        overwrite: Si True, remplace le UserForm existant

    Returns:
        VBAModuleInfo du UserForm importé

    Raises:
        VBAModuleAlreadyExistsError: Si overwrite=False et UserForm existe
        VBAImportError: Si le fichier .frx est manquant ou l'import échoue
    """
    # Vérifier que le fichier .frx existe (obligatoire pour les UserForms)
    frx_file = module_file.with_suffix('.frx')
    if not frx_file.exists():
        raise VBAImportError(
            str(module_file),
            f"Fichier .frx manquant : {frx_file}"
        )

    try:
        # Import direct via VBComponents.Import()
        component = vb_project.VBComponents.Import(str(module_file.resolve()))

        # Récupérer le nom du UserForm
        module_name = component.Name

        # Vérifier si un UserForm avec ce nom existe déjà
        if not overwrite:
            existing = _find_component(vb_project, module_name)
            if existing is not None and existing != component:
                # Annuler l'import
                vb_project.VBComponents.Remove(component)
                raise VBAModuleAlreadyExistsError(module_name, vb_project.Name)

        # Construire VBAModuleInfo
        lines_count = component.CodeModule.CountOfLines
        return VBAModuleInfo(
            name=module_name,
            module_type="userform",
            lines_count=lines_count,
            has_predeclared_id=True  # UserForms ont toujours PredeclaredId=True
        )

    except pywintypes.com_error as e:
        raise VBAImportError(str(module_file), f"Erreur COM: {e}") from e
```

**Points d'attention** :
- Les UserForms ont TOUJOURS deux fichiers : .frm (code) et .frx (layout binaire)
- Le fichier .frx doit être dans le même dossier que le .frm
- `Import()` charge automatiquement les deux fichiers

## Tests à implémenter

Créer `tests/test_vba_manager_import.py` :

```python
import pytest
from pathlib import Path
from unittest.mock import Mock, PropertyMock
import pywintypes

from xlmanage.vba_manager import VBAManager
from xlmanage.exceptions import (
    VBAImportError,
    VBAModuleAlreadyExistsError,
    VBAProjectAccessError,
)


def test_import_module_file_not_found(mock_excel_manager):
    """Test error when module file doesn't exist."""
    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAImportError, match="introuvable"):
        vba_mgr.import_module(Path("nonexistent.bas"))


def test_import_standard_module_success(mock_excel_manager, tmp_path):
    """Test successful import of .bas module."""
    # Créer un fichier .bas
    bas_file = tmp_path / "Module1.bas"
    bas_content = '''Attribute VB_Name = "Module1"
Sub Hello()
    MsgBox "Hello"
End Sub
'''
    bas_file.write_text(bas_content, encoding='windows-1252')

    # Mock du VBProject
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.CodeModule.CountOfLines = 4

    mock_vb_project = Mock()
    mock_vb_project.VBComponents.Import.return_value = mock_component
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.import_module(bas_file)

    assert result.name == "Module1"
    assert result.module_type == "standard"
    assert result.lines_count == 4
    mock_vb_project.VBComponents.Import.assert_called_once()


def test_import_class_module_success(mock_excel_manager, tmp_path):
    """Test successful import of .cls module."""
    cls_file = tmp_path / "MyClass.cls"
    cls_content = '''VERSION 1.0 CLASS
Attribute VB_Name = "MyClass"
Attribute VB_PredeclaredId = True
Option Explicit

Public Sub Test()
End Sub
'''
    cls_file.write_text(cls_content, encoding='windows-1252')

    # Mock du VBProject
    mock_component = Mock()
    mock_component.Name = "MyClass"
    mock_component.CodeModule.CountOfLines = 3
    mock_component.Properties = Mock(return_value=Mock())

    mock_vb_project = Mock()
    mock_vb_project.VBComponents.Add.return_value = mock_component
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    result = vba_mgr.import_module(cls_file)

    assert result.name == "MyClass"
    assert result.module_type == "class"
    assert result.has_predeclared_id is True


def test_import_module_already_exists_no_overwrite(mock_excel_manager, tmp_path):
    """Test error when module exists and overwrite=False."""
    bas_file = tmp_path / "Module1.bas"
    bas_content = 'Attribute VB_Name = "Module1"\\nSub Test()\\nEnd Sub'
    bas_file.write_text(bas_content, encoding='windows-1252')

    # Mock d'un module existant
    existing_component = Mock()
    existing_component.Name = "Module1"

    mock_vb_project = Mock()
    mock_vb_project.VBComponents.__iter__ = Mock(return_value=iter([existing_component]))

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleAlreadyExistsError):
        vba_mgr.import_module(bas_file, overwrite=False)


def test_import_userform_missing_frx(mock_excel_manager, tmp_path):
    """Test error when .frx file is missing for UserForm."""
    frm_file = tmp_path / "UserForm1.frm"
    frm_file.write_text("Dummy content", encoding='windows-1252')
    # Pas de fichier .frx

    mock_excel_manager.app.ActiveWorkbook.Name = "test.xlsm"

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAImportError, match=".frx manquant"):
        vba_mgr.import_module(frm_file)
```

## Dépendances

- Epic 9, Story 1 (exceptions)
- Epic 9, Story 2 (fonctions utilitaires)
- Epic 9, Story 3 (VBAManager init)
- Epic 6 (WorkbookManager pour _resolve_workbook)

## Définition of Done

- [ ] `import_module()` est implémentée avec tous les paramètres
- [ ] Import de modules .bas fonctionne
- [ ] Import de modules .cls avec parsing fonctionne
- [ ] Import de UserForms .frm/.frx fonctionne
- [ ] Le paramètre `overwrite` fonctionne correctement
- [ ] Tous les tests passent (10+ tests)
- [ ] Couverture de code > 95%
- [ ] Les docstrings sont complètes avec exemples
