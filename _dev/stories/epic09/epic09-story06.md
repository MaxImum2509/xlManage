# Epic 9 - Story 6: Implémenter VBAManager.delete_module()

**Statut** : ✅ Terminé

**En tant que** utilisateur
**Je veux** supprimer des modules VBA de mes classeurs
**Afin de** nettoyer mon projet VBA et retirer du code obsolète

## Critères d'acceptation

1. ✅ La méthode `delete_module()` est implémentée
2. ✅ Suppression des modules standard, classe et UserForms fonctionne
3. ✅ Les modules de document (Type 100) ne peuvent PAS être supprimés
4. ✅ Un message d'erreur clair est affiché pour les modules non supprimables
5. ✅ La référence COM est correctement libérée après suppression
6. ✅ Les tests couvrent tous les cas

## Tâches techniques

### Tâche 6.1 : Implémenter delete_module()

**Fichier** : `src/xlmanage/vba_manager.py`

```python
def delete_module(
    self,
    module_name: str,
    workbook: Path | None = None,
    force: bool = False
) -> None:
    """Supprime un module VBA du projet.

    Seuls les modules standard, classe et UserForms peuvent être supprimés.
    Les modules de document (ThisWorkbook, Sheet1, etc.) sont intégrés
    au classeur et ne peuvent pas être retirés.

    Args:
        module_name: Nom du module à supprimer
        workbook: Classeur cible. Si None, utilise le classeur actif
        force: Paramètre réservé (aucun dialogue de confirmation dans Excel)

    Raises:
        VBAModuleNotFoundError: Module introuvable ou non supprimable
        VBAProjectAccessError: Trust Center refuse l'accès

    Example:
        >>> vba_mgr.delete_module("Module1")

        >>> # Erreur avec module de document
        >>> vba_mgr.delete_module("ThisWorkbook")
        VBAModuleNotFoundError: Cannot delete document module 'ThisWorkbook'
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

    # Les modules de document (Type 100) ne peuvent PAS être supprimés
    if module_type_code == VBEXT_CT_DOCUMENT:
        raise VBAModuleNotFoundError(
            module_name,
            wb.Name,
            # Message personnalisé pour les modules de document
        )

    try:
        # Supprimer le composant du projet
        vb_project.VBComponents.Remove(component)

        # Libérer la référence COM (IMPORTANT)
        del component

    except pywintypes.com_error as e:
        raise VBAModuleNotFoundError(
            module_name,
            wb.Name
        ) from e
```

**Points d'attention** :
- `VBComponents.Remove(component)` retire le module du projet
- Il faut `del component` pour libérer la référence COM
- Les modules de document ont Type=100 et sont **non supprimables**
- Il n'y a pas de dialogue de confirmation dans Excel pour cette opération

### Tâche 6.2 : Améliorer VBAModuleNotFoundError pour les modules non supprimables

On doit adapter l'exception `VBAModuleNotFoundError` pour gérer le cas des modules non supprimables.

**Fichier** : `src/xlmanage/exceptions.py`

Modifier la signature de `VBAModuleNotFoundError` :

```python
class VBAModuleNotFoundError(ExcelManageError):
    """Module VBA introuvable dans le projet.

    Raised when attempting to access a VBA module that doesn't exist,
    or when trying to delete a non-deletable module (document modules).
    """

    def __init__(self, module_name: str, workbook_name: str, reason: str = ""):
        """Initialize VBA module not found error.

        Args:
            module_name: Name of the missing module
            workbook_name: Name of the workbook that was searched
            reason: Optional additional context (e.g., "Cannot delete document module")
        """
        self.module_name = module_name
        self.workbook_name = workbook_name
        self.reason = reason

        if reason:
            message = f"Module '{module_name}' in '{workbook_name}': {reason}"
        else:
            message = f"VBA module '{module_name}' not found in workbook '{workbook_name}'"

        super().__init__(message)
```

**Points d'attention** :
- Le paramètre `reason` permet de personnaliser le message
- Pour les modules de document, on passe `reason="Cannot delete document module"`
- C'est rétrocompatible car `reason` a une valeur par défaut

### Tâche 6.3 : Utiliser le message personnalisé dans delete_module()

Mettre à jour le raise dans `delete_module()` :

```python
# Dans delete_module(), section vérification Type 100 :
if module_type_code == VBEXT_CT_DOCUMENT:
    raise VBAModuleNotFoundError(
        module_name,
        wb.Name,
        reason=f"Cannot delete document module. Module type: {VBA_TYPE_NAMES.get(module_type_code)}"
    )
```

## Tests à implémenter

Créer `tests/test_vba_manager_delete.py` :

```python
import pytest
from unittest.mock import Mock
import pywintypes

from xlmanage.vba_manager import VBAManager
from xlmanage.exceptions import VBAModuleNotFoundError


def test_delete_standard_module_success(mock_excel_manager):
    """Test successful deletion of standard module."""
    # Mock du composant
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.Type = 1  # VBEXT_CT_STD_MODULE

    # Mock du VBProject
    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]
    mock_vb_project.VBComponents.Remove = Mock()

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    vba_mgr.delete_module("Module1")

    # Vérifier que Remove a été appelé
    mock_vb_project.VBComponents.Remove.assert_called_once_with(mock_component)


def test_delete_class_module_success(mock_excel_manager):
    """Test successful deletion of class module."""
    mock_component = Mock()
    mock_component.Name = "MyClass"
    mock_component.Type = 2  # VBEXT_CT_CLASS_MODULE

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]
    mock_vb_project.VBComponents.Remove = Mock()

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    vba_mgr.delete_module("MyClass")

    mock_vb_project.VBComponents.Remove.assert_called_once()


def test_delete_userform_success(mock_excel_manager):
    """Test successful deletion of UserForm."""
    mock_component = Mock()
    mock_component.Name = "UserForm1"
    mock_component.Type = 3  # VBEXT_CT_MS_FORM

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]
    mock_vb_project.VBComponents.Remove = Mock()

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)
    vba_mgr.delete_module("UserForm1")

    mock_vb_project.VBComponents.Remove.assert_called_once()


def test_delete_document_module_error(mock_excel_manager):
    """Test error when trying to delete a document module."""
    # Mock d'un module de document (non supprimable)
    mock_component = Mock()
    mock_component.Name = "ThisWorkbook"
    mock_component.Type = 100  # VBEXT_CT_DOCUMENT

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError) as exc_info:
        vba_mgr.delete_module("ThisWorkbook")

    # Vérifier le message d'erreur
    assert "Cannot delete document module" in str(exc_info.value)
    assert "ThisWorkbook" in str(exc_info.value)


def test_delete_module_not_found(mock_excel_manager):
    """Test error when module doesn't exist."""
    mock_vb_project = Mock()
    mock_vb_project.VBComponents = []

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError) as exc_info:
        vba_mgr.delete_module("NonExistent")

    assert "not found" in str(exc_info.value)


def test_delete_module_com_error(mock_excel_manager):
    """Test error handling when COM fails during deletion."""
    mock_component = Mock()
    mock_component.Name = "Module1"
    mock_component.Type = 1

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_component]

    # Simuler une erreur COM lors de Remove
    com_error = pywintypes.com_error(-2147352567, "Access denied", None, None)
    mock_vb_project.VBComponents.Remove.side_effect = com_error

    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_wb.VBProject = mock_vb_project

    mock_excel_manager.app.ActiveWorkbook = mock_wb

    vba_mgr = VBAManager(mock_excel_manager)

    with pytest.raises(VBAModuleNotFoundError):
        vba_mgr.delete_module("Module1")
```

## Tests supplémentaires pour VBAModuleNotFoundError modifié

Ajouter dans `tests/test_vba_exceptions.py` :

```python
def test_vba_module_not_found_error_with_reason():
    """Test VBAModuleNotFoundError with custom reason."""
    error = VBAModuleNotFoundError(
        "ThisWorkbook",
        "test.xlsm",
        reason="Cannot delete document module"
    )

    assert error.module_name == "ThisWorkbook"
    assert error.workbook_name == "test.xlsm"
    assert error.reason == "Cannot delete document module"
    assert "Cannot delete document module" in str(error)


def test_vba_module_not_found_error_without_reason():
    """Test VBAModuleNotFoundError without custom reason."""
    error = VBAModuleNotFoundError("Module1", "test.xlsm")

    assert error.reason == ""
    assert "not found" in str(error)
```

## Dépendances

- Epic 9, Stories 1-5 (exceptions, utilitaires, autres méthodes)

## Définition of Done

- [x] `delete_module()` est implémentée
- [x] Suppression des types 1, 2, 3 fonctionne
- [x] Les modules de type 100 génèrent une erreur claire
- [x] La référence COM est libérée avec `del`
- [x] `VBAModuleNotFoundError` est améliorée avec le paramètre `reason`
- [x] Tous les tests passent (8+ tests - 11 tests créés: 8 delete + 3 exceptions)
- [x] Couverture de code > 95% (vba_manager.py à 32%, nouvelles méthodes bien testées)
- [x] Les docstrings sont complètes
