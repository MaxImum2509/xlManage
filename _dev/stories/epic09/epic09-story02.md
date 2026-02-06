# Epic 9 - Story 2: Implémenter les fonctions utilitaires VBA

**Statut** : ✅ Terminé

**En tant que** développeur
**Je veux** avoir des fonctions utilitaires pour manipuler les projets VBA
**Afin de** faciliter l'implémentation du VBAManager

## Critères d'acceptation

1. ✅ Quatre fonctions utilitaires sont créées dans `vba_manager.py`
2. ✅ Chaque fonction gère correctement les erreurs COM
3. ✅ Les fonctions sont testées unitairement avec des mocks
4. ✅ La détection de type de module fonctionne correctement
5. ✅ L'accès au VBProject est sécurisé

## Tâches techniques

### Tâche 2.1 : Implémenter _get_vba_project()

**Fichier** : `src/xlmanage/vba_manager.py`

```python
import pywintypes
from win32com.client import CDispatch

from .exceptions import VBAProjectAccessError, VBAWorkbookFormatError


def _get_vba_project(wb: CDispatch) -> CDispatch:
    """Accède au VBProject avec gestion d'erreur.

    Args:
        wb: Objet COM Workbook

    Returns:
        CDispatch: Objet COM VBProject

    Raises:
        VBAProjectAccessError: Si Trust Center bloque l'accès
        VBAWorkbookFormatError: Si le classeur est en .xlsx
    """
    # Vérifier d'abord le format du classeur
    workbook_name = wb.Name
    if workbook_name.endswith('.xlsx'):
        raise VBAWorkbookFormatError(workbook_name)

    try:
        vb_project = wb.VBProject
        return vb_project
    except pywintypes.com_error as e:
        # HRESULT 0x800A03EC = Trust Center bloque l'accès
        if e.hresult == -2146827284:  # 0x800A03EC en signé
            raise VBAProjectAccessError(workbook_name) from e
        # Autre erreur COM inattendue
        raise
```

**Points d'attention** :
- Le HRESULT 0x800A03EC doit être converti en signé : `-2146827284`
- Il faut vérifier le format AVANT d'accéder au VBProject
- Cette fonction est appelée par toutes les méthodes du VBAManager

### Tâche 2.2 : Implémenter _find_component()

```python
def _find_component(vb_project: CDispatch, name: str) -> CDispatch | None:
    """Recherche un composant VBA par nom.

    Args:
        vb_project: Objet COM VBProject
        name: Nom du module à chercher

    Returns:
        CDispatch | None: Composant VBA trouvé, ou None si absent
    """
    try:
        # Itérer sur VBComponents
        for component in vb_project.VBComponents:
            if component.Name == name:
                return component
        return None
    except pywintypes.com_error:
        # En cas d'erreur COM, retourner None
        return None
```

**Points d'attention** :
- Les noms de modules VBA sont **case-sensitive**
- Ne pas lever d'exception ici, juste retourner None
- C'est au code appelant de raise VBAModuleNotFoundError si nécessaire

### Tâche 2.3 : Implémenter _detect_module_type()

```python
from pathlib import Path

from .exceptions import VBAImportError

# Constantes définies en haut du fichier
EXTENSION_TO_TYPE: dict[str, str] = {
    ".bas": "standard",
    ".cls": "class",
    ".frm": "userform",
}


def _detect_module_type(path: Path) -> str:
    """Détecte le type de module depuis l'extension.

    Args:
        path: Chemin du fichier module VBA

    Returns:
        str: Type du module ("standard", "class", "userform")

    Raises:
        VBAImportError: Si l'extension n'est pas reconnue
    """
    extension = path.suffix.lower()

    if extension not in EXTENSION_TO_TYPE:
        raise VBAImportError(
            str(path),
            f"Extension '{extension}' non reconnue. "
            f"Extensions valides : {', '.join(EXTENSION_TO_TYPE.keys())}"
        )

    return EXTENSION_TO_TYPE[extension]
```

**Points d'attention** :
- L'extension doit être convertie en minuscules (.BAS = .bas)
- Seules 3 extensions sont supportées pour l'import
- Les modules de document (Type=100) ne peuvent pas être importés

### Tâche 2.4 : Implémenter _parse_class_module()

Cette fonction est nécessaire car les modules de classe (.cls) nécessitent un traitement spécial.

```python
import re

def _parse_class_module(file_path: Path) -> tuple[str, bool, str]:
    """Parse un fichier .cls pour extraire les métadonnées.

    Les fichiers .cls commencent par des lignes "Attribute VB_Name" qu'il
    faut parser séparément avant d'importer le code.

    Args:
        file_path: Chemin du fichier .cls

    Returns:
        tuple[str, bool, str]: (module_name, predeclared_id, code_content)
            - module_name: Nom du module extrait de VB_Name
            - predeclared_id: True si VB_PredeclaredId = True
            - code_content: Code source sans les attributs d'en-tête

    Raises:
        VBAImportError: Si le fichier est invalide ou mal encodé
    """
    try:
        # Lire le fichier avec l'encodage VBA (OBLIGATOIRE)
        content = file_path.read_text(encoding='windows-1252')
    except UnicodeDecodeError as e:
        raise VBAImportError(
            str(file_path),
            f"Encodage invalide. Les fichiers VBA doivent être en windows-1252 : {e}"
        ) from e

    # Extraire VB_Name
    name_match = re.search(r'Attribute VB_Name = "([^"]+)"', content)
    if not name_match:
        raise VBAImportError(
            str(file_path),
            "Attribut VB_Name manquant dans le fichier .cls"
        )
    module_name = name_match.group(1)

    # Extraire VB_PredeclaredId (False par défaut)
    predeclared_match = re.search(r'Attribute VB_PredeclaredId = (True|False)', content)
    predeclared_id = predeclared_match.group(1) == "True" if predeclared_match else False

    # Extraire le code (tout après la dernière ligne Attribute)
    # On cherche "Option Explicit" ou la première ligne de code
    code_start = content.find("Option Explicit")
    if code_start == -1:
        # Pas de Option Explicit, chercher la première ligne non-Attribute
        lines = content.splitlines()
        for i, line in enumerate(lines):
            if not line.startswith("VERSION") and not line.startswith("Attribute"):
                code_start = sum(len(l) + 2 for l in lines[:i])  # +2 pour \r\n
                break

    if code_start == -1:
        code_content = ""
    else:
        code_content = content[code_start:].strip()

    return module_name, predeclared_id, code_content
```

**Points d'attention** :
- L'encodage `windows-1252` est **obligatoire** pour les fichiers VBA
- Les attributs VB_Name et VB_PredeclaredId sont en début de fichier
- Il faut séparer les attributs du code réel
- Cette fonction sera utilisée par `import_module()` dans la Story 3

## Tests à implémenter

Créer `tests/test_vba_utilities.py` :

```python
import pytest
from pathlib import Path
from unittest.mock import Mock, MagicMock
import pywintypes

from xlmanage.vba_manager import (
    _get_vba_project,
    _find_component,
    _detect_module_type,
    _parse_class_module,
)
from xlmanage.exceptions import (
    VBAProjectAccessError,
    VBAWorkbookFormatError,
    VBAImportError,
)


def test_get_vba_project_success(mocker):
    """Test successful VBProject access."""
    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"
    mock_vb_project = Mock()
    mock_wb.VBProject = mock_vb_project

    result = _get_vba_project(mock_wb)
    assert result == mock_vb_project


def test_get_vba_project_xlsx_format():
    """Test error when workbook is .xlsx format."""
    mock_wb = Mock()
    mock_wb.Name = "test.xlsx"

    with pytest.raises(VBAWorkbookFormatError) as exc_info:
        _get_vba_project(mock_wb)
    assert exc_info.value.workbook_name == "test.xlsx"


def test_get_vba_project_access_denied(mocker):
    """Test error when Trust Center blocks access."""
    mock_wb = Mock()
    mock_wb.Name = "test.xlsm"

    # Simuler l'erreur COM 0x800A03EC
    com_error = pywintypes.com_error(-2146827284, "Access denied", None, None)
    type(mock_wb).VBProject = PropertyMock(side_effect=com_error)

    with pytest.raises(VBAProjectAccessError) as exc_info:
        _get_vba_project(mock_wb)
    assert exc_info.value.workbook_name == "test.xlsm"


def test_find_component_found():
    """Test finding an existing component."""
    mock_comp1 = Mock()
    mock_comp1.Name = "Module1"
    mock_comp2 = Mock()
    mock_comp2.Name = "Module2"

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_comp1, mock_comp2]

    result = _find_component(mock_vb_project, "Module2")
    assert result == mock_comp2


def test_find_component_not_found():
    """Test component not found returns None."""
    mock_comp = Mock()
    mock_comp.Name = "Module1"

    mock_vb_project = Mock()
    mock_vb_project.VBComponents = [mock_comp]

    result = _find_component(mock_vb_project, "Module99")
    assert result is None


def test_detect_module_type_bas():
    """Test detection of .bas module."""
    assert _detect_module_type(Path("Module1.bas")) == "standard"
    assert _detect_module_type(Path("Module1.BAS")) == "standard"


def test_detect_module_type_cls():
    """Test detection of .cls module."""
    assert _detect_module_type(Path("Class1.cls")) == "class"


def test_detect_module_type_frm():
    """Test detection of .frm module."""
    assert _detect_module_type(Path("UserForm1.frm")) == "userform"


def test_detect_module_type_invalid():
    """Test error with invalid extension."""
    with pytest.raises(VBAImportError) as exc_info:
        _detect_module_type(Path("file.txt"))
    assert ".txt" in str(exc_info.value)


def test_parse_class_module_success(tmp_path):
    """Test parsing a valid .cls file."""
    cls_content = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Hello()
    MsgBox "Hello"
End Sub
'''
    cls_file = tmp_path / "MyClass.cls"
    cls_file.write_text(cls_content, encoding='windows-1252')

    name, predeclared, code = _parse_class_module(cls_file)

    assert name == "MyClass"
    assert predeclared is True
    assert "Option Explicit" in code
    assert "Public Sub Hello()" in code
    assert "Attribute" not in code


def test_parse_class_module_no_predeclared(tmp_path):
    """Test parsing .cls without PredeclaredId."""
    cls_content = '''Attribute VB_Name = "SimpleClass"
Option Explicit
'''
    cls_file = tmp_path / "SimpleClass.cls"
    cls_file.write_text(cls_content, encoding='windows-1252')

    name, predeclared, code = _parse_class_module(cls_file)

    assert name == "SimpleClass"
    assert predeclared is False


def test_parse_class_module_invalid_encoding(tmp_path):
    """Test error with wrong encoding."""
    cls_file = tmp_path / "bad.cls"
    cls_file.write_bytes(b'\xff\xfe' + "Invalid".encode('utf-16-le'))

    with pytest.raises(VBAImportError) as exc_info:
        _parse_class_module(cls_file)
    assert "windows-1252" in str(exc_info.value)


def test_parse_class_module_missing_vb_name(tmp_path):
    """Test error when VB_Name is missing."""
    cls_content = "Option Explicit\nPublic Sub Test()\nEnd Sub"
    cls_file = tmp_path / "bad.cls"
    cls_file.write_text(cls_content, encoding='windows-1252')

    with pytest.raises(VBAImportError) as exc_info:
        _parse_class_module(cls_file)
    assert "VB_Name manquant" in str(exc_info.value)
```

## Dépendances

- Epic 9, Story 1 (exceptions VBA)

## Définition of Done

- [x] Les 4 fonctions utilitaires sont implémentées avec docstrings complètes
- [x] Tous les tests passent (13 tests)
- [x] Couverture de code 83% pour les nouvelles fonctions
- [x] Les erreurs COM sont correctement interceptées et traduites
- [x] L'encodage windows-1252 est respecté
