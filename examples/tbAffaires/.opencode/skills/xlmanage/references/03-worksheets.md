# WorksheetManager - CRUD Feuilles

## Import

```python
from xlmanage import WorksheetManager, WorksheetInfo
```

## WorksheetManager API

### Instanciation

```python
mgr = ExcelManager()
mgr.start()

ws_mgr = WorksheetManager(mgr.app)
```

### Méthodes CRUD

| Méthode | Paramètres | Retour | Description |
|----------|------------|--------|-------------|
| `create(name, workbook=None)` | `name: str, workbook: Path|None` | `WorksheetInfo` | Crée nouvelle feuille |
| `delete(name, workbook=None)` | `name: str, workbook: Path|None` | `None` | Supprime feuille |
| `copy(name, new_name, workbook=None)` | `name: str, new_name: str, workbook: Path|None` | `WorksheetInfo` | Copie feuille |
| `list(workbook=None)` | `workbook: Path|None` | `List[WorksheetInfo]` | Liste toutes les feuilles |

## Usage Patterns

### Créer une Nouvelle Feuille

```python
from pathlib import Path

with ExcelManager(visible=False) as mgr:
    mgr.start()
    ws_mgr = WorksheetManager(mgr.app)

    # Créer dans classeur actif
    info = ws_mgr.create("Data")
    print(f"Created sheet: {info.name} at index {info.index}")

    # Créer dans classeur spécifique
    info = ws_mgr.create("Summary", workbook=Path("data.xlsx"))
```

### Copier une Feuille

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    ws_mgr = WorksheetManager(mgr.app)

    # Copier dans classeur actif
    info = ws_mgr.copy("Template", "Copy1")
    print(f"Copied: {info.name}")

    # Copier dans classeur spécifique
    info = ws_mgr.copy("Template", "Backup",
                       workbook=Path("data.xlsx"))
```

### Supprimer une Feuille

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    ws_mgr = WorksheetManager(mgr.app)

    # Supprimer dans classeur actif
    ws_mgr.delete("TempSheet")

    # Supprimer dans classeur spécifique
    ws_mgr.delete("Draft", workbook=Path("data.xlsx"))
```

### Lister les Feuilles

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    ws_mgr = WorksheetManager(mgr.app)

    # Lister feuilles du classeur actif
    sheets = ws_mgr.list()

    for sheet in sheets:
        print(f"{sheet.name}: index={sheet.index}, "
              f"visible={sheet.visible}, "
              f"used={sheet.rows_used}x{sheet.columns_used}")

    # Lister feuilles d'un classeur spécifique
    sheets = ws_mgr.list(workbook=Path("data.xlsx"))
```

## WorksheetInfo Structure

```python
dataclass WorksheetInfo:
    name: str            # Nom de la feuille
    index: int           # Position dans le classeur (1-based)
    visible: bool         # Visibilité (True=visible, False=masquée)
    rows_used: int       # Nombre de lignes utilisées
    columns_used: int    # Nombre de colonnes utilisées
```

## Exceptions

| Exception | Condition | Contexte |
|-----------|-----------|-----------|
| `WorksheetAlreadyExistsError` | Nom déjà utilisé | `create()`, `copy()` |
| `WorksheetNotFoundError` | Feuille introuvable | `delete()`, `copy()` |
| `WorksheetNameError` | Nom invalide | `create()`, `copy()` |
| `WorksheetDeleteError` | Suppression impossible | `delete()` (ex: dernière feuille) |

## Workflow Complet

```python
from pathlib import Path
from xlmanage import (
    ExcelManager, WorksheetManager,
    WorksheetNotFoundError, WorksheetDeleteError
)

def setup_worksheets(workbook_path: Path):
    """Configure les feuilles d'un classeur"""
    with ExcelManager(visible=False) as mgr:
        mgr.start()
        ws_mgr = WorksheetManager(mgr.app)

        try:
            # Lister feuilles existantes
            sheets = ws_mgr.list(workbook=workbook_path)
            print(f"Existing sheets: {[s.name for s in sheets]}")

            # Créer nouvelle feuille
            data_sheet = ws_mgr.create("Data", workbook=workbook_path)
            print(f"Created: {data_sheet.name}")

            # Copier depuis template
            summary = ws_mgr.copy("Template", "Summary",
                                 workbook=workbook_path)
            print(f"Copied: {summary.name}")

            # Nettoyer feuilles temporaires
            for sheet in sheets:
                if sheet.name.startswith("_temp"):
                    try:
                        ws_mgr.delete(sheet.name, workbook=workbook_path)
                        print(f"Deleted: {sheet.name}")
                    except WorksheetDeleteError as e:
                        print(f"Cannot delete {sheet.name}: {e.reason}")

        except WorksheetAlreadyExistsError as e:
            print(f"Sheet '{e.name}' already exists")
        except WorksheetNameError as e:
            print(f"Invalid sheet name: {e.reason}")

# Usage
setup_worksheets(Path("data/workbook.xlsx"))
```

## COM Access Direct

Après création/récupération via WorksheetManager :

```python
info = ws_mgr.create("Data")
app = mgr.app

# Accéder au COM object de la feuille
ws = app.Worksheets(info.name)

# Manipulation directe COM
ws.Range("A1").Value = "Header"
ws.Range("A1:D100").Formula = "=ROW()"
```

## Trucs et Astuces

### Renommer une Feuille (via COM direct)

```python
# WorksheetManager n'a pas de méthode rename()
# Utiliser COM direct après avoir l'info feuille
info = ws_mgr.create("Temp")
app.Worksheets(info.name).Name = "Final"
```

### Masquer/Afficher une Feuille (via COM direct)

```python
ws = app.Worksheets("Sheet1")
ws.Visible = False          # False = masquée
ws.Visible = xlSheetVisible  # True = visible (-1)
```

### Vérifier Existence

```python
def sheet_exists(ws_mgr: WorksheetManager, name: str) -> bool:
    sheets = ws_mgr.list()
    return any(s.name == name for s in sheets)
```

## Documentation API

Pour voir l'API complète, utiliser les docstrings Python :

```python
from xlmanage import WorksheetManager
import inspect
print(inspect.getdoc(WorksheetManager))
```
