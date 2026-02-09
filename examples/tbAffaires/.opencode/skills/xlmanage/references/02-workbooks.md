# WorkbookManager - CRUD Classeurs

## Import

```python
from xlmanage import WorkbookManager, WorkbookInfo
```

## WorkbookManager API

### Instanciation

```python
# Nécessite un ExcelManager démarré
mgr = ExcelManager()
mgr.start()

wb_mgr = WorkbookManager(mgr.app)
```

### Méthodes CRUD

| Méthode | Paramètres | Retour | Description |
|----------|------------|--------|-------------|
| `create(path)` | `path: Path` | `WorkbookInfo` | Crée nouveau classeur |
| `open(path)` | `path: Path` | `WorkbookInfo` | Ouvre classeur existant |
| `close(name, save=True)` | `name: str, save: bool` | `None` | Ferme classeur |
| `save(name=None, path=None)` | `name: str|None, path: Path|None` | `None` | Sauvegarde classeur |
| `list()` | - | `List[WorkbookInfo]` | Liste tous les classeurs |

## Usage Patterns

### Créer un Nouveau Classeur

```python
from pathlib import Path

with ExcelManager(visible=False) as mgr:
    mgr.start()
    wb_mgr = WorkbookManager(mgr.app)

    info = wb_mgr.create(Path("output/new_workbook.xlsx"))
    print(f"Created: {info.name} with {info.sheets_count} sheets")
```

### Ouvrir un Classeur Existant

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    wb_mgr = WorkbookManager(mgr.app)

    info = wb_mgr.open(Path("data/existing.xlsm"))
    print(f"Opened: {info.full_path}, read_only={info.read_only}")
```

### Sauvegarder un Classeur

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    wb_mgr = WorkbookManager(mgr.app)

    info = wb_mgr.open(Path("data/workbook.xlsx"))

    # Sauvegarder dans le même fichier
    wb_mgr.save(name=info.name)

    # Sauvegarder dans un nouveau fichier (Save As)
    wb_mgr.save(name=info.name, path=Path("output/backup.xlsx"))
```

### Fermer un Classeur

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    wb_mgr = WorkbookManager(mgr.app)

    info = wb_mgr.open(Path("data/temp.xlsx"))

    # Fermer avec sauvegarde
    wb_mgr.close(name=info.name, save=True)

    # Fermer sans sauvegarde
    wb_mgr.close(name=info.name, save=False)
```

### Lister les Classeurs Ouverts

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    wb_mgr = WorkbookManager(mgr.app)

    workbooks = wb_mgr.list()

    for wb in workbooks:
        print(f"{wb.name}: {wb.sheets_count} sheets, "
              f"saved={wb.saved}, read_only={wb.read_only}")
```

## WorkbookInfo Structure

```python
dataclass WorkbookInfo:
    name: str           # Nom du classeur (ex: "Workbook1.xlsx")
    full_path: str       # Chemin complet du fichier
    read_only: bool      # Mode lecture seule
    saved: bool          # État de sauvegarde
    sheets_count: int     # Nombre de feuilles
```

## Exceptions

| Exception | Condition | Contexte |
|-----------|-----------|-----------|
| `WorkbookNotFoundError` | Fichier inexistant | `open()` |
| `WorkbookAlreadyOpenError` | Déjà ouvert | `open()` |
| `WorkbookSaveError` | Échec sauvegarde | `save()` |

## Workflow Complet

```python
from pathlib import Path
from xlmanage import ExcelManager, WorkbookManager, WorkbookNotFoundError

def process_workbook(input_path: Path, output_path: Path):
    """Ouvre, modifie, sauvegarde un classeur"""
    with ExcelManager(visible=False) as mgr:
        mgr.start()
        wb_mgr = WorkbookManager(mgr.app)

        try:
            # Ouvrir
            info = wb_mgr.open(input_path)
            print(f"Opened: {info.name}")

            # ... modifications ici via mgr.app.Workbooks(info.name) ...

            # Sauvegarder
            wb_mgr.save(name=info.name, path=output_path)
            print(f"Saved to: {output_path}")

            # Fermer
            wb_mgr.close(name=info.name, save=False)

        except WorkbookNotFoundError as e:
            print(f"File not found: {e.path}")
        except WorkbookSaveError as e:
            print(f"Save failed: {e.message}")

# Usage
process_workbook(
    Path("data/input.xlsx"),
    Path("output/result.xlsx")
)
```

## COM Access Direct

Après ouverture via WorkbookManager, accéder au COM object :

```python
info = wb_mgr.open(Path("data.xlsx"))
app = mgr.app
wb = app.Workbooks(info.name)

# Maintenant manipulation directe COM possible
ws = wb.Worksheets("Sheet1")
ws.Range("A1").Value = "Hello"
```

## Documentation API

Pour voir l'API complète, utiliser les docstrings Python :

```python
from xlmanage import WorkbookManager
import inspect
print(inspect.getdoc(WorkbookManager))
```
