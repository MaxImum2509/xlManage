# TableManager - CRUD ListObjects (Tables Excel)

## Import

```python
from xlmanage import TableManager, TableInfo
```

## TableManager API

### Instanciation

```python
mgr = ExcelManager()
mgr.start()

tbl_mgr = TableManager(mgr.app)
```

### Méthodes CRUD

| Méthode | Paramètres | Retour | Description |
|----------|------------|--------|-------------|
| `create(name, range_ref, workbook=None, worksheet=None)` | `name, range_ref: str, workbook: Path|None, worksheet: str|None` | `TableInfo` | Crée nouvelle table |
| `delete(name, workbook=None, worksheet=None)` | `name: str, workbook: Path|None, worksheet: str|None` | `None` | Supprime table |
| `list(workbook=None, worksheet=None)` | `workbook: Path|None, worksheet: str|None` | `List[TableInfo]` | Liste toutes les tables |

## Usage Patterns

### Créer une Nouvelle Table

```python
from pathlib import Path

with ExcelManager(visible=False) as mgr:
    mgr.start()
    tbl_mgr = TableManager(mgr.app)

    # Créer dans feuille active
    info = tbl_mgr.create(
        name="tblSales",
        range_ref="A1:D100"
    )
    print(f"Created: {info.name} with {info.rows_count} rows")

    # Créer dans feuille spécifique
    info = tbl_mgr.create(
        name="tblCustomers",
        range_ref="A1:E50",
        workbook=Path("data.xlsx"),
        worksheet="Customers"
    )
```

### Supprimer une Table

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    tbl_mgr = TableManager(mgr.app)

    # Supprimer dans feuille active
    tbl_mgr.delete("OldTable")

    # Supprimer dans feuille spécifique
    tbl_mgr.delete("TempTable",
                   workbook=Path("data.xlsx"),
                   worksheet="Data")
```

### Lister les Tables

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    tbl_mgr = TableManager(mgr.app)

    # Lister toutes les tables (tous classeurs/feuilles)
    tables = tbl_mgr.list()

    for tbl in tables:
        print(f"{tbl.name}: {tbl.rows_count} rows, "
              f"{len(tbl.columns)} cols, "
              f"sheet={tbl.worksheet_name}, "
              f"range={tbl.range_address}")

    # Lister tables d'une feuille spécifique
    tables = tbl_mgr.list(workbook=Path("data.xlsx"),
                          worksheet="Data")
```

## TableInfo Structure

```python
dataclass TableInfo:
    name: str              # Nom de la table (ex: "tblSales")
    worksheet_name: str     # Nom de la feuille conteneur
    range_address: str      # Plage Excel (ex: "$A$1:$D$100")
    columns: List[str]     # Liste des noms de colonnes
    rows_count: int        # Nombre de lignes de données (excl. header)
    header_row: str        # Adresse de la ligne header (ex: "$A$1:$D$1")
```

## Exceptions

| Exception | Condition | Contexte |
|-----------|-----------|-----------|
| `TableAlreadyExistsError` | Nom déjà utilisé | `create()` |
| `TableNotFoundError` | Table introuvable | `delete()` |
| `TableNameError` | Nom invalide | `create()` |
| `TableRangeError` | Plage invalide | `create()` |

## Workflow Complet

```python
from pathlib import Path
from xlmanage import (
    ExcelManager, TableManager,
    TableNotFoundError, TableAlreadyExistsError, TableRangeError
)

def setup_sales_tables(data_path: Path):
    """Configure les tables du classeur sales"""
    with ExcelManager(visible=False) as mgr:
        mgr.start()
        tbl_mgr = TableManager(mgr.app)

        try:
            # Créer table principale
            sales_tbl = tbl_mgr.create(
                name="tblSales",
                range_ref="A1:F1000",
                workbook=data_path,
                worksheet="Sales"
            )
            print(f"Created {sales_tbl.name}: {sales_tbl.rows_count} rows, "
                  f"columns: {', '.join(sales_tbl.columns)}")

            # Créer table de lookup
            products_tbl = tbl_mgr.create(
                name="tblProducts",
                range_ref="A1:E50",
                workbook=data_path,
                worksheet="Products"
            )
            print(f"Created {products_tbl.name}")

            # Lister toutes les tables
            all_tables = tbl_mgr.list()
            print(f"\nAll tables in workbook:")
            for tbl in all_tables:
                print(f"  - {tbl.name} ({tbl.rows_count} rows)")

        except TableAlreadyExistsError as e:
            print(f"Table '{e.name}' already exists in {e.workbook_name}")
        except TableRangeError as e:
            print(f"Invalid range '{e.range_ref}': {e.reason}")
        except TableNameError as e:
            print(f"Invalid table name: {e.reason}")

# Usage
setup_sales_tables(Path("data/sales.xlsx"))
```

## COM Access Direct

Après création via TableManager, manipuler via COM :

```python
info = tbl_mgr.create("tblData", "A1:D100")
app = mgr.app
wb = app.Workbooks(info.worksheet_name + ".xlsx")  # Adapter
ws = wb.Worksheets(info.worksheet_name)
tbl = ws.ListObjects(info.name)

# Maintenant manipulation directe COM possible
tbl.ListRows.Add()
tbl.ListColumns("Total").DataBodyRange.Formula = "=SUM([@[Qty]]*[@[Price]])"
```

## Trucs et Astuces

### Déterminer la Plage Automatiquement

```python
def detect_range(ws, start_cell="A1") -> str:
    """Détecte la plage utilisée depuis une cellule"""
    # Utiliser UsedRange ou dernière ligne/colonne
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp
    last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column  # xlToLeft
    return f"{start_cell}:{ws.Cells(last_row, last_col).Address(False, False)}"
```

### Récupérer les Données d'une Table

```python
def get_table_data(tbl_mgr, name, workbook=None, worksheet=None):
    """Récupère données table sous forme de liste de dictionnaires"""
    info = next((t for t in tbl_mgr.list(workbook, worksheet)
                 if t.name == name), None)
    if not info:
        return None

    app = mgr.app
    ws = app.Workbooks(workbook or app.ActiveWorkbook.Name).Worksheets(info.worksheet_name)
    tbl = ws.ListObjects(info.name)

    # Lire via COM
    data_range = tbl.DataBodyRange.Value
    columns = info.columns

    # Convertir en liste de dict
    result = []
    for i in range(1, len(data_range) + 1):
        row_data = {}
        for j, col in enumerate(columns):
            row_data[col] = data_range[i, j + 1]
        result.append(row_data)

    return result
```

### Mettre à Jour une Table

```python
def update_table_data(tbl_mgr, name, data, workbook=None, worksheet=None):
    """Met à jour table avec données (remplace tout)"""
    info = next((t for t in tbl_mgr.list(workbook, worksheet)
                 if t.name == name), None)
    if not info:
        raise TableNotFoundError(name, worksheet or "Active")

    app = mgr.app
    wb = app.Workbooks(workbook or app.ActiveWorkbook.Name)
    ws = wb.Worksheets(info.worksheet_name)
    tbl = ws.ListObjects(info.name)

    # Préparer tableau 2D
    data_2d = [[row[col] for col in info.columns] for row in data]

    # Écrire via COM (batch)
    if tbl.DataBodyRange is not None:
        tbl.DataBodyRange.Delete()  # Effacer existant

    tbl.Resize(ws.Range(info.range_address).Resize(
        len(data) + 1, len(info.columns)
    ))

    tbl.DataBodyRange.Value = data_2d
```

## Documentation API

Pour voir l'API complète, utiliser les docstrings Python :

```python
from xlmanage import TableManager
import inspect
print(inspect.getdoc(TableManager))
```
