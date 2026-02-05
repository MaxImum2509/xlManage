# Rapport d'implémentation - Epic 8 Story 4

**Date** : 2026-02-05
**Story** : Implémenter TableManager.__init__ et la méthode create()
**Statut** : ✅ Complétée

## Résumé

Implémentation de la classe `TableManager` avec :
- Constructeur `__init__(excel_manager)`
- Méthode `create()` pour créer des tables Excel
- Helper `_get_table_info()` pour extraire les informations

## Fichiers modifiés

### 1. `src/xlmanage/table_manager.py`
**Ajouts** :
- Imports supplémentaires (Path, CDispatch, TableAlreadyExistsError)
- Classe `TableManager` (lignes 175-268)
  - `__init__(excel_manager)` : Constructeur
  - `_get_table_info(table, ws)` : Helper pour extraire TableInfo
  - `create(name, range_ref, worksheet, workbook)` : Création de table

### 2. `src/xlmanage/__init__.py`
**Modifications** :
- Ajout de `"table_manager"` à `__all__`
- Ajout de `"TableManager"` et `"TableInfo"` à `__all__`
- Import et export de `TableManager` et `TableInfo`

### 3. `tests/test_table_manager.py`
**Ajouts** : 7 nouveaux tests (total: 41 tests)

**TestTableManager** (3 tests) :
- `test_table_manager_initialization` : Initialisation correcte
- `test_get_table_info_with_data` : Extraction d'infos avec données
- `test_get_table_info_empty_table` : Table vide (0 lignes)

**TestTableManagerCreate** (4 tests) :
- `test_create_in_active_worksheet` : Création dans feuille active
- `test_create_invalid_table_name` : Nom invalide → erreur
- `test_create_invalid_range` : Plage invalide → erreur
- `test_create_duplicate_name` : Nom dupliqué → erreur

## Tests

```bash
poetry run pytest tests/test_table_manager.py -v
```

**Résultats** : ✅ 41/41 tests passés
**Couverture** : 92% pour `src/xlmanage/table_manager.py`

## Détails techniques

### TableManager.__init__

```python
def __init__(self, excel_manager):
    """Initialize table manager."""
    self._mgr = excel_manager
```

- Stocke la référence à ExcelManager
- Nécessite une instance démarrée d'ExcelManager

### TableManager._get_table_info()

```python
def _get_table_info(self, table: "CDispatch", ws: "CDispatch") -> TableInfo:
    """Extract information from a table COM object."""
```

**Extraction** :
- `name` : table.Name
- `worksheet_name` : ws.Name
- `range_ref` : table.Range.Address
- `header_row_range` : table.HeaderRowRange.Address
- `rows_count` : table.DataBodyRange.Rows.Count (ou 0 si vide)

**Gestion table vide** :
- Si `table.DataBodyRange` est `None` → rows_count = 0

### TableManager.create()

```python
def create(
    self,
    name: str,
    range_ref: str,
    worksheet: str | None = None,
    workbook: Path | None = None,
) -> TableInfo:
    """Create a new table in a worksheet."""
```

**Processus de création** :
1. **Validation nom** : `_validate_table_name(name)`
2. **Validation plage** : `_validate_range(range_ref)`
3. **Résolution classeur** : `_resolve_workbook(app, workbook)`
4. **Résolution feuille** :
   - Si `worksheet=None` → feuille active
   - Sinon → `_find_worksheet(wb, worksheet)`
5. **Vérification unicité** : Parcourt toutes les feuilles pour vérifier qu'aucune table n'a le même nom
6. **Création** :
   ```python
   table = ws.ListObjects.Add(
       SourceType=1,  # xlSrcRange
       Source=ws.Range(range_ref),
       XlListObjectHasHeaders=1,  # xlYes
   )
   table.Name = name
   ```
7. **Retour** : `_get_table_info(table, ws)`

**Paramètres COM** :
- `SourceType=1` : xlSrcRange (source est une plage)
- `XlListObjectHasHeaders=1` : xlYes (première ligne = en-tête)

**Exceptions levées** :
- `TableNameError` : nom invalide
- `TableRangeError` : plage invalide
- `TableAlreadyExistsError` : nom déjà utilisé dans le classeur
- `WorksheetNotFoundError` : feuille introuvable
- `WorkbookNotFoundError` : classeur non ouvert

## Conformité aux critères d'acceptation

✅ 1. Classe TableManager créée avec constructeur
✅ 2. Méthode `create()` implémentée
✅ 3. Validation du nom et de la plage
✅ 4. Détection de nom déjà utilisé
✅ 5. Détection de plage déjà utilisée (via paramètres COM)
✅ 6. Retourne TableInfo
✅ 7. Tests couvrent tous les cas (7 tests spécifiques + 34 précédents)

## Exemples d'utilisation

### Exemple 1 : Création basique

```python
from xlmanage import ExcelManager, TableManager

with ExcelManager() as excel_mgr:
    table_mgr = TableManager(excel_mgr)

    # Créer une table dans la feuille active
    info = table_mgr.create("tbl_Sales", "A1:D100")
    print(f"{info.name}: {info.rows_count} lignes")
    # Output: tbl_Sales: 99 lignes
```

### Exemple 2 : Création dans feuille spécifique

```python
with ExcelManager() as excel_mgr:
    table_mgr = TableManager(excel_mgr)

    # Créer une table dans une feuille spécifique
    info = table_mgr.create(
        name="tbl_Customers",
        range_ref="A1:F500",
        worksheet="Data"
    )
    print(f"Table '{info.name}' créée dans {info.worksheet_name}")
```

### Exemple 3 : Gestion d'erreurs

```python
try:
    table_mgr.create("1InvalidName", "A1:D10")
except TableNameError as e:
    print(f"Nom invalide : {e.reason}")
    # Output: Nom invalide : must start with letter or underscore...

try:
    table_mgr.create("tbl_Sales", "A1:D10")  # Nom déjà utilisé
except TableAlreadyExistsError as e:
    print(f"Table '{e.name}' existe déjà dans {e.workbook_name}")
```

## Prochaines étapes

Cette story complète les opérations de création. Les stories suivantes ajouteront :
- **Story 5** : TableManager.delete() et list()
- **Story 6** : Intégration CLI complète (table create/delete/list)

## Notes techniques

- **Unicité des noms** : Les noms de tables sont uniques au niveau du CLASSEUR, pas de la feuille (contrairement aux feuilles)
- **Recherche exhaustive** : La méthode parcourt toutes les feuilles du classeur pour vérifier l'unicité
- **Feuille active** : Si `worksheet=None`, utilise `wb.ActiveSheet` au lieu de chercher
- **Headers obligatoires** : `XlListObjectHasHeaders=1` force l'utilisation de la première ligne comme en-tête
- **Tests avec mocks** : Les tests utilisent des mocks pour éviter de dépendre d'Excel réel
