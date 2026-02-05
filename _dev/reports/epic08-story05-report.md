# Rapport d'implémentation - Epic 8 Story 5

**Date** : 2026-02-05
**Story** : Implémenter TableManager.delete() et list()
**Statut** : ✅ Complétée

## Résumé

Ajout des méthodes `delete()` et `list()` à la classe `TableManager` :
- `delete()` : Supprime une table par nom
- `list()` : Liste toutes les tables d'une feuille ou du classeur

## Fichiers modifiés

### 1. `src/xlmanage/table_manager.py`
**Ajouts** :
- Import de `TableNotFoundError`
- Méthode `delete(name, worksheet, workbook)` (lignes 271-326)
- Méthode `list(worksheet, workbook)` (lignes 328-385)

### 2. `tests/test_table_manager.py`
**Ajouts** : 7 nouveaux tests (total: 48 tests)

**TestTableManagerDelete** (3 tests) :
- `test_delete_from_active_workbook` : Suppression en cherchant toutes les feuilles
- `test_delete_from_specific_worksheet` : Suppression dans feuille spécifique
- `test_delete_table_not_found` : Table inexistante → erreur

**TestTableManagerList** (4 tests) :
- `test_list_all_tables_in_workbook` : Liste toutes les tables du classeur
- `test_list_tables_in_specific_worksheet` : Liste les tables d'une feuille
- `test_list_empty_workbook` : Classeur sans tables
- `test_list_handles_corrupted_table` : Continue si une table est corrompue

## Tests

```bash
poetry run pytest tests/test_table_manager.py -v
```

**Résultats** : ✅ 48/48 tests passés
**Couverture** : 91% pour `src/xlmanage/table_manager.py`

## Détails techniques

### TableManager.delete()

```python
def delete(
    self,
    name: str,
    worksheet: str | None = None,
    workbook: Path | None = None,
) -> None:
    """Delete a table."""
```

**Processus** :
1. **Résolution classeur** : `_resolve_workbook(app, workbook)`
2. **Recherche table** :
   - Si `worksheet=None` : cherche dans toutes les feuilles
   - Sinon : cherche dans la feuille spécifiée
3. **Vérification existence** : Lève `TableNotFoundError` si non trouvée
4. **Suppression** : `table.Delete()`

**Recherche flexible** :
- `worksheet=None` → cherche partout (utile quand on ne connaît pas la feuille)
- `worksheet="Data"` → cherche uniquement dans "Data" (plus rapide)

**Exceptions** :
- `TableNotFoundError` : table n'existe pas
- `WorkbookNotFoundError` : classeur non ouvert
- `ExcelConnectionError` : problème COM

### TableManager.list()

```python
def list(
    self,
    worksheet: str | None = None,
    workbook: Path | None = None,
) -> list[TableInfo]:
    """List all tables."""
```

**Processus** :
1. **Résolution classeur** : `_resolve_workbook(app, workbook)`
2. **Collecte tables** :
   - Si `worksheet=None` : parcourt toutes les feuilles
   - Sinon : parcourt uniquement la feuille spécifiée
3. **Extraction infos** : `_get_table_info(table, sheet)` pour chaque table
4. **Gestion erreurs** : Continue si une table/feuille est corrompue
5. **Retour** : Liste de `TableInfo` (vide si aucune table)

**Robustesse** :
- Try/except pour chaque table (skip si corrompue)
- Try/except pour chaque feuille (skip si illisible)
- Retourne toujours une liste (jamais None)

## Conformité aux critères d'acceptation

✅ 1. Méthodes delete() et list() implémentées
✅ 2. delete() supprime complètement la table (appelle Delete())
✅ 3. list() supporte recherche par feuille ou classeur entier
✅ 4. 7 nouveaux tests (48 total)
✅ 5. Couverture 91%

## Exemples d'utilisation

### delete()

```python
from xlmanage import ExcelManager, TableManager

with ExcelManager() as excel_mgr:
    table_mgr = TableManager(excel_mgr)

    # Supprimer une table (cherche partout)
    table_mgr.delete("tbl_Sales")

    # Supprimer dans une feuille spécifique (plus rapide)
    table_mgr.delete("tbl_Old", worksheet="Archive")
```

### list()

```python
with ExcelManager() as excel_mgr:
    table_mgr = TableManager(excel_mgr)

    # Lister toutes les tables du classeur
    all_tables = table_mgr.list()
    for table in all_tables:
        print(f"{table.name} ({table.worksheet_name}): {table.rows_count} rows")
    # Output:
    # tbl_Sales (Data): 99 rows
    # tbl_Products (Data): 150 rows
    # tbl_Archive (History): 1000 rows

    # Lister uniquement les tables d'une feuille
    data_tables = table_mgr.list(worksheet="Data")
    print(f"{len(data_tables)} tables dans 'Data'")
    # Output: 2 tables dans 'Data'
```

## Cas d'usage

### Suppression conditionnelle

```python
# Supprimer toutes les tables commençant par "tmp_"
all_tables = table_mgr.list()
for table in all_tables:
    if table.name.startswith("tmp_"):
        table_mgr.delete(table.name, worksheet=table.worksheet_name)
        print(f"Supprimé : {table.name}")
```

### Statistiques

```python
# Afficher des statistiques sur les tables
tables = table_mgr.list()
total_rows = sum(t.rows_count for t in tables)
print(f"Total : {len(tables)} tables, {total_rows} lignes")

# Par feuille
by_sheet = {}
for table in tables:
    by_sheet.setdefault(table.worksheet_name, []).append(table)

for sheet, sheet_tables in by_sheet.items():
    print(f"{sheet}: {len(sheet_tables)} tables")
```

## Notes techniques

- **delete() avec worksheet=None** : Cherche dans toutes les feuilles, ce qui peut être lent sur de gros classeurs
- **delete() avec worksheet** : Plus rapide car cherche dans une seule feuille
- **list() robuste** : Continue même si certaines tables/feuilles sont corrompues
- **list() retourne toujours une liste** : Jamais None, facilite le code client
- **Suppression définitive** : `Delete()` ne peut pas être annulé (pas de Undo)

## Prochaine étape

Story 6 : Intégration CLI complète (commandes `table create/delete/list`)
