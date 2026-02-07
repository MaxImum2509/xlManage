# Epic 13 - Story 1: Mise en conformite de `table_manager.py`

**Statut** : ⚠️ Implémenté - Tests en cours de mise à jour (2026-02-07)

**Priorite** : P1 - Critique

**En tant que** mainteneur du projet
**Je veux** que `table_manager.py` soit conforme a l'architecture documentee
**Afin de** garantir la coherence API et eviter les bugs de comportement

## Contexte

L'audit du 2026-02-06 a identifie 9 anomalies dans `table_manager.py` (8 critiques, 1 mineure). Ce module est le plus impacte du projet. Les anomalies touchent la dataclass `TableInfo`, les signatures de fonctions utilitaires, et la logique metier de `create()` et `delete()`.

**Reference architecture** : section 4.5 (`_dev/architecture.md`, lignes 823-971)

## Anomalies a corriger

| ID      | Severite | Description                                                    |
| ------- | -------- | -------------------------------------------------------------- |
| TBL-001 | Critique | Champ `columns: list[str]` manquant dans `TableInfo`           |
| TBL-002 | Critique | Noms de champs incorrects (`range_ref` / `header_row_range`)   |
| TBL-003 | Critique | Signature `_find_table()` incorrecte (feuille vs classeur)     |
| TBL-004 | Critique | Signature `_validate_range()` incorrecte (pas de param COM)    |
| TBL-005 | Critique | Pas de verification de chevauchement dans `create()`           |
| TBL-006 | Critique | Extraction des colonnes absente dans `_get_table_info()`       |
| TBL-007 | Critique | `delete()` ignore le parametre `force`                         |
| TBL-008 | Critique | Ordre des parametres incorrect dans `create()`                 |
| TBL-009 | Mineur   | Notation regex legerement differente                           |

## Taches techniques

### Tache 1.1 : Corriger la dataclass `TableInfo` (TBL-001, TBL-002)

**Fichier** : `src/xlmanage/table_manager.py`

**Avant** :
```python
@dataclass
class TableInfo:
    name: str
    worksheet_name: str
    range_ref: str
    header_row_range: str
    rows_count: int
```

**Apres** :
```python
@dataclass
class TableInfo:
    name: str
    worksheet_name: str
    range_address: str       # Renomme depuis range_ref
    columns: list[str]       # NOUVEAU - noms des colonnes d'en-tete
    rows_count: int
    header_row: str          # Renomme depuis header_row_range
```

**Impact** : Ce renommage impacte `_get_table_info()`, tous les tests, et les references dans `cli.py` (champs `info.range_ref` et `info.header_row_range`).

### Tache 1.2 : Reecrire `_find_table()` pour chercher dans tout le classeur (TBL-003)

**Fichier** : `src/xlmanage/table_manager.py`

**Avant** :
```python
def _find_table(ws: CDispatch, name: str) -> CDispatch | None:
    # Cherche dans une seule feuille
```

**Apres** :
```python
def _find_table(wb: CDispatch, name: str) -> tuple[CDispatch, CDispatch] | None:
    """Recherche une table dans tout le classeur (noms uniques au classeur).

    Itere toutes les feuilles, puis tous les ListObjects de chaque feuille.

    Args:
        wb: Workbook COM object
        name: Name of the table to find

    Returns:
        (worksheet, listobject) si trouvee, None sinon.
    """
    for ws in wb.Worksheets:
        for table in ws.ListObjects:
            try:
                if table.Name == name:
                    return (ws, table)
            except Exception:
                continue
    return None
```

**Impact** : Tous les appelants de `_find_table()` doivent etre mis a jour (`delete()`, `create()` verification d'unicite).

### Tache 1.3 : Reecrire `_validate_range()` avec parametre COM (TBL-004, TBL-005)

**Fichier** : `src/xlmanage/table_manager.py`

**Avant** :
```python
def _validate_range(range_ref: str) -> None:
    # Validation purement textuelle (regex)
```

**Apres** :
```python
def _validate_range(ws: CDispatch, range_ref: str) -> CDispatch:
    """Valide et retourne un objet Range COM.

    Verifie :
    1. ws.Range(range_ref) ne raise pas (syntaxe valide)
    2. La plage ne chevauche pas un ListObject existant

    Args:
        ws: Worksheet COM object
        range_ref: Range reference (e.g., "A1:D10")

    Returns:
        Objet COM Range.

    Raises:
        TableRangeError: Si la plage est invalide ou chevauche une table.
    """
    if not range_ref or not range_ref.strip():
        raise TableRangeError(range_ref, "range cannot be empty")

    try:
        range_obj = ws.Range(range_ref)
    except Exception:
        raise TableRangeError(range_ref, "invalid range syntax")

    # Verifier le chevauchement avec les tables existantes
    for table in ws.ListObjects:
        try:
            existing_range = table.Range
            if _ranges_overlap(range_obj, existing_range):
                raise TableRangeError(
                    range_ref,
                    f"range overlaps with existing table '{table.Name}'"
                )
        except TableRangeError:
            raise
        except Exception:
            continue

    return range_obj
```

**Fonction helper** a ajouter :
```python
def _ranges_overlap(range1: CDispatch, range2: CDispatch) -> bool:
    """Verifie si deux plages COM se chevauchent via Intersect."""
    try:
        app = range1.Application
        intersection = app.Intersect(range1, range2)
        return intersection is not None
    except Exception:
        return False
```

### Tache 1.4 : Mettre a jour `_get_table_info()` pour extraire les colonnes (TBL-006)

**Fichier** : `src/xlmanage/table_manager.py`

**Avant** :
```python
def _get_table_info(self, table, ws) -> TableInfo:
    return TableInfo(
        name=table.Name,
        worksheet_name=ws.Name,
        range_ref=table.Range.Address,
        header_row_range=table.HeaderRowRange.Address,
        rows_count=table.DataBodyRange.Rows.Count if table.DataBodyRange else 0,
    )
```

**Apres** :
```python
def _get_table_info(self, table, ws) -> TableInfo:
    # Extraire les noms des colonnes
    columns = [col.Name for col in table.ListColumns]

    return TableInfo(
        name=table.Name,
        worksheet_name=ws.Name,
        range_address=table.Range.Address,
        columns=columns,
        rows_count=table.DataBodyRange.Rows.Count if table.DataBodyRange else 0,
        header_row=table.HeaderRowRange.Address,
    )
```

### Tache 1.5 : Implementer `Unlist()` vs `Delete()` selon `force` (TBL-007)

**Fichier** : `src/xlmanage/table_manager.py`

**Avant** :
```python
# Toujours table_found.Delete()
```

**Apres** :
```python
if force:
    table_found.Delete()    # Supprime table ET donnees
else:
    table_found.Unlist()    # Conserve les donnees, supprime la structure
```

### Tache 1.6 : Corriger l'ordre des parametres de `create()` (TBL-008)

**Fichier** : `src/xlmanage/table_manager.py`

**Avant** :
```python
def create(self, name, range_ref, worksheet, workbook) -> TableInfo:
```

**Apres** :
```python
def create(self, name, range_ref, workbook=None, worksheet=None) -> TableInfo:
```

Cet ordre est coherent avec les autres managers (`workbook` avant `worksheet`).

### Tache 1.7 : Harmoniser la notation regex (TBL-009)

**Fichier** : `src/xlmanage/table_manager.py:41`

**Avant** : `r"^[a-zA-Z_][a-zA-Z0-9_]*$"`
**Apres** : `r'^[A-Za-z_][A-Za-z0-9_]*$'`

Fonctionnellement equivalent mais consistant avec la spec.

### Tache 1.8 : Mettre a jour `create()` et `delete()` pour utiliser les nouvelles signatures

Adapter les corps de `create()` et `delete()` pour :
- Utiliser `_find_table(wb, name)` au lieu de `_find_table(ws, name)`
- Utiliser `_validate_range(ws, range_ref)` qui retourne un Range COM
- Passer le Range COM a `ListObjects.Add(Source=range_obj)`

### Tache 1.9 : Mettre a jour les references dans `cli.py`

**Fichier** : `src/xlmanage/cli.py`

Remplacer :
- `info.range_ref` -> `info.range_address`
- `info.header_row_range` -> `info.header_row`

### Tache 1.10 : Mettre a jour les tests

**Fichier** : `tests/test_table_manager.py`

- Adapter les tests aux nouveaux noms de champs (`range_address`, `columns`, `header_row`)
- Adapter les tests aux nouvelles signatures (`_find_table(wb, ...)`, `_validate_range(ws, ...)`)
- Ajouter tests pour :
  - Extraction des colonnes
  - Verification de chevauchement
  - `delete(force=False)` -> Unlist vs `delete(force=True)` -> Delete
  - Ordre des parametres de `create()`

## Criteres d'acceptation

1. [ ] `TableInfo` contient les 6 champs conformes a l'architecture (dont `columns: list[str]`)
2. [ ] `_find_table()` cherche dans tout le classeur et retourne `(ws, table)`
3. [ ] `_validate_range()` accepte un `ws` COM et retourne un `Range` COM
4. [ ] `create()` verifie le chevauchement avant creation
5. [ ] `delete()` utilise `Unlist()` par defaut et `Delete()` avec `force=True`
6. [ ] L'ordre des parametres de `create()` est `(name, range_ref, workbook, worksheet)`
7. [ ] Les colonnes sont extraites dans `_get_table_info()`
8. [ ] Tous les tests existants sont mis a jour et passent
9. [ ] `cli.py` utilise les nouveaux noms de champs

## Dependances

- Aucune dependance bloquante (module autonome)
- Impacte `cli.py` (champs renommes) et `__init__.py` (export `TableInfo`)

## Definition of Done

- [ ] Les 9 anomalies TBL-001 a TBL-009 sont corrigees
- [ ] Tous les tests `test_table_manager.py` passent
- [ ] Les tests CLI impliquant les tables passent
- [ ] Couverture > 90% pour `table_manager.py`
- [ ] `ruff check` passe sans erreur
