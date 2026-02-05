# Epic 8 - Story 3: Implémenter les fonctions utilitaires _find_table et _validate_range

**Statut** : ✅ Implémentée

**En tant que** développeur
**Je veux** des fonctions pour chercher une table et valider une plage Excel
**Afin de** faciliter la manipulation des tables dans les feuilles

## Critères d'acceptation

1. ✅ Fonction `_find_table()` implémentée
2. ✅ Recherche case-sensitive
3. ✅ Fonction `_validate_range()` implémentée
4. ✅ Validation syntaxe plage Excel
5. ✅ Tests couvrent tous les scénarios

## Tâches techniques

### Tâche 3.1 : Implémenter _find_table

**Fichier** : `src/xlmanage/table_manager.py`

```python
def _find_table(ws: CDispatch, name: str) -> CDispatch | None:
    """Find a table by name in a worksheet.

    Searches for a table (ListObject) with the given name.
    Note: Table names are case-SENSITIVE in Excel.

    Args:
        ws: Worksheet COM object to search in
        name: Name of the table to find

    Returns:
        Table COM object if found, None otherwise

    Examples:
        >>> table = _find_table(ws, "tbl_Sales")
        >>> if table:
        ...     print(f"Found: {table.Name}")

    Note:
        Unlike worksheet names, Excel table names are case-sensitive.
        "tbl_Sales" and "TBL_SALES" are different tables.
    """
    # Iterate through all tables in the worksheet
    for table in ws.ListObjects:
        try:
            if table.Name == name:  # Case-sensitive comparison
                return table
        except Exception:
            # Skip tables that can't be read
            continue

    return None
```

### Tâche 3.2 : Implémenter _validate_range

```python
def _validate_range(range_ref: str) -> None:
    """Validate an Excel range reference.

    Checks that the range has valid syntax and structure.

    Args:
        range_ref: Range reference to validate (e.g., "A1:D10")

    Raises:
        TableRangeError: If the range is invalid

    Examples:
        >>> _validate_range("A1:D10")  # OK
        >>> _validate_range("Sheet1!A1:D10")  # OK
        >>> _validate_range("A1:Z")  # Raises: invalid syntax
    """
    if not range_ref or not range_ref.strip():
        raise TableRangeError(range_ref, "range cannot be empty")

    # Remove sheet reference if present (e.g., "Sheet1!" or "'Sheet Name'!")
    clean_range = range_ref
    if "!" in clean_range:
        parts = clean_range.split("!", 1)
        if len(parts) == 2:
            clean_range = parts[1]

    # Must contain at least one colon (for start:end range)
    if ":" not in clean_range:
        raise TableRangeError(range_ref, "range must have format A1:Z99")

    # Basic pattern check for Excel ranges
    pattern = r'^[A-Z]+\d+:[A-Z]+\d+$|^[rR]\d+[cC]\d+:[rR]\d+[cC]\d+$'
    if not re.match(pattern, clean_range.replace("$", "")):
        raise TableRangeError(range_ref, "invalid range syntax")
```

## Dépendances

- Story 2 (TableInfo) - ✅ À créer avant
- Story 1 (Exceptions) - ✅ À créer avant

## Définition of Done

- [x] Fonction `_find_table()` implémentée
- [x] Fonction `_validate_range()` implémentée
- [x] Recherche case-sensitive pour les tables
- [x] Validation syntaxe plage Excel
- [x] Tous les tests passent (14 tests pour ces fonctions)
- [x] Couverture de code 95%
- [x] Documentation complète avec exemples
