# Epic 8 - Story 5: Implémenter TableManager.delete() et list()

**Statut** : ✅ À implémenter

**En tant que** utilisateur
**Je veux** supprimer une table et lister toutes les tables
**Afin de** gérer les tables de mes données

## Critères d'acceptation

1. ✅ Méthode `delete()` implémentée
2. ✅ Confirmation avant suppression
3. ✅ Méthode `list()` implémentée
4. ✅ Liste toutes les tables du classeur ou de la feuille
5. ✅ Tests couvrent tous les cas

## Tâches techniques

### Tâche 5.1 : Implémenter delete()

**Signature** :
```python
def delete(
    self,
    name: str,
    worksheet: str | None = None,
    workbook: Path | None = None,
) -> None:
    """Delete a table.

    Deletes the specified table from the worksheet.

    Args:
        name: Name of the table to delete
        worksheet: Worksheet containing the table (if None, search all worksheets)
        workbook: Target workbook path (if None, uses active workbook)

    Raises:
        TableNotFoundError: If the table doesn't exist
        WorkbookNotFoundError: If the specified workbook is not open
        ExcelConnectionError: If COM connection fails
    """
```

**Points d'attention** :
1. **Recherche** : chercher dans la feuille spécifiée ou toutes les feuilles
2. **Suppression** : table.Delete() supprime complètement la table
3. **Cleanup** : nettoyer la référence COM après suppression

### Tâche 5.2 : Implémenter list()

**Signature** :
```python
def list(
    self,
    worksheet: str | None = None,
    workbook: Path | None = None,
) -> list[TableInfo]:
    """List all tables.

    Returns information about all tables in the worksheet(s).

    Args:
        worksheet: Worksheet name to search (if None, list all tables in workbook)
        workbook: Target workbook path (if None, uses active workbook)

    Returns:
        List of TableInfo for each table
        Returns empty list if no tables found

    Raises:
        WorkbookNotFoundError: If the specified workbook is not open
        ExcelConnectionError: If COM connection fails

    Examples:
        >>> manager = TableManager(excel_mgr)
        >>> tables = manager.list(worksheet="Data")
        >>> for table in tables:
        ...     print(f"{table.name}: {table.rows_count} rows")
    """
```

**Points d'attention** :
1. **Portée** : si worksheet fourni, lister uniquement les tables de cette feuille
2. **Portée globale** : si worksheet=None, lister toutes les tables du classeur
3. **Gestion d'erreur** : continuer si une table est corrompue

## Dépendances

- Story 1-4 (Toutes les stories précédentes)

## Définition of Done

- [ ] Méthodes delete() et list() implémentées
- [ ] delete() supprime complètement la table
- [ ] list() supporte recherche par feuille ou classeur entier
- [ ] Tous les tests passent (minimum 8 tests)
- [ ] Couverture de code > 95%
