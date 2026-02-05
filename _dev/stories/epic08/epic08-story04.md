# Epic 8 - Story 4: Implémenter TableManager.__init__ et la méthode create()

**Statut** : ✅ Implémentée

**En tant que** utilisateur
**Je veux** créer une nouvelle table dans une feuille
**Afin de** structurer mes données dans Excel

## Critères d'acceptation

1. ✅ Classe TableManager créée avec constructeur
2. ✅ Méthode `create()` implémentée
3. ✅ Validation du nom et de la plage
4. ✅ Détection de nom déjà utilisé
5. ✅ Détection de plage déjà utilisée
6. ✅ Retourne TableInfo
7. ✅ Tests couvrent tous les cas

## Tâches techniques

### Tâche 4.1 : Créer la classe TableManager

```python
class TableManager:
    """Manager for Excel table (ListObject) CRUD operations.

    This class provides methods to create, delete, and list tables.
    It depends on ExcelManager for COM access.

    Note:
        The ExcelManager instance must be started before using this manager.
    """

    def __init__(self, excel_manager: ExcelManager):
        """Initialize table manager.

        Args:
            excel_manager: An ExcelManager instance (must be started)

        Example:
            >>> with ExcelManager() as excel_mgr:
            ...     table_mgr = TableManager(excel_mgr)
            ...     info = table_mgr.create("tbl_Sales", "A1:D100", worksheet="Data")
        """
        self._mgr = excel_manager
```

### Tâche 4.2 : Implémenter la méthode create()

La méthode `create()` doit :
1. Valider le nom de table
2. Valider la plage
3. Résoudre la feuille cible
4. Vérifier l'unicité du nom (au niveau du classeur)
5. Vérifier la plage n'existe pas déjà
6. Créer la table
7. Retourner TableInfo

**Signature** :
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

**Points d'attention** :
1. **Unicité du nom** : au niveau du CLASSEUR, pas de la feuille
2. **Validation plage** : syntaxe Excel valide, pas de chevauchement
3. **Résolution feuille** : trouver la bonne feuille dans le classeur
4. **Retour** : TableInfo avec tous les détails de la table créée

### Tâche 4.3 : Implémenter _get_table_info() helper

```python
def _get_table_info(self, table: CDispatch, ws: CDispatch) -> TableInfo:
    """Extract information from a table COM object."""
    return TableInfo(
        name=table.Name,
        worksheet_name=ws.Name,
        range_ref=table.Range.Address,
        header_row_range=table.HeaderRowRange.Address,
        rows_count=table.DataBodyRange.Rows.Count if table.DataBodyRange else 0,
    )
```

## Dépendances

- Story 1-3 (Toutes les stories précédentes)

## Définition of Done

- [x] TableManager.__init__ implémenté
- [x] Méthode create() implémentée avec toutes les validations
- [x] Helper _get_table_info() implémenté
- [x] TableManager et TableInfo exportés dans __init__.py
- [x] Tous les tests passent (41 tests au total, 7 pour TableManager)
- [x] Couverture de code 92%
- [x] Documentation complète avec exemples
