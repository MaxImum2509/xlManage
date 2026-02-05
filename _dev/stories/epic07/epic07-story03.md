# Epic 7 - Story 3: Implémenter les fonctions utilitaires _resolve_workbook et _find_worksheet

**Statut** : ✅ Terminé

**En tant que** développeur
**Je veux** des fonctions pour résoudre le classeur cible et chercher une feuille
**Afin de** faciliter la manipulation des feuilles dans différents classeurs

## Critères d'acceptation

1. ✅ Fonction `_resolve_workbook()` implémentée ✓
2. ✅ Support classeur explicite ou classeur actif ✓
3. ✅ Fonction `_find_worksheet()` implémentée ✓
4. ✅ Recherche case-insensitive ✓
5. ✅ Tests couvrent tous les scénarios ✓

## Tâches techniques

### Tâche 3.1 : Implémenter _resolve_workbook

**Fichier** : `src/xlmanage/worksheet_manager.py`

```python
def _resolve_workbook(app: CDispatch, workbook: Path | None) -> CDispatch:
    """Resolve the target workbook.

    If workbook is provided, finds or opens that specific workbook.
    If workbook is None, returns the active workbook.

    Args:
        app: Excel Application COM object
        workbook: Optional path to a specific workbook.
                  If None, uses the active workbook.

    Returns:
        Workbook COM object

    Raises:
        WorkbookNotFoundError: If the specified workbook is not open
        ExcelConnectionError: If no active workbook when workbook=None

    Examples:
        >>> # Use active workbook
        >>> wb = _resolve_workbook(app, None)

        >>> # Use specific workbook
        >>> wb = _resolve_workbook(app, Path("C:/data/test.xlsx"))

    Note:
        This function does NOT open the workbook if it's not already open.
        Use WorkbookManager.open() to open a workbook first.
        This function is shared by WorksheetManager, TableManager, and VBAManager.
    """
    if workbook is None:
        # Use active workbook
        try:
            wb = app.ActiveWorkbook
            if wb is None:
                from .exceptions import ExcelConnectionError
                raise ExcelConnectionError(
                    0x80080005,
                    "No active workbook. Open a workbook first."
                )
            return wb
        except Exception as e:
            from .exceptions import ExcelConnectionError
            if hasattr(e, "hresult"):
                raise ExcelConnectionError(
                    getattr(e, "hresult"),
                    f"Failed to get active workbook: {str(e)}"
                ) from e
            else:
                raise
    else:
        # Find the specified workbook
        from .workbook_manager import _find_open_workbook
        from .exceptions import WorkbookNotFoundError

        wb = _find_open_workbook(app, workbook)
        if wb is None:
            raise WorkbookNotFoundError(
                workbook,
                f"Workbook is not open: {workbook.name}"
            )
        return wb
```

### Tâche 3.2 : Implémenter _find_worksheet

```python
def _find_worksheet(wb: CDispatch, name: str) -> CDispatch | None:
    """Find a worksheet by name in a workbook.

    Searches for a worksheet with the given name.
    The search is case-insensitive (Excel behavior).

    Args:
        wb: Workbook COM object to search in
        name: Name of the worksheet to find

    Returns:
        Worksheet COM object if found, None otherwise

    Examples:
        >>> ws = _find_worksheet(wb, "Sheet1")
        >>> if ws:
        ...     print(f"Found: {ws.Name}")

        >>> # Case-insensitive
        >>> ws = _find_worksheet(wb, "SHEET1")  # Finds "Sheet1"

    Note:
        Excel worksheet names are case-insensitive but case-preserving.
        "Sheet1" and "SHEET1" refer to the same worksheet.
    """
    # Normalize search name to lowercase
    search_name = name.lower()

    # Iterate through all worksheets
    for ws in wb.Worksheets:
        try:
            # Compare case-insensitive
            if ws.Name.lower() == search_name:
                return ws
        except Exception:
            # Skip worksheets that can't be read
            continue

    return None
```

## Dépendances

- Story 2 (WorksheetInfo) - ✅ À créer avant
- Story 1 (Exceptions) - ✅ À créer avant

## Définition of Done

- [x] Fonction `_resolve_workbook()` implémentée
- [x] Fonction `_find_worksheet()` implémentée
- [x] Support classeur actif et classeur spécifique
- [x] Recherche case-insensitive
- [x] Tous les tests passent (18 tests au lieu de 12)
- [x] Couverture de code 93% (lignes manquantes: imports alternatifs)
- [x] Documentation complète avec exemples

## Rapport d'implémentation

**Date** : 2026-02-05
**Développeur** : Claude Sonnet 4.5

### Résumé

Implémentation complète et réussie des fonctions utilitaires `_resolve_workbook()` et `_find_worksheet()` dans le module `worksheet_manager.py`. Ces fonctions sont essentielles pour la manipulation des feuilles Excel dans différents classeurs.

### Implémentation

#### Fichiers modifiés

1. **src/xlmanage/worksheet_manager.py**
   - Ajout des imports nécessaires (Path, CDispatch)
   - Implémentation de `_resolve_workbook()` (lignes 95-158)
   - Implémentation de `_find_worksheet()` (lignes 161-199)

2. **tests/test_worksheet_manager.py**
   - Ajout de la classe `TestResolveWorkbook` avec 7 tests
   - Ajout de la classe `TestFindWorksheet` avec 11 tests
   - Ajout des imports nécessaires (patch, Mock)

#### Fonction `_resolve_workbook()`

**Emplacement** : src/xlmanage/worksheet_manager.py:95-158

**Fonctionnalités implémentées :**
- Résolution du classeur actif quand `workbook=None`
- Recherche d'un classeur spécifique par chemin
- Gestion des erreurs COM avec hresult
- Gestion des exceptions génériques
- Validation de la présence d'un classeur actif
- Utilisation de `_find_open_workbook()` du module workbook_manager

**Gestion d'erreurs :**
- `ExcelConnectionError` : Quand aucun classeur actif n'est disponible
- `ExcelConnectionError` : Quand une erreur COM se produit (avec hresult)
- `WorkbookNotFoundError` : Quand le classeur spécifié n'est pas ouvert
- Re-lève les exceptions non-COM telles quelles

#### Fonction `_find_worksheet()`

**Emplacement** : src/xlmanage/worksheet_manager.py:161-199

**Fonctionnalités implémentées :**
- Recherche case-insensitive (comportement Excel natif)
- Normalisation en minuscules pour la comparaison
- Itération sécurisée sur toutes les feuilles
- Gestion robuste des erreurs de lecture
- Support des noms Unicode
- Support des caractères spéciaux

**Comportement :**
- Retourne le premier worksheet correspondant
- Retourne `None` si aucune feuille ne correspond
- Ignore les feuilles qui provoquent des erreurs de lecture

### Tests

#### Tests pour `_resolve_workbook()` (7 tests)

1. `test_resolve_workbook_with_none_returns_active` : Résolution du classeur actif
2. `test_resolve_workbook_with_none_no_active_raises` : Erreur si pas de classeur actif
3. `test_resolve_workbook_with_none_com_error_raises` : Gestion erreur COM
4. `test_resolve_workbook_with_path_finds_open` : Recherche classeur par chemin
5. `test_resolve_workbook_with_path_not_open_raises` : Erreur si classeur fermé
6. `test_resolve_workbook_preserves_workbook_object` : Préservation de l'objet
7. `test_resolve_workbook_with_none_non_com_error_raises` : Erreur non-COM

#### Tests pour `_find_worksheet()` (11 tests)

1. `test_find_worksheet_exact_match` : Correspondance exacte
2. `test_find_worksheet_case_insensitive` : Recherche insensible à la casse
3. `test_find_worksheet_not_found` : Feuille inexistante
4. `test_find_worksheet_empty_workbook` : Classeur vide
5. `test_find_worksheet_multiple_sheets` : Plusieurs feuilles
6. `test_find_worksheet_handles_read_error` : Gestion erreur de lecture
7. `test_find_worksheet_all_error_returns_none` : Toutes les feuilles en erreur
8. `test_find_worksheet_unicode_names` : Noms Unicode
9. `test_find_worksheet_special_characters` : Caractères spéciaux
10. `test_find_worksheet_returns_first_match` : Première correspondance
11. `test_find_worksheet_preserves_worksheet_object` : Préservation de l'objet

### Résultats des tests

```
============================= test session starts =============================
Platform: Windows (Python 3.14.2)
Tests collected: 229
Tests passed: 228
Tests xfailed: 1

Coverage Results:
- Global: 90.64% (objectif 90% dépassé)
- worksheet_manager.py: 93%
- Lignes non couvertes: 27-28, 32-33 (imports alternatifs)

Test Duration: 28.66s
Status: ✅ ALL TESTS PASSED
```

### Couverture de code

**Couverture globale du fichier** : 93%

**Lignes non couvertes** :
- Lignes 27-28 : Import alternatif pour CDispatch
- Lignes 32-33 : Import alternatif pour exceptions

Ces lignes sont des imports fallback qui ne sont pas exécutés dans l'environnement de test actuel où les imports principaux réussissent. Elles assurent la compatibilité avec différentes configurations d'import.

**Branches testées** :
- ✅ Classeur actif vs classeur spécifique
- ✅ Erreurs COM avec et sans hresult
- ✅ Classeur trouvé vs non trouvé
- ✅ Feuille trouvée vs non trouvée
- ✅ Recherche case-insensitive
- ✅ Gestion des erreurs de lecture
- ✅ Collections vides vs peuplées

### Qualité du code

**Points forts :**
- Documentation complète avec docstrings détaillés
- Exemples d'utilisation dans les docstrings
- Gestion robuste des erreurs
- Tests exhaustifs couvrant les cas nominaux et d'erreur
- Code conforme aux standards du projet (Ruff, MyPy)
- Utilisation appropriée des types hints

**Architecture :**
- Séparation claire des responsabilités
- Réutilisation de `_find_open_workbook()` existant
- Imports dynamiques pour éviter les dépendances circulaires
- Fonction `_find_worksheet()` stateless et testable

### Dépendances

**Story 2 (WorksheetInfo)** : ✅ Utilisé
**Story 1 (Exceptions)** : ✅ Utilisé

Les fonctions implémentées utilisent :
- `WorkbookNotFoundError` de Story 1
- `ExcelConnectionError` de Story 1
- `_find_open_workbook()` du module workbook_manager

### Compatibilité

**Compatibilité ascendante** : ✅ Maintenue
- Aucune modification des fonctions existantes
- Ajout uniquement de nouvelles fonctions
- Tous les tests existants continuent de passer

**Intégration** : ✅ Prête
- Les fonctions sont prêtes à être utilisées par :
  - WorksheetManager (Epic 7 Story 4+)
  - TableManager (Epic futur)
  - VBAManager (Epic futur)

### Conclusions

L'implémentation est complète et de haute qualité :
- ✅ Tous les critères d'acceptation satisfaits
- ✅ Couverture de code excellente (93%)
- ✅ Tests exhaustifs (18 tests pour 2 fonctions)
- ✅ Documentation complète
- ✅ Code robuste et maintenable
- ✅ Prêt pour la production

**Prochaine étape recommandée** : Epic 7 - Story 4 (Implémentation WorksheetManager)
