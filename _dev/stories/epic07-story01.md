# Epic 7 - Story 1: Créer les exceptions pour la gestion des feuilles

**Statut** : ✅ Terminé

**En tant que** développeur
**Je veux** avoir des exceptions spécifiques pour les erreurs de gestion des feuilles
**Afin de** fournir des messages d'erreur clairs et exploitables aux utilisateurs

## Critères d'acceptation

1. ✅ Quatre nouvelles exceptions sont créées dans `src/xlmanage/exceptions.py`
2. ✅ Toutes héritent de `ExcelManageError`
3. ✅ Chaque exception a des attributs métier appropriés
4. ✅ Les exceptions sont exportées dans `__init__.py`
5. ✅ Les tests couvrent tous les cas d'usage

## Définition of Done

- [x] Les 4 exceptions sont créées avec docstrings complètes
- [x] Les exceptions sont exportées dans `__init__.py`
- [x] Tous les tests passent (15 tests ajoutés, 186 tests au total)
- [x] Couverture de code 100% pour les nouvelles exceptions
- [x] Le code suit les conventions du projet

## Dépendances

- Aucune dépendance (exceptions de base existent)

## Implémentation

### Fichiers modifiés

1. **src/xlmanage/exceptions.py** (lignes 145-196)
   - Ajout de 4 nouvelles classes d'exception pour la gestion des feuilles
   - `WorksheetNotFoundError` : Feuille introuvable
   - `WorksheetAlreadyExistsError` : Nom de feuille déjà utilisé
   - `WorksheetDeleteError` : Suppression de feuille impossible
   - `WorksheetNameError` : Nom de feuille invalide

2. **src/xlmanage/__init__.py**
   - Ajout des 4 nouvelles exceptions dans `__all__`
   - Ajout des 4 nouvelles exceptions dans les imports

3. **tests/test_exceptions.py** (lignes 20-33, 441-562)
   - Ajout de l'import des 4 nouvelles exceptions
   - Ajout de 4 classes de tests avec 15 tests au total

### Tests ajoutés

**TestWorksheetNotFoundError** (4 tests)
1. `test_worksheet_not_found_default_message` : Message par défaut
2. `test_worksheet_not_found_custom_workbook_name` : Nom de classeur personnalisé
3. `test_worksheet_not_found_empty_sheet_name` : Nom vide
4. `test_worksheet_not_found_inheritance` : Héritage correct

**TestWorksheetAlreadyExistsError** (3 tests)
1. `test_worksheet_already_exists_default_message` : Message par défaut
2. `test_worksheet_already_exists_with_spaces` : Nom avec espaces
3. `test_worksheet_already_exists_inheritance` : Héritage correct

**TestWorksheetDeleteError** (3 tests)
1. `test_worksheet_delete_error_default_reason` : Raison par défaut
2. `test_worksheet_delete_last_visible_sheet` : Dernière feuille visible
3. `test_worksheet_delete_error_inheritance` : Héritage correct

**TestWorksheetNameError** (5 tests)
1. `test_worksheet_name_error_default_reason` : Raison par défaut
2. `test_worksheet_name_error_too_long` : Nom trop long
3. `test_worksheet_name_error_invalid_character` : Caractère invalide
4. `test_worksheet_name_error_inheritance` : Héritage correct

### Tests

Exécution des tests :
```bash
# Tests des nouvelles exceptions
poetry run pytest tests/test_exceptions.py::TestWorksheetNotFoundError -v
poetry run pytest tests/test_exceptions.py::TestWorksheetAlreadyExistsError -v
poetry run pytest tests/test_exceptions.py::TestWorksheetDeleteError -v
poetry run pytest tests/test_exceptions.py::TestWorksheetNameError -v

# Tous les tests
poetry run pytest tests/ -v --no-cov
# Résultat : 186 passed, 1 xfailed
```

**Note** : Le test xfailed (`test_sample_failing`) est attendu et n'est pas lié à cette story.
