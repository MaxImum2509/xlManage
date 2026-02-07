# Epic 13 - Story 6: Corrections mineures et entetes licence

**Statut** : A faire

**Priorite** : P2 (TEST-002) + P3 (CONF-001, CONF-002)

**En tant que** mainteneur du projet
**Je veux** corriger les anomalies mineures de configuration et ajouter les entetes de licence manquantes
**Afin de** garantir la conformite licence GPL v3 et la proprete du depot

## Contexte

L'audit du 2026-02-06 a identifie plusieurs anomalies mineures : une valeur incorrecte dans `pyproject.toml`, un fichier binaire suivi par erreur dans git, et 6 fichiers de test sans entete GPL.

## Anomalies a corriger

| ID       | Severite  | Description                                          |
| -------- | --------- | ---------------------------------------------------- |
| CONF-001 | Mineur    | `target-version = "py313"` au lieu de `"py314"`      |
| CONF-002 | Mineur    | `.coverage` (binaire) suivi par git                  |
| TEST-002 | Important | 6 fichiers de test sans entete GPL v3                |

## Taches techniques

### Tache 6.1 : Corriger `target-version` dans `pyproject.toml` (CONF-001)

**Fichier** : `pyproject.toml:53`

**Avant** :
```toml
target-version = "py313"
```

**Apres** :
```toml
target-version = "py314"
```

**Justification** : Le projet cible Python 3.14 (cf. `requires-python = ">=3.14"` a la ligne 11). La version de ruff doit correspondre.

**Note** : Cette modification est autorisee par l'exception [EXP-001] car c'est une configuration d'outil (`[tool.ruff]`), pas une dependance.

### Tache 6.2 : Retirer `.coverage` du suivi git (CONF-002)

**Commande** :
```bash
git rm --cached .coverage
```

Le fichier `.coverage` est un binaire SQLite genere par pytest-cov. Il est deja dans `.gitignore` mais a ete committe par erreur.

**Verification** : Verifier que `.coverage` est bien dans `.gitignore`. Si non, l'ajouter.

### Tache 6.3 : Ajouter les entetes GPL v3 aux 6 fichiers de test (TEST-002)

**Fichiers impactes** :
1. `tests/test_sample.py`
2. `tests/test_coverage.py`
3. `tests/test_excel_manager.py`
4. `tests/test_workbook_manager.py`
5. `tests/test_vba_utilities.py`
6. `tests/test_vba_manager_init.py`

**Entete a ajouter** (en tete de chaque fichier, avant les imports) :

```python
"""
[Description breve du fichier de test]

This file is part of xlManage.

xlManage is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

xlManage is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with xlManage.  If not, see <https://www.gnu.org/licenses/>.
"""
```

**Descriptions suggerees** :
- `test_sample.py` : "Basic sample tests for xlManage."
- `test_coverage.py` : "Coverage verification tests for xlManage."
- `test_excel_manager.py` : "Tests for ExcelManager lifecycle operations."
- `test_workbook_manager.py` : "Tests for WorkbookManager CRUD operations."
- `test_vba_utilities.py` : "Tests for VBA utility functions."
- `test_vba_manager_init.py` : "Tests for VBAManager initialization."

## Criteres d'acceptation

1. [ ] `target-version` est `"py314"` dans `pyproject.toml`
2. [ ] `.coverage` n'est plus suivi par git (mais present dans `.gitignore`)
3. [ ] Les 6 fichiers de test ont l'entete GPL v3
4. [ ] Tous les tests continuent de passer
5. [ ] `ruff check` passe sans erreur

## Dependances

- Aucune dependance bloquante
- Peut etre executee en parallele avec les autres stories

## Definition of Done

- [ ] Les 3 anomalies CONF-001, CONF-002, TEST-002 sont corrigees
- [ ] Tous les tests passent
- [ ] `git status` ne montre pas `.coverage` dans les fichiers suivis
