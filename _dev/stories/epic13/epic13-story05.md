# Epic 13 - Story 5: Corriger les 28 tests en echec

**Statut** : A faire

**Priorite** : P2 - Important

**En tant que** mainteneur du projet
**Je veux** que tous les tests passent au vert
**Afin de** garantir la stabilite du CI/CD et une couverture > 90%

## Contexte

L'audit du 2026-02-06 a identifie 28 tests en echec sur 525 (528 apres Epic 12). Les echecs se concentrent sur 2 fichiers :

- **17 echecs dans `test_excel_manager.py`** : fonctions utilitaires et cas limites stop
- **11 echecs dans `test_cli.py`** : commande stop et integration

La couverture actuelle est de 89.32% (seuil requis : 90%).

## Tests en echec

### `test_excel_manager.py` - 17 echecs

| Classe de test                           | Nombre | Cause probable                               |
| ---------------------------------------- | ------ | -------------------------------------------- |
| `TestExcelManagerNewMethods`             | 2      | `list_running_instances` ROT mock incorrect  |
| `TestExcelManagerStopEdgeCases`          | 3      | Gestion d'erreurs stop/del_app              |
| `TestListRunningInstancesEdgeCases`      | 3      | Fallback et erreurs multiples               |
| `TestUtilityFunctions`                   | 9      | `enumerate_excel_instances`, `connect_by_pid/hwnd`, `enumerate_excel_pids` |

### `test_cli.py` - 11 echecs

| Classe de test                           | Nombre | Cause probable                               |
| ---------------------------------------- | ------ | -------------------------------------------- |
| `TestStopCommand`                        | 9      | Mock incorrect de la commande stop           |
| `TestCLIIntegration`                     | 2      | Workflow start+stop                          |

## Taches techniques

### Tache 5.1 : Analyser les causes racines des echecs

Pour chaque test en echec :
1. Executer le test isolement avec `-vvs` pour voir le traceback complet
2. Identifier si la cause est un mock incorrect, une signature changee, ou un bug dans le code source
3. Classifier : correction du test vs correction du code

**Commande** :
```bash
pytest tests/test_excel_manager.py::TestUtilityFunctions -vvs --no-header
pytest tests/test_cli.py::TestStopCommand -vvs --no-header
```

### Tache 5.2 : Corriger les tests `TestUtilityFunctions` (9 tests)

**Fichier** : `tests/test_excel_manager.py`

Ces tests couvrent les fonctions module-level :
- `enumerate_excel_instances()` - enumeration via ROT
- `enumerate_excel_pids()` - fallback via tasklist
- `connect_by_hwnd()` - connexion par handle de fenetre
- `connect_by_pid()` - connexion par PID (si existante)

Verifier que les mocks correspondent aux signatures actuelles des fonctions dans `excel_manager.py`.

### Tache 5.3 : Corriger les tests `TestExcelManagerNewMethods` et `TestExcelManagerStopEdgeCases` (5 tests)

**Fichier** : `tests/test_excel_manager.py`

Verifier :
- `list_running_instances()` : le mock du ROT et la construction d'`InstanceInfo`
- `stop()` : la gestion des exceptions COM et la liberation des references

### Tache 5.4 : Corriger les tests `TestListRunningInstancesEdgeCases` (3 tests)

**Fichier** : `tests/test_excel_manager.py`

Cas limites :
- `get_instance_info` qui leve une exception
- Fallback `connect_by_hwnd` qui echoue
- Les deux methodes (ROT + fallback) echouent

### Tache 5.5 : Corriger les tests `TestStopCommand` (9 tests)

**Fichier** : `tests/test_cli.py`

Verifier que les mocks de la commande `stop` correspondent a :
- La signature actuelle de `ExcelManager.stop_instance()`
- La signature actuelle de `ExcelManager.stop_all()`
- La signature actuelle de `ExcelManager.list_running_instances()`
- La signature actuelle de `ExcelManager.force_kill()`

### Tache 5.6 : Corriger les tests `TestCLIIntegration` (2 tests)

**Fichier** : `tests/test_cli.py`

Tests de workflow `start` + `stop` qui necessitent des mocks coordonnes.

### Tache 5.7 : Verifier la couverture globale

Apres correction de tous les tests :
```bash
pytest --cov=src/ --cov-report=term --cov-fail-under=90
```

Objectif : >= 90% de couverture globale.

Si la couverture est encore insuffisante, identifier les modules sous le seuil et ajouter des tests cibles.

## Criteres d'acceptation

1. [ ] Les 28 tests en echec sont corriges et passent au vert
2. [ ] Aucune regression sur les tests existants (528+ tests passent)
3. [ ] La couverture globale atteint >= 90%
4. [ ] `cli.py` atteint >= 80% de couverture
5. [ ] `excel_manager.py` atteint >= 90% de couverture

## Dependances

- Doit etre executee APRES les Stories 1-4 (les modifications de code peuvent impacter les tests)
- En particulier, les changements dans `table_manager.py` (Story 1) et les optimizers (Story 3) vont necessiter des mises a jour de tests supplementaires

## Definition of Done

- [ ] 0 tests en echec
- [ ] Couverture >= 90%
- [ ] `pytest` sort avec exit code 0
