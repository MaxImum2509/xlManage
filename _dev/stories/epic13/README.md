# Epic 13 : Mise en conformite architecture

**Objectif** : Corriger les 24 anomalies identifiees par l'audit d'architecture du 2026-02-06 pour rendre le code 100% conforme a `_dev/architecture.md` v1.0.0.

**Rapport d'audit** : `_dev/reports/audit-architecture-2026-02-06.md`

---

## Resume

| Metrique                       | Avant audit     | Cible apres Epic 13 |
| ------------------------------ | --------------- | -------------------- |
| Modules conformes              | 5/11 (45%)      | 11/12 (92%)          |
| Commandes CLI                  | 20/21 (95%)     | 21/21 (100%)         |
| Tests en echec                 | 28              | 0                    |
| Couverture                     | 89.32%          | >= 90%               |
| Anomalies critiques restantes  | 18              | 0                    |
| Anomalies importantes restantes| 4               | 0                    |
| Anomalies mineures restantes   | 4               | 0                    |

---

## Stories

| Story | Titre                                          | Priorite | Anomalies          | Fichiers impactes                                    |
| ----- | ---------------------------------------------- | -------- | ------------------ | ---------------------------------------------------- |
| S1    | Mise en conformite `table_manager.py`          | P1       | TBL-001 a TBL-009  | `table_manager.py`, `cli.py`, tests                  |
| S2    | Mise en conformite `worksheet_manager.py`      | P1+P2    | WS-001 a WS-003    | `worksheet_manager.py`, tests                        |
| S3    | Refactoriser les 3 optimizers                  | P1+P2    | OPT-001 a OPT-003  | `excel_optimizer.py`, `screen_optimizer.py`, `calculation_optimizer.py`, `cli.py` |
| S4    | Commande `run-macro` CLI + `__init__.py`       | P0+P2    | CLI-001, INIT-001/2 | `cli.py`, `__init__.py`, tests                       |
| S5    | Corriger les 28 tests en echec                 | P2       | TEST-001            | `test_excel_manager.py`, `test_cli.py`               |
| S6    | Corrections mineures et entetes licence        | P2+P3    | CONF-001/2, TEST-002| `pyproject.toml`, `.coverage`, 6 fichiers de test    |

---

## Ordre d'execution recommande

```
S6 (corrections mineures)        ← independant, peut demarrer immediatement
  |
S2 (worksheet_manager.py)        ← pas de dependance, impacte peu de fichiers
  |
S1 (table_manager.py)            ← le plus complexe, 9 anomalies
  |
S3 (optimizers)                  ← changement transversal (3 fichiers + cli)
  |
S4 (run-macro CLI + __init__.py) ← depend de S3 pour les imports corrects
  |
S5 (corriger les tests)          ← en dernier, apres toutes les modifications de code
```

**Justification** :
- **S6 en premier** : corrections triviales sans risque de regression
- **S2 avant S1** : plus simple (3 anomalies vs 9), permet de valider l'approche
- **S1 avant S3** : `table_manager.py` est autonome, les optimizers impactent le CLI
- **S4 apres S3** : `__init__.py` doit importer les optimizers refactorises
- **S5 en dernier** : les corrections de code des S1-S4 vont changer les tests

**Parallelisation possible** : S6 peut etre fait en parallele avec S2. S1 et S2 sont independantes.

---

## Matrice anomalies -> stories

| ID       | Story | Severite  | Description                                              |
| -------- | ----- | --------- | -------------------------------------------------------- |
| TBL-001  | S1    | Critique  | `columns: list[str]` manquant dans TableInfo             |
| TBL-002  | S1    | Critique  | Noms de champs incorrects                                |
| TBL-003  | S1    | Critique  | `_find_table()` signature incorrecte                     |
| TBL-004  | S1    | Critique  | `_validate_range()` signature incorrecte                 |
| TBL-005  | S1    | Critique  | Pas de verification chevauchement                        |
| TBL-006  | S1    | Critique  | Extraction colonnes absente                              |
| TBL-007  | S1    | Critique  | `delete()` ignore `force`                                |
| TBL-008  | S1    | Critique  | Ordre parametres `create()` incorrect                    |
| TBL-009  | S1    | Mineur    | Notation regex differente                                |
| WS-001   | S2    | Critique  | `SHEET_NAME_FORBIDDEN_CHARS` incorrecte                  |
| WS-002   | S2    | Important | Type annotation manquante sur `__init__`                 |
| WS-003   | S2    | Important | Nettoyage COM incomplet                                  |
| OPT-001  | S3    | Critique  | Injection dependances absente (3 fichiers)               |
| OPT-002  | S3    | Important | `OptimizationState` mal renseignee                       |
| OPT-003  | S3    | Important | `MaxIterations`/`MaxChange` manquants                    |
| CLI-001  | S4    | Critique  | Commande `run-macro` manquante                           |
| INIT-001 | S4    | Important | Exports manquants dans `__init__.py`                     |
| INIT-002 | S4    | Mineur    | Aliases inutiles dans `__init__.py`                      |
| TEST-001 | S5    | Important | 28 tests en echec                                        |
| TEST-002 | S6    | Important | Entetes GPL manquantes (6 fichiers)                      |
| CONF-001 | S6    | Mineur    | `target-version` incorrect                               |
| CONF-002 | S6    | Mineur    | `.coverage` dans git                                     |

---

## Verification Epic 12

L'Epic 12 a ete testee le 2026-02-07 :

- **Story 1** (_parse_macro_args) : 15 tests OK
- **Story 2** (MacroRunner) : 16 tests OK
- **Story 3** (CLI run-macro) : A faire → integree dans cette Epic 13, Story 4

**Total Epic 12** : 31/31 tests passent (Stories 1 et 2 conformes).

---

## Estimation de charge

| Story | Complexite | Fichiers modifies | Tests a modifier/creer |
| ----- | ---------- | ----------------- | ---------------------- |
| S1    | Haute      | 3                 | ~20 tests a adapter    |
| S2    | Faible     | 1                 | ~5 tests a ajouter     |
| S3    | Moyenne    | 4                 | ~15 tests a adapter    |
| S4    | Moyenne    | 2                 | ~8 tests a creer       |
| S5    | Haute      | 2                 | 28 tests a corriger    |
| S6    | Faible     | 8                 | 0 test a modifier      |
