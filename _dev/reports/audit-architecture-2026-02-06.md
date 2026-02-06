# Audit de conformite architecture - xlManage

**Date :** 2026-02-06
**Auditeur :** Claude Opus 4.6
**Reference :** `_dev/architecture.md` v1.0.0

---

## Resume executif

| Module                     | Statut           | Anomalies                          |
| -------------------------- | ---------------- | ---------------------------------- |
| `exceptions.py`            | **CONFORME**     | 0                                  |
| `excel_manager.py`         | **CONFORME**     | 0                                  |
| `workbook_manager.py`      | **CONFORME**     | 0                                  |
| `worksheet_manager.py`     | **NON CONFORME** | 3 (1 critique, 2 importants)       |
| `table_manager.py`         | **NON CONFORME** | 9 (8 critiques, 1 mineur)          |
| `vba_manager.py`           | **CONFORME**     | 0                                  |
| `macro_runner.py`          | **ABSENT**       | Fichier manquant (Epic 12)         |
| `excel_optimizer.py`       | **NON CONFORME** | 3 (1 critique, 2 importants)       |
| `screen_optimizer.py`      | **NON CONFORME** | 1 (critique, meme cause)           |
| `calculation_optimizer.py` | **NON CONFORME** | 1 (critique, meme cause)           |
| `cli.py`                   | **NON CONFORME** | 1 (commande `run-macro` manquante) |
| `__init__.py`              | **INCOMPLET**    | Exports manquants                  |
| Tests                      | **PARTIEL**      | 28 tests KO, couverture 88.82%     |
| `pyproject.toml`           | **ANOMALIE**     | `target-version` incorrect         |

**Bilan global : 18 anomalies critiques, 4 anomalies importantes, 2 anomalies mineures**

---

## 1. exceptions.py - CONFORME

Toutes les 22 classes d'exceptions requises sont presentes avec :

- Heritage correct depuis `ExcelManageError`
- Parametres `__init__` conformes a la spec
- Attributs metier correctement stockes
- Entete licence GPL v3 presente

**Verdict : Aucune action requise.**

---

## 2. excel_manager.py - CONFORME

- Dataclass `InstanceInfo` avec les 4 champs requis (pid, visible, workbooks_count, hwnd)
- Pattern RAII (`__enter__`/`__exit__`) correctement implemente
- **Aucun appel a `app.Quit()`** (verifie)
- Liberation ordonnee des references COM (del + gc.collect)
- `Dispatch()` vs `DispatchEx()` correctement utilise
- ctypes pour extraction PID via `GetWindowThreadProcessId`
- 3 fonctions utilitaires module-level presentes (enumerate_excel_instances, enumerate_excel_pids, connect_by_hwnd)
- Entete licence GPL v3 presente

**Verdict : Aucune action requise.**

---

## 3. workbook_manager.py - CONFORME

- Dataclass `WorkbookInfo` avec les 5 champs requis
- Constante `FILE_FORMAT_MAP` correcte (51, 52, 56, 50)
- Toutes les methodes CRUD presentes (open, create, close, save, list)
- `Path.exists()` verifie avant open
- Verification "deja ouvert" via `_find_open_workbook()`
- `Path.resolve()` systematique avant appels COM
- Injection de dependances via `ExcelManager`
- Entete licence GPL v3 presente

**Verdict : Aucune action requise.**

---

## 4. worksheet_manager.py - NON CONFORME (3 anomalies)

### [WS-001] CRITIQUE : Constante SHEET_NAME_FORBIDDEN_CHARS incorrecte

**Fichier :** `worksheet_manager.py:37`
**Attendu :** `r'\/*?:[]'`
**Actuel :** `r"\\/\*\?:\[\]"`

La double echappement dans une raw string rend le pattern regex incorrect. Les caracteres interdits (`/`, `*`, `?`, etc.) ne seront pas detectes par `_validate_sheet_name()`. Des noms de feuilles invalides passent la validation et provoqueront des erreurs COM.

### [WS-002] IMPORTANT : Type annotation manquante sur `__init__`

**Fichier :** `worksheet_manager.py:212`
**Attendu :** `def __init__(self, excel_manager: ExcelManager):`
**Actuel :** `def __init__(self, excel_manager):`

Non conforme au pattern d'injection de dependances documente dans l'architecture (section 3.2).

### [WS-003] IMPORTANT : Nettoyage COM incomplet

Les methodes `create()` et `copy()` retournent des resultats sans `del ws` prealable. Seule `delete()` fait correctement `del ws`. L'architecture exige la liberation ordonnee (Annexe C, regle 2).

---

## 5. table_manager.py - NON CONFORME (9 anomalies)

### [TBL-001] CRITIQUE : Champ `columns: list[str]` manquant dans TableInfo

**Attendu :** `columns: list[str]` contenant les noms des colonnes d'en-tete
**Actuel :** Champ absent de la dataclass

### [TBL-002] CRITIQUE : Noms de champs incorrects dans TableInfo

| Champ attendu   | Champ actuel       |
| --------------- | ------------------ |
| `range_address` | `range_ref`        |
| `header_row`    | `header_row_range` |

### [TBL-003] CRITIQUE : Signature `_find_table()` incorrecte

**Attendu :** `_find_table(wb: CDispatch, name: str) -> tuple[CDispatch, CDispatch] | None`
Recherche dans tout le classeur (noms uniques au classeur), retourne `(worksheet, table)`.

**Actuel :** `_find_table(ws: CDispatch, name: str) -> CDispatch | None`
Recherche dans une seule feuille, retourne seulement le ListObject.

### [TBL-004] CRITIQUE : Signature `_validate_range()` incorrecte

**Attendu :** `_validate_range(ws: CDispatch, range_ref: str) -> CDispatch`
Doit recevoir la feuille COM, verifier les chevauchements, et retourner l'objet Range COM.

**Actuel :** `_validate_range(range_ref: str) -> None`
Validation purement textuelle (regex), sans verification COM ni chevauchement.

### [TBL-005] CRITIQUE : Pas de verification de chevauchement dans `create()`

L'architecture exige (section 4.5, ligne 960) que `_validate_range()` verifie que la plage ne chevauche pas un ListObject existant. Cette verification est absente.

### [TBL-006] CRITIQUE : Extraction des colonnes absente dans `_get_table_info()`

La methode ne remplit pas le champ `columns` avec `[col.Name for col in table.ListColumns]` comme requis.

### [TBL-007] CRITIQUE : `delete()` ignore le parametre `force`

**Attendu :**

- `force=False` : `table.Unlist()` (conserve les donnees, supprime la structure)
- `force=True` : `table.Delete()` (supprime table ET donnees)

**Actuel :** Toujours `table.Delete()` (ligne 325), quel que soit `force`.

### [TBL-008] CRITIQUE : Ordre des parametres incorrect dans `create()`

**Attendu :** `create(name, range_ref, workbook, worksheet)`
**Actuel :** `create(name, range_ref, worksheet, workbook)`

Incohérent avec les autres managers qui utilisent tous `workbook` avant `worksheet`.

### [TBL-009] MINEUR : Notation regex legerement differente

**Attendu :** `r'^[A-Za-z_][A-Za-z0-9_]*$'`
**Actuel :** `r"^[a-zA-Z_][a-zA-Z0-9_]*$"`

Fonctionnellement equivalent mais inconsistant avec la spec.

---

## 6. vba_manager.py - CONFORME

- Toutes les constantes presentes (VBEXT*CT*\*, VBA_TYPE_NAMES, EXTENSION_TO_TYPE, VBA_ENCODING)
- Dataclass `VBAModuleInfo` avec les 4 champs requis
- Process d'import classe en 5 etapes correctement implemente
- Export document (Type=100) via `CodeModule.Lines()` au lieu de `Export()`
- Modules document non supprimables (Type=100)
- Gestion erreur Trust Center (0x800A03EC -> VBAProjectAccessError)
- Encodage windows-1252 applique
- Verification format .xlsm
- Entete licence GPL v3 presente

**Verdict : Aucune action requise.**

---

## 7. macro_runner.py - ABSENT

**Le fichier `src/xlmanage/macro_runner.py` n'existe pas.**

L'architecture (section 4.7) requiert :

- Dataclass `MacroResult` avec 5 champs
- Classe `MacroRunner` avec methode `run()`
- 3 fonctions utilitaires : `_build_macro_reference()`, `_parse_macro_args()`, `_format_return_value()`

**Impact :** Epic 12 entierement non implemente. Commande CLI `run-macro` impossible sans ce module.

---

## 8. Optimizers - NON CONFORME (3 anomalies partagees)

### [OPT-001] CRITIQUE : Injection de dependances absente (3 fichiers)

Les 3 optimizers acceptent `CDispatch` au lieu de `ExcelManager` :

| Fichier                       | Actuel                           | Attendu                                       |
| ----------------------------- | -------------------------------- | --------------------------------------------- |
| `excel_optimizer.py:69`       | `__init__(self, app: CDispatch)` | `__init__(self, excel_manager: ExcelManager)` |
| `screen_optimizer.py:49`      | `__init__(self, app: CDispatch)` | `__init__(self, excel_manager: ExcelManager)` |
| `calculation_optimizer.py:51` | `__init__(self, app: CDispatch)` | `__init__(self, excel_manager: ExcelManager)` |

L'architecture (section 3.2) est explicite : "chaque manager recoit un `ExcelManager` en parametre".
La section 4.8 precise : "Dependencies: `ExcelManager` (au lieu de `gencache.EnsureDispatch` autonome)".

### [OPT-002] IMPORTANT : OptimizationState mal renseignee dans ExcelOptimizer

`apply()` retourne un `OptimizationState` avec `screen={}` et `calculation={}` vides au lieu de les remplir avec les proprietes correspondantes.

### [OPT-003] IMPORTANT : ExcelOptimizer manque MaxIterations et MaxChange

L'optimizer "complet" ne gere que 6 des 10 proprietes attendues. Il manque `MaxIterations` et `MaxChange` (proprietes de calcul).

---

## 9. cli.py - NON CONFORME (1 anomalie)

### [CLI-001] CRITIQUE : Commande `run-macro` manquante

L'arbre des commandes (section 4.9, ligne 1367) specifie :

```
run-macro  MACRO_NAME  --workbook  --args  --timeout
```

20 des 21 commandes sont implementees (95%). Seule `run-macro` est absente.
Les 20 commandes existantes sont conformes : couche mince, formatage Rich, gestion d'exceptions.

---

## 10. **init**.py - INCOMPLET

### Exports manquants

| Element                | Dans **all** ? | Importe ?            |
| ---------------------- | -------------- | -------------------- |
| `VBAManager`           | Non            | Non                  |
| `VBAModuleInfo`        | Non            | Non                  |
| `MacroRunner`          | Non            | Non (fichier absent) |
| `MacroResult`          | Non            | Non (fichier absent) |
| `ExcelOptimizer`       | Non            | Non                  |
| `ScreenOptimizer`      | Non            | Non                  |
| `CalculationOptimizer` | Non            | Non                  |
| `OptimizationState`    | Non            | Non                  |

### Aliases inutiles

Les imports utilisent des aliases confus (`WorkbookInfoClass`, `WorksheetInfoData`, `TableInfoData`) avant d'etre re-assignes. Import direct preferable.

### Entete licence manquante

Le fichier `__init__.py` n'a pas l'entete GPL v3 requise par CLAUDE.md.

---

## 11. Tests - PARTIEL

### Tests en echec : 28 / 525

- **17 echecs dans `test_excel_manager.py`** : Fonctions utilitaires et cas limites stop
- **11 echecs dans `test_cli.py`** : Commande stop et integration

### Couverture : 88.82% (seuil requis : 90%)

Modules sous le seuil :

- `cli.py` : 78% (tire vers le bas par les commandes non testables)
- `excel_manager.py` : 87%

### Entetes licence manquantes (6 fichiers)

- `tests/test_sample.py`
- `tests/test_coverage.py`
- `tests/test_excel_manager.py`
- `tests/test_workbook_manager.py`
- `tests/test_vba_utilities.py`
- `tests/test_vba_manager_init.py`

### Tests manquants

- Aucun test pour `macro_runner.py` (fichier source absent)
- Aucun test CLI pour `run-macro` (commande absente)

---

## 12. pyproject.toml - ANOMALIE

### [CONF-001] MINEUR : target-version incorrect

**Fichier :** `pyproject.toml:53`
**Actuel :** `target-version = "py313"`
**Attendu :** `target-version = "py314"` (le projet cible Python 3.14)

### [CONF-002] MINEUR : .coverage dans git

Le fichier `.coverage` (binaire SQLite) est suivi par git malgre sa presence dans `.gitignore`. Necessite `git rm --cached .coverage`.

---

## Plan d'actions par priorite

### P0 - Bloquant (module absent)

| ID        | Action                                     | Impact                               |
| --------- | ------------------------------------------ | ------------------------------------ |
| MACRO-001 | Creer `macro_runner.py` (Epic 12)          | Module entier manquant               |
| CLI-001   | Ajouter commande `run-macro` dans `cli.py` | Fonctionnalite utilisateur manquante |

### P1 - Critique (comportement incorrect)

| ID      | Action                                                                      | Impact                |
| ------- | --------------------------------------------------------------------------- | --------------------- |
| TBL-001 | Ajouter `columns: list[str]` a `TableInfo`                                  | Schema incomplet      |
| TBL-002 | Renommer `range_ref` -> `range_address`, `header_row_range` -> `header_row` | API non conforme      |
| TBL-003 | Reecrire `_find_table()` pour chercher dans tout le classeur                | Recherche limitee     |
| TBL-004 | Reecrire `_validate_range()` avec parametre `ws` et retour `CDispatch`      | Pas de validation COM |
| TBL-005 | Ajouter verification chevauchement dans `create()`                          | Risque de corruption  |
| TBL-006 | Extraire colonnes dans `_get_table_info()`                                  | Donnees manquantes    |
| TBL-007 | Implementer `Unlist()` vs `Delete()` selon `force`                          | Perte de donnees      |
| TBL-008 | Corriger ordre parametres `create(name, range_ref, workbook, worksheet)`    | API inconsistante     |
| WS-001  | Corriger `SHEET_NAME_FORBIDDEN_CHARS`                                       | Validation cassee     |
| OPT-001 | Modifier les 3 optimizers pour accepter `ExcelManager`                      | Pattern architectural |

### P2 - Important (qualite / conformite)

| ID       | Action                                                                 | Impact                  |
| -------- | ---------------------------------------------------------------------- | ----------------------- |
| WS-002   | Ajouter type annotation `ExcelManager` sur `WorksheetManager.__init__` | Type checking           |
| WS-003   | Ajouter `del ws` dans `create()` et `copy()`                           | Fuite COM               |
| OPT-002  | Remplir `screen` et `calculation` dans `OptimizationState`             | Donnees incompletes     |
| OPT-003  | Ajouter `MaxIterations` / `MaxChange` a `ExcelOptimizer`               | Optimisation incomplete |
| TEST-001 | Corriger les 28 tests en echec                                         | CI/CD rouge             |
| TEST-002 | Ajouter entetes GPL a 6 fichiers de test                               | Licence                 |
| INIT-001 | Completer `__init__.py` (VBAManager, optimizers, etc.)                 | Exports publics         |

### P3 - Mineur

| ID       | Action                                        | Impact        |
| -------- | --------------------------------------------- | ------------- |
| TBL-009  | Harmoniser notation regex                     | Consistance   |
| CONF-001 | Corriger `target-version` dans pyproject.toml | Linting       |
| CONF-002 | Retirer `.coverage` du suivi git              | Proprete repo |
| INIT-002 | Supprimer aliases inutiles dans `__init__.py` | Lisibilite    |

---

## Statistiques

| Metrique               | Valeur                         |
| ---------------------- | ------------------------------ |
| Modules source audites | 11 / 12 (macro_runner absent)  |
| Modules conformes      | 5 / 11 (45%)                   |
| Tests totaux           | 525 (497 OK, 28 KO, 1 xfailed) |
| Couverture             | 88.82% (seuil : 90%)           |
| Anomalies critiques    | 18                             |
| Anomalies importantes  | 4                              |
| Anomalies mineures     | 4                              |

---

_Rapport genere le 2026-02-06 par Claude Opus 4.6_
