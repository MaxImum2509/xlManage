# Epic 13 - Story 3: Refactoriser les 3 optimizers pour injection de dependances

**Statut** : A faire

**Priorite** : P1 (OPT-001) + P2 (OPT-002, OPT-003)

**En tant que** mainteneur du projet
**Je veux** que les 3 optimizers acceptent `ExcelManager` au lieu de `CDispatch`
**Afin de** respecter le pattern d'injection de dependances et completer les donnees retournees

## Contexte

L'audit du 2026-02-06 a identifie 3 anomalies partagees par les 3 fichiers d'optimisation. La plus critique est l'injection de dependances : l'architecture (section 3.2) exige que "chaque manager recoit un `ExcelManager` en parametre". Actuellement, les 3 optimizers acceptent directement un `CDispatch` (`app`).

**Reference architecture** : section 4.8 (`_dev/architecture.md`, lignes 1277-1332)

## Anomalies a corriger

| ID      | Severite  | Fichiers impactes                                    | Description                                  |
| ------- | --------- | ---------------------------------------------------- | -------------------------------------------- |
| OPT-001 | Critique  | `excel_optimizer.py`, `screen_optimizer.py`, `calculation_optimizer.py` | `__init__` accepte `CDispatch` au lieu de `ExcelManager` |
| OPT-002 | Important | `excel_optimizer.py`                                  | `apply()` retourne `OptimizationState` avec `screen={}` et `calculation={}` vides |
| OPT-003 | Important | `excel_optimizer.py`                                  | `ExcelOptimizer` ne gere pas `MaxIterations` et `MaxChange` |

## Taches techniques

### Tache 3.1 : Modifier `ExcelOptimizer.__init__` (OPT-001)

**Fichier** : `src/xlmanage/excel_optimizer.py:69`

**Avant** :
```python
def __init__(self, app: CDispatch) -> None:
    self._app = app
    self._original_settings: dict[str, Any] = {}
```

**Apres** :
```python
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .excel_manager import ExcelManager

# ...

def __init__(self, excel_manager: "ExcelManager") -> None:
    self._mgr = excel_manager
    self._app = excel_manager.app
    self._original_settings: dict[str, Any] = {}
```

**Note** : On conserve `self._app` en interne pour ne pas changer tout le code des methodes `_save_current_settings()`, `_apply_optimizations()`, etc. On ajoute `self._mgr` pour la coherence architecturale.

### Tache 3.2 : Modifier `ScreenOptimizer.__init__` (OPT-001)

**Fichier** : `src/xlmanage/screen_optimizer.py:49`

Meme changement que Tache 3.1 :
```python
def __init__(self, excel_manager: "ExcelManager") -> None:
    self._mgr = excel_manager
    self._app = excel_manager.app
    self._original_settings: dict[str, Any] = {}
```

### Tache 3.3 : Modifier `CalculationOptimizer.__init__` (OPT-001)

**Fichier** : `src/xlmanage/calculation_optimizer.py:51`

Meme changement que Tache 3.1 :
```python
def __init__(self, excel_manager: "ExcelManager") -> None:
    self._mgr = excel_manager
    self._app = excel_manager.app
    self._original_settings: dict[str, Any] = {}
```

### Tache 3.4 : Remplir `screen` et `calculation` dans `ExcelOptimizer.apply()` (OPT-002)

**Fichier** : `src/xlmanage/excel_optimizer.py`

**Avant** :
```python
def apply(self) -> OptimizationState:
    self._save_current_settings()
    self._apply_optimizations()
    return OptimizationState(
        screen={},         # VIDE
        calculation={},    # VIDE
        full=self._original_settings.copy(),
        applied_at=datetime.now().isoformat(),
        optimizer_type="all",
    )
```

**Apres** :
```python
def apply(self) -> OptimizationState:
    self._save_current_settings()
    self._apply_optimizations()

    # Extraire les sous-ensembles screen et calculation
    screen_keys = {"ScreenUpdating", "DisplayStatusBar", "EnableAnimations"}
    calc_keys = {"Calculation", "Iteration", "MaxIterations", "MaxChange"}

    return OptimizationState(
        screen={k: v for k, v in self._original_settings.items() if k in screen_keys},
        calculation={k: v for k, v in self._original_settings.items() if k in calc_keys},
        full=self._original_settings.copy(),
        applied_at=datetime.now().isoformat(),
        optimizer_type="all",
    )
```

### Tache 3.5 : Ajouter `MaxIterations` et `MaxChange` a `ExcelOptimizer` (OPT-003)

**Fichier** : `src/xlmanage/excel_optimizer.py`

L'optimizer "complet" ne gere actuellement que 8 proprietes. Il manque `MaxIterations` et `MaxChange` (proprietes de calcul).

**Dans `_save_current_settings()`** :
```python
self._original_settings = {
    "ScreenUpdating": self._app.ScreenUpdating,
    "DisplayStatusBar": self._app.DisplayStatusBar,
    "EnableAnimations": self._app.EnableAnimations,
    "Calculation": self._app.Calculation,
    "EnableEvents": self._app.EnableEvents,
    "DisplayAlerts": self._app.DisplayAlerts,
    "AskToUpdateLinks": self._app.AskToUpdateLinks,
    "Iteration": self._app.Iteration,
    "MaxIterations": self._app.MaxIterations,    # NOUVEAU
    "MaxChange": self._app.MaxChange,            # NOUVEAU
}
```

**Dans `get_current_settings()`** - meme ajout des 2 proprietes.

**Note** : `_apply_optimizations()` n'a pas besoin de modifier `MaxIterations` et `MaxChange` (ce ne sont pas des proprietes d'optimisation a proprement parler, mais elles doivent etre sauvegardees/restaurees).

### Tache 3.6 : Mettre a jour `cli.py` pour passer `ExcelManager` aux optimizers

**Fichier** : `src/xlmanage/cli.py`

**Avant** :
```python
screen_opt = ScreenOptimizer(app_com)
calc_opt = CalculationOptimizer(app_com)
excel_opt = ExcelOptimizer(app_com)
optimizer = ScreenOptimizer(app_com)
optimizer = CalculationOptimizer(app_com)
optimizer = ExcelOptimizer(app_com)
```

**Apres** :
```python
screen_opt = ScreenOptimizer(excel_mgr)
calc_opt = CalculationOptimizer(excel_mgr)
excel_opt = ExcelOptimizer(excel_mgr)
optimizer = ScreenOptimizer(excel_mgr)
optimizer = CalculationOptimizer(excel_mgr)
optimizer = ExcelOptimizer(excel_mgr)
```

**Impact** : Les fonctions `_display_optimization_status()` et `_restore_optimizations()` qui creent aussi des optimizers doivent recevoir `excel_mgr` au lieu de `app_com`.

### Tache 3.7 : Mettre a jour les tests

**Fichiers** : `tests/test_excel_optimizer.py`, `tests/test_screen_optimizer.py`, `tests/test_calculation_optimizer.py`

- Adapter les fixtures pour passer un mock `ExcelManager` au lieu d'un mock `CDispatch`
- Verifier que `apply()` retourne un `OptimizationState` avec `screen` et `calculation` remplis
- Verifier que `MaxIterations` et `MaxChange` sont sauvegardes/restaures

## Criteres d'acceptation

1. [ ] Les 3 optimizers acceptent `ExcelManager` en parametre
2. [ ] `ExcelOptimizer.apply()` retourne un `OptimizationState` avec `screen` et `calculation` remplis
3. [ ] `ExcelOptimizer` gere les 10 proprietes (8 + MaxIterations + MaxChange)
4. [ ] `cli.py` passe `excel_mgr` au lieu de `app_com` aux optimizers
5. [ ] Le context manager (`__enter__`/`__exit__`) continue de fonctionner
6. [ ] Tous les tests passent

## Dependances

- Depend de `ExcelManager` (deja existant)
- Impacte `cli.py` (appels aux constructeurs des optimizers)

## Definition of Done

- [ ] Les 3 anomalies OPT-001 a OPT-003 sont corrigees
- [ ] Tous les tests des 3 optimizers passent
- [ ] Tests CLI `optimize` passent
- [ ] Couverture > 90% pour les 3 fichiers
