# Optimiseurs Excel - Performance et RAII

## Imports

```python
from xlmanage import (
    ExcelOptimizer,
    ScreenOptimizer,
    CalculationOptimizer,
    OptimizationState
)
```

## Types d'Optimiseurs

### ScreenOptimizer - Affichage

Désactive les mises à jour d'écran pour accélérer les opérations visuelles.

| Propriété | Avant Optimisation | Après Optimisation |
|-----------|-------------------|-------------------|
| ScreenUpdating | True (activé) | False (désactivé) |
| DisplayStatusBar | True (visible) | False (caché) |
| EnableAnimations | True (activé) | False (désactivé) |

### CalculationOptimizer - Calcul

Désactive le calcul automatique pour les opérations massives.

| Propriété | Avant Optimisation | Après Optimisation |
|-----------|-------------------|-------------------|
| Calculation | Automatic (-4105) | Manual (-4135) |
| Iteration | True/False (user) | False (désactivé) |
| MaxIterations | Défaut Excel | 100 |
| MaxChange | Défaut Excel | 0.0001 |

### ExcelOptimizer - Optimisation Complète

Combine `ScreenOptimizer` + `CalculationOptimizer` + événements.

| Propriété | Avant Optimisation | Après Optimisation |
|-----------|-------------------|-------------------|
| ScreenUpdating | True | False |
| DisplayStatusBar | True | False |
| EnableAnimations | True | False |
| Calculation | Automatic | Manual |
| EnableEvents | True | False |
| DisplayAlerts | False (user) | False (désactivé) |
| AskToUpdateLinks | True (user) | False (désactivé) |

## Usage avec Context Manager (Recommandé)

### ScreenOptimizer

```python
from xlmanage import ExcelManager, ScreenOptimizer

with ExcelManager(visible=False) as mgr:
    mgr.start()

    with ScreenOptimizer(mgr):
        # Écran optimisé ici
        for i in range(1000):
            ws.Cells(i, 1).Value = i
    # Paramètres restaurés automatiquement
```

### CalculationOptimizer

```python
from xlmanage import ExcelManager, CalculationOptimizer

with ExcelManager(visible=False) as mgr:
    mgr.start()

    with CalculationOptimizer(mgr):
        # Calcul manuel activé
        wb = mgr.app.Workbooks.Open("complex.xlsx")
        # Modifications sans recalcul
        mgr.app.Calculate()  # Recalcul manuel quand voulu
    # Paramètres restaurés automatiquement
```

### ExcelOptimizer (Complet)

```python
from xlmanage import ExcelManager, ExcelOptimizer

with ExcelManager(visible=False) as mgr:
    mgr.start()

    with ExcelOptimizer(mgr):
        # Excel complètement optimisé ici
        wb = mgr.app.Workbooks.Open("large_file.xlsx")

        for ws in wb.Worksheets:
            # Opérations massives sans ralentissements
            ws.Range("A1:Z1000").Value = data

        mgr.app.Calculate()  # Recalcul manuel final
    # Tous les paramètres restaurés automatiquement
```

## Usage Manuel avec apply/restore

Quand vous avez besoin d'un contrôle précis du moment d'application/restauration :

### ScreenOptimizer Manuel

```python
from xlmanage import ExcelManager, ScreenOptimizer

with ExcelManager(visible=False) as mgr:
    mgr.start()

    optimizer = ScreenOptimizer(mgr)

    # Sauvegarder et appliquer optimisation
    state = optimizer.apply()

    # ... opérations avec écran optimisé ...

    # Restaurer quand vous voulez
    optimizer.restore()
```

### ExcelOptimizer Manuel

```python
from xlmanage import ExcelManager, ExcelOptimizer

with ExcelManager(visible=False) as mgr:
    mgr.start()

    optimizer = ExcelOptimizer(mgr)

    # État sauvegardé
    state = optimizer.apply()

    try:
        # ... opérations optimisées ...

        # Recalculer manuellement
        mgr.app.Calculate()

    finally:
        # Toujours restaurer
        optimizer.restore()
```

## Inspection des Paramètres Courants

```python
from xlmanage import ExcelManager, ExcelOptimizer

with ExcelManager(visible=False) as mgr:
    mgr.start()

    optimizer = ExcelOptimizer(mgr)

    # Lire paramètres actuels sans modifier
    settings = optimizer.get_current_settings()

    print("Current Excel settings:")
    for prop, value in settings.items():
        print(f"  {prop}: {value}")
```

## OptimizationState Structure

```python
dataclass OptimizationState:
    screen: dict[str, object]      # État propriétés écran (ScreenOptimizer)
    calculation: dict[str, object]  # État propriétés calcul (CalculationOptimizer)
    full: dict[str, object]        # État complet (ExcelOptimizer)
    applied_at: str                # Timestamp ISO application
    optimizer_type: str            # "screen", "calculation", "all"
```

## Patterns Avancés

### Optimisation Conditionnelle

```python
def should_optimize(file_size_mb: float) -> bool:
    """Détermine si l'optimisation est nécessaire"""
    return file_size_mb > 5  # Optimiser > 5 MB

with ExcelManager(visible=False) as mgr:
    mgr.start()
    file_path = Path("data.xlsx")
    file_size = file_path.stat().st_size / (1024 * 1024)  # MB

    if should_optimize(file_size):
        with ExcelOptimizer(mgr):
            wb = mgr.app.Workbooks.Open(str(file_path))
            # ... traitement optimisé ...
    else:
        wb = mgr.app.Workbooks.Open(str(file_path))
        # ... traitement normal ...
```

### Optimisation Nested (Scoping)

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()

    with ScreenOptimizer(mgr):
        # Écran optimisé pour toutes les opérations

        # Calcul optimisé uniquement pour une section
        with CalculationOptimizer(mgr):
            wb = mgr.app.Workbooks.Open("complex.xlsx")
            # ... calculs massifs ...
            mgr.app.Calculate()

        # Retour à calcul automatique, écran toujours optimisé

        wb2 = mgr.app.Workbooks.Open("simple.xlsx")
        # ... calculs normaux ...
    # Écran restauré
```

### Mesurer Impact Performance

```python
import time

def process_with_timing(with_optimizer: bool) -> float:
    """Mesure le temps de traitement"""
    with ExcelManager(visible=False) as mgr:
        mgr.start()

        start = time.time()

        if with_optimizer:
            with ExcelOptimizer(mgr):
                # ... opérations ...
                pass
        else:
            # ... opérations ...
            pass

        return time.time() - start

time_normal = process_with_timing(False)
time_optimized = process_with_timing(True)

print(f"Normal: {time_normal:.2f}s")
print(f"Optimized: {time_optimized:.2f}s")
print(f"Speedup: {time_normal/time_optimized:.1f}x")
```

## Anti-Patterns

### ❌ Oublier restauration (sans context manager)

```python
# MAUVAIS - restauration manuelle facile à oublier
optimizer = ExcelOptimizer(mgr)
optimizer.apply()
# ... opérations ...
# OUPS: oublié optimizer.restore() !
```

### ✅ Utiliser context manager

```python
# BON - restauration automatique garantie
with ExcelOptimizer(mgr):
    # ... opérations ...
# Restauration automatique
```

### ❌ Optimiser pour opérations triviales

```python
# MAUVAIS - surcharge inutile
with ExcelOptimizer(mgr):
    ws.Range("A1").Value = "Hello"  # Opération trop simple
```

### ✅ Optimiser pour opérations massives

```python
# BON - gain réel
with ExcelOptimizer(mgr):
    for i in range(1000):
        ws.Cells(i, 1).Value = data[i]
```

## Documentation API

Pour voir l'API complète, utiliser les docstrings Python :

```python
from xlmanage import ExcelOptimizer, ScreenOptimizer, CalculationOptimizer
import inspect

print(inspect.getdoc(ExcelOptimizer))
print(inspect.getdoc(ScreenOptimizer))
print(inspect.getdoc(CalculationOptimizer))
```
