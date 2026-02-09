# Fondamentaux - ExcelManager et Lifecycle COM

## ExcelManager - Gestion Lifecycle Excel

### Context Manager RAII (Recommandé)

```python
from xlmanage import ExcelManager

# Pattern standard : ouverture et fermeture automatiques
with ExcelManager(visible=False) as mgr:
    mgr.start()
    # ... opérations Excel ...
# Fermeture automatique garantie à la sortie du with
```

### Méthodes Principales

| Méthode | Description | Usage |
|----------|-------------|-------|
| `start(new=False)` | Démarre/connexion Excel | `mgr.start()` |
| `stop(save=True)` | Arrête proprement l'instance | `mgr.stop()` |
| `stop_all(save=True)` | Arrête TOUTES les instances | `mgr.stop_all()` |
| `stop_instance(pid, save=True)` | Arrête instance par PID | `mgr.stop_instance(12345)` |
| `force_kill(pid)` | Force-kill brutal (last resort) | `mgr.force_kill(12345)` |
| `list_running_instances()` | Liste toutes les instances | `[inst for inst in ...]` |
| `get_running_instance()` | Instance active courante | `mgr.get_running_instance()` |
| `get_instance_info(app)` | Info sur instance COM | `mgr.get_instance_info(app)` |

### Propriété `app`

Accès au COM Application Excel :

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    app = mgr.app
    wb = app.Workbooks.Open("data.xlsx")
```

## ⚠️ Règles Critiques

### NEVER call `app.Quit()`

**JAMAIS appeler `Quit()`** - provoque RPC error (0x800706be).

**Pourquoi :**
1. `Quit()` termine Excel immédiatement
2. Python détient encore des références COM
3. GC libère ensuite sur processus mort → RPC fatal

**Solution :** Utiliser le context manager `with` ou `stop()`.

### RAII Pattern

```python
# CORRECT : RAII automatique
with ExcelManager(visible=False) as mgr:
    mgr.start()
    wb = mgr.app.Workbooks.Open("test.xlsx")

# INCORRECT : gestion manuelle fragile
mgr = ExcelManager(visible=False)
mgr.start()
wb = mgr.app.Workbooks.Open("test.xlsx")
mgr.stop()  # Facile à oublier en cas d'erreur
```

## Instance Management

### Dispatch vs DispatchEx

```python
# win32.Dispatch() - Réutilise ou crée instance via ROT (partagé)
mgr.start(new=False)  # Défaut : Dispatch()

# win32.DispatchEx() - Crée NOUVELLE instance isolée
mgr.start(new=True)   # DispatchEx : processus séparé
```

**Usage `new=True` :** Opérations parallèles pour éviter interférences inter-instances.

### Enumeration des Instances

```python
mgr = ExcelManager()
instances = mgr.list_running_instances()

for inst in instances:
    print(f"PID {inst.pid}: {inst.workbooks_count} classeurs, "
          f"visible={inst.visible}")
```

### InstanceInfo Structure

```python
dataclass InstanceInfo:
    pid: int              # Process ID
    visible: bool          # Visibilité Excel
    workbooks_count: int   # Nombre de classeurs ouverts
    hwnd: int             # Window handle (ID unique)
```

## Shutdown Protocol

### `stop(save=True)`

1. Désactive les alertes Excel
2. Ferme tous les classeurs (avec sauvegarde optionnelle)
3. Libère les références COM (`del`)
4. Force garbage collection
5. Set `_app = None`

### `stop_instance(pid, save=True)`

Connecte à l'instance via ROT/HWND, puis applique `stop()`.

### `stop_all(save=True)`

Énumère via ROT et applique `stop_instance()` pour chaque.

### `force_kill(pid)` ⚠️

**Dernier recours uniquement !** Termine brutalement via `taskkill /f /pid <pid>`.

```python
try:
    mgr.stop_instance(12345)
except ExcelRPCError:
    # Clean shutdown failed, force kill
    mgr.force_kill(12345)
```

## Error Handling

### Exceptions Principales

| Exception | Quand lever | Contient |
|-----------|-------------|----------|
| `ExcelConnectionError` | COM indisponible | hresult, message |
| `ExcelInstanceNotFoundError` | Instance introuvable | instance_id, message |
| `ExcelRPCError` | Instance déconnectée | hresult, message |
| `ExcelManageError` | Base toutes exceptions | message |

### Pattern d'Error Handling

```python
from xlmanage import ExcelManager, ExcelConnectionError, ExcelRPCError

try:
    with ExcelManager(visible=False) as mgr:
        mgr.start()
        # ... opérations ...
except ExcelConnectionError as e:
    print(f"COM unavailable: {e.message}, HRESULT=0x{e.hresult:X}")
except ExcelRPCError as e:
    print(f"RPC error (instance likely dead): 0x{e.hresult:X}")
except ExcelManageError as e:
    print(f"xlManage error: {e.message}")
```

## Documentation API

Pour voir l'API complète, utiliser les docstrings Python :

```python
from xlmanage import ExcelManager
import inspect
print(inspect.getdoc(ExcelManager))
```
