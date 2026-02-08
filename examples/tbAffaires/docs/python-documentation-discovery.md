# Procédure méthodologique : Lecture de la documentation d'un paquet Python

## Objectif

Guider un agent IA pour extraire efficacement la doc d'un paquet Python (classes, fonctions, modules) via introspection runtime.

## Étapes séquentielles

### 1. Vérification installation

```bash
poetry show <paquet>  # Confirme version/install
```

### 2. Exploration module racine

```bash
poetry run python -c "import <paquet>; print(dir(<paquet>)); help(<paquet>)"
```

- `dir()` : Liste attributs/classes/fonctions
- `help()` : Doc complète du module

### 3. Docstring spécifique (classe/fonction)

```bash
poetry run python -c "import <paquet>; print(<paquet>.Classe.__doc__)"
poetry run python -c "from <paquet>.module import fonction; help(fonction)"
```

### 4. Inspection avancée (inspect)

```python
import inspect
print(inspect.getdoc(<paquet>.Classe))
print(inspect.signature(<paquet>.Classe.__init__))
print(inspect.getsource(<paquet>.Classe.methode))  # Source si disponible
```

### 5. Recherche ciblée (par besoin)

| Besoin               | Commande                                                                                                            |
| -------------------- | ------------------------------------------------------------------------------------------------------------------- |
| Attributs classe     | `poetry run python -c "import <paquet>; print(dir(<paquet>.Classe))"`                                               |
| Signature `__init__` | `poetry run python -c "import inspect; print(inspect.signature(<paquet>.Classe.__init__))"`                         |
| Doc complète         | `poetry run python -c "from <paquet> import Classe; help(Classe)"`                                                  |
| Instances running    | `poetry run python -c "from <paquet>.excel_manager import ExcelManager; help(ExcelManager.list_running_instances)"` |

## Exemple : xlmanage.ExcelManager

```bash
poetry run python -c "from xlmanage.excel_manager import ExcelManager, InstanceInfo; help(ExcelManager); print(InstanceInfo.__doc__)"
```

**Résultat** : RAII context manager, méthodes `start/stop/list_running_instances`, attributs `InstanceInfo(pid, visible, workbooks_count, hwnd)`.

## Bonnes pratiques

- **Batch** : Multiple `help()` en parallèle via plusieurs tool calls
- **Erreurs** : `ModuleNotFoundError` → `poetry add <paquet>`
- **Source** : `inspect.getsource()` pour code complet
- **Cache** : Note docstrings clés dans AGENTS.md pour usage fréquent
