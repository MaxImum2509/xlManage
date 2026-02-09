# xlManage Python Skill

Documentation modulaire pour l'utilisation du package Python xlManage (contrôle Excel via COM automation).

## 📦 Installation de la Skill

La skill est enregistrée dans `.opencode/skills/xlmanage/` et sera automatiquement disponible pour les agents Claude Code.

## 📖 Structure

```
xlmanage/
├── SKILL.md                    # Point d'entrée (metadata + overview)
└── references/
    ├── 01-fondamentaux.md       # ExcelManager, lifecycle, RAII
    ├── 02-workbooks.md          # WorkbookManager, CRUD classeurs
    ├── 03-worksheets.md         # WorksheetManager, CRUD feuilles
    ├── 04-tables.md            # TableManager, ListObjects
    ├── 05-vba.md               # VBAManager, MacroRunner
    ├── 06-optimisation.md      # Optimiseurs, performances
    └── 07-exceptions.md        # Hiérarchie exceptions, handling
```

## 🎯 Chargement Intelligent

La skill utilise un système de **progressive disclosure** :

1. **SKILL.md** (~100 tokens) : Toujours chargé - Vue d'ensemble + index des sections
2. **References** (charge à la demande) : Chaque section est chargée uniquement quand nécessaire

Cette architecture minimise l'usage du contexte tout en maintenant une vision complète de la bibliothèque.

## 📚 Quand Charger Chaque Section

| Section | Scénario d'utilisation | Contenu clé |
|----------|----------------------|-------------|
| **01-fondamentaux** | TOUTE interaction xlManage | ExcelManager, lifecycle COM, RAII pattern |
| **02-workbooks** | Manipulation fichiers .xlsx/.xlsm | WorkbookManager, open/save/close classeurs |
| **03-worksheets** | Manipulation de feuilles | WorksheetManager, création/suppression feuilles |
| **04-tables** | Opérations ListObjects | TableManager, CRUD tables Excel |
| **05-vba** | Import/export/exécution VBA | VBAManager, MacroRunner |
| **06-optimisation** | Performance intensive | Optimiseurs, RAII patterns, calculs massifs |
| **07-exceptions** | Error handling / debugging | Hiérarchie complète des exceptions |

## 🚀 Quick Start

```python
from xlmanage import ExcelManager

# Pattern RAII standard (recommandé)
with ExcelManager(visible=False) as mgr:
    mgr.start()
    # ... vos opérations Excel ...
    # Fermeture automatique garantie
```

## 🔍 Navigation

Pour charger une section spécifique, lisez le fichier correspondant dans `references/` :

```bash
# Exemple : lire la section sur les workbooks
read ".opencode/skills/xlmanage/references/02-workbooks.md"
```

## 📖 Accès à la Documentation

Pour obtenir la documentation complète de xlManage, l'agent peut utiliser Python pour lire les docstrings :

```python
# Méthode 1 : help()
from xlmanage import ExcelManager
help(ExcelManager)

# Méthode 2 : inspect.getdoc()
import inspect
from xlmanage import ExcelManager
print(inspect.getdoc(ExcelManager))

# Méthode 3 : explorer tous les modules
from xlmanage import *
for name in dir():
    obj = eval(name)
    if hasattr(obj, '__doc__') and obj.__doc__:
        print(f"\n=== {name} ===")
        print(obj.__doc__)
```

Cette méthode garantit un accès fiable à la documentation la plus à jour sans dépendre de ressources externes.

## ⚠️ Règles Critiques

### 1. NEVER call `app.Quit()`

```python
# ❌ MAUVAIS - provoque RPC error
excel.Quit()

# ✅ BON - utiliser context manager
with ExcelManager(visible=False) as mgr:
    mgr.start()
```

### 2. Windows-1252 Encoding pour VBA

```python
# Tous les fichiers VBA doivent utiliser Windows-1252 avec CRLF
with open("module.bas", "w", encoding="windows-1252", newline="\r\n") as f:
    f.write(vba_code)
```

### 3. Toujours utiliser `with` statement

```python
# ❌ MAUVAIS - gestion manuelle fragile
mgr = ExcelManager()
mgr.start()
# ... risques d'oublier mgr.stop()

# ✅ BON - fermeture automatique garantie
with ExcelManager() as mgr:
    mgr.start()
```

## 🛠️ Modules Principaux

| Module | Responsabilité |
|---------|----------------|
| `ExcelManager` | Gestion lifecycle Excel (start/stop) |
| `WorkbookManager` | CRUD classeurs (open/save/close) |
| `WorksheetManager` | CRUD feuilles (create/delete/copy) |
| `TableManager` | CRUD ListObjects (tables Excel) |
| `VBAManager` | CRUD modules VBA (import/export) |
| `MacroRunner` | Exécution macros VBA (Sub/Function) |
| `ExcelOptimizer` | Optimisation performances complète |
| `ScreenOptimizer` | Optimisation affichage |
| `CalculationOptimizer` | Optimisation calcul |

## 📊 Statistiques

- **SKILL.md** : ~100 tokens (metadata + overview)
- **01-fondamentaux.md** : ~6000 tokens
- **02-workbooks.md** : ~5000 tokens
- **03-worksheets.md** : ~6000 tokens
- **04-tables.md** : ~7500 tokens
- **05-vba.md** : ~8500 tokens
- **06-optimisation.md** : ~7500 tokens
- **07-exceptions.md** : ~12000 tokens

**Total** : ~52500 tokens (seulement quand toutes les sections sont chargées)

En pratique, seulement 2-3 sections sont nécessaires par tâche typique, réduisant l'usage à ~15000-20000 tokens.

## 🤝 Contributeur

Pour mettre à jour la skill :

1. Modifier les fichiers `.md` dans `references/`
2. Mettre à jour le `description` dans `SKILL.md` si nécessaire
3. Tester avec des cas d'utilisation réels
4. Itérer en fonction des besoins

## 📄 License

Cette skill fait partie du projet xlManage.
