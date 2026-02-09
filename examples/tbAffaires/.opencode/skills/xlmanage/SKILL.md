# Skill: xlmanage-python

name: xlmanage-python
description: Guide for using xlManage Python package to control Excel via COM automation. Use when working with Excel files, VBA modules, ListObjects, or performance optimization in Python projects using xlManage.

# xlManage Python - Usage du Package

**xlManage** est un utilitaire CLI Python pour contrôler Excel via l'automatisation COM (pywin32). Cette skill fournit des guidelines pour l'utilisation Python de xlManage.

## 📋 Structure Modulaire

Cette skill est organisée en **sections modulaires chargeables à la demande** pour minimiser l'usage du contexte.

```
xlmanage-python/
├── SKILL.md                    # Ce fichier : Vue d'ensemble + index
├── references/
│   ├── 01-fondamentaux.md       # ExcelManager, lifecycle, RAII, COM basics
│   ├── 02-workbooks.md          # WorkbookManager, WorkbookInfo, CRUD
│   ├── 03-worksheets.md         # WorksheetManager, WorksheetInfo, CRUD
│   ├── 04-tables.md            # TableManager, TableInfo, ListObjects
│   ├── 05-vba.md               # VBAManager, MacroRunner, import/export
│   ├── 06-optimisation.md      # Optimiseurs, performances, RAII patterns
│   └── 07-exceptions.md        # Hiérarchie des exceptions, handling patterns
```

## 🎯 Quand Charger Chaque Section

| Section | Quand charger | Contenu clé |
|----------|---------------|-------------|
| **01-fondamentaux** | TOUTE interaction xlManage | ExcelManager, contexte COM, lifecycle |
| **02-workbooks** | Manipulation de fichiers .xlsx/.xlsm | WorkbookManager, open/save/close |
| **03-worksheets** | Manipulation de feuilles | WorksheetManager, création/suppression |
| **04-tables** | Opérations sur ListObjects | TableManager, CRUD tables |
| **05-vba** | Import/export/exécution VBA | VBAManager, MacroRunner |
| **06-optimisation** | Performance intensive | Optimiseurs, RAII patterns |
| **07-exceptions** | Error handling / debugging | Hiérarchie exceptions |

## 🚀 Quick Start Pattern

```python
from xlmanage import ExcelManager

# Pattern RAII standard (recommandé)
with ExcelManager(visible=False) as mgr:
    mgr.start()
    # ... vos opérations Excel ...
    # Fermeture automatique garantie
```

## 📚 Modules et Managers Principaux

### Modules Core (Toujours nécessaires)
- `ExcelManager` - Gestion lifecycle Excel
- `WorkbookManager` - CRUD classeurs
- `WorksheetManager` - CRUD feuilles
- `TableManager` - CRUD ListObjects
- `VBAManager` - CRUD modules VBA
- `MacroRunner` - Exécution macros

### Optimiseurs (Performance)
- `ScreenOptimizer` - Optimisation affichage
- `CalculationOptimizer` - Optimisation calcul
- `ExcelOptimizer` - Optimisation complète

### Exceptions (Error Handling)
- `ExcelManageError` - Base exception
- Spécialisées : Workbook*, Worksheet*, Table*, VBA*, Excel*...

## 📖 Accès aux Docstrings Python

Pour obtenir la documentation complète des fonctions xlManage, l'agent peut utiliser Python pour lire les docstrings :

```python
# Méthode 1 : help()
from xlmanage import ExcelManager
help(ExcelManager)

# Méthode 2 : inspect.getdoc()
import inspect
from xlmanage import ExcelManager
print(inspect.getdoc(ExcelManager))

# Méthode 3 : attribut __doc__
from xlmanage import ExcelManager
print(ExcelManager.__doc__)
```

Cette méthode permet d'accéder à la documentation la plus à jour sans dépendre de ressources externes.

## 🔗 Chargement des Sections

Pour charger une section spécifique, lisez le fichier correspondant dans `references/`.
