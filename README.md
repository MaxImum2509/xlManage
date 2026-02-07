# xlManage - Automatisation Excel par Ligne de Commande

[![Python 3.14+](https://img.shields.io/badge/python-3.14+-blue.svg)](https://www.python.org/downloads/)
[![License: GPL-3.0](https://img.shields.io/badge/license-GPL--3.0-green.svg)](https://www.gnu.org/licenses/gpl-3.0.html)
[![Coverage: 90%+](https://img.shields.io/badge/coverage-90%25%2B-brightgreen.svg)](#tests)
[![Poetry](https://img.shields.io/badge/dependency%20management-poetry-blue.svg)](https://python-poetry.org/)

xlManage est un outil CLI Windows en Python qui permet de piloter Microsoft Excel via l'automatisation COM (pywin32).
Il offre un contrôle programmatique complet sur Excel : démarrage/arrêt d'instances, gestion des classeurs, feuilles, tables, modules VBA et exécution de macros.

**Destiné aux agents LLM et développeurs** qui ont besoin d'une interface déclarative et robuste pour automatiser Excel.

---

## ✨ Fonctionnalités principales

### Gestion du cycle de vie Excel
- ✅ Démarrer/arrêter des instances Excel
- ✅ Énumérer les instances actives
- ✅ Contrôler la visibilité et les propriétés

### Opérations CRUD Classeurs
- ✅ Ouvrir/créer/fermer des classeurs
- ✅ Sauvegarder et exporter
- ✅ Lister les classeurs ouverts

### Gestion des feuilles de calcul
- ✅ Créer/supprimer/copier des feuilles
- ✅ Lister les feuilles avec infos (dimensions, visibilité)
- ✅ Validation des noms

### Tables Excel (ListObjects)
- ✅ Créer/supprimer des tables
- ✅ Lister les tables avec colonnes et données
- ✅ Validation des plages et unicité

### Automatisation VBA
- ✅ Importer/exporter modules VBA (.bas, .cls, .frm)
- ✅ Lister les modules avec métadonnées
- ✅ Supprimer les modules
- ✅ Gérer les UserForms

### Exécution de macros
- ✅ Exécuter Sub et Function VBA
- ✅ Passer des arguments typés (str, int, float, bool)
- ✅ Capturer les retours
- ✅ Gestion complète des erreurs VBA

### Optimisation de performances
- ✅ Désactiver les mises à jour écran
- ✅ Optimiser les calculs
- ✅ Désactiver les événements
- ✅ Modes avec/sans context manager

---

## 🚀 Installation

### Prérequis
- **Windows** (pywin32 et COM ne fonctionnent que sur Windows)
- **Python 3.14+**
- **Microsoft Excel** installé avec licence valide
- **Poetry** pour la gestion des dépendances

### Installation depuis PyPI
```bash
pip install xlmanage
```

### Installation en mode développement
```bash
git clone https://github.com/MaxImum2509/xlManage.git
cd xlManage
poetry install
poetry run xlmanage --help
```

---

## 📖 Utilisation rapide

### Commandes principales
```bash
# Démarrer une instance Excel
xlmanage start --visible

# Ouvrir un classeur
xlmanage workbook open C:\data\mon-fichier.xlsx

# Créer une nouvelle feuille
xlmanage worksheet create "Nouvelle feuille"

# Lister les feuilles
xlmanage worksheet list

# Créer une table
xlmanage table create "tbVentes" A1:D100 --worksheet "Données"

# Exécuter une macro
xlmanage run-macro Module1.MyMacro --args '"arg1",42,true'

# Arrêter proprement
xlmanage stop --save
```

### Exemple d'automatisation
```bash
# Scénario : ouvrir, optimiser, créer table, exécuter macro
xlmanage start --visible
xlmanage workbook open workbook.xlsm
xlmanage optimize --all
xlmanage table create "tbData" A1:Z1000 --worksheet "Import"
xlmanage run-macro "ProcessData" --timeout 60
xlmanage workbook save
xlmanage stop --save
```

---

## 🏗️ Architecture

xlManage suit une architecture modulaire en 3 couches :

```
┌─────────────────────────────────────┐
│        COUCHE CLI (cli.py)          │  ← Interface utilisateur (Typer + Rich)
├─────────────────────────────────────┤
│   Managers (6 modules)              │  ← Logique métier
│ • ExcelManager                      │
│ • WorkbookManager                   │
│ • WorksheetManager                  │
│ • TableManager                      │
│ • VBAManager                        │
│ • MacroRunner                       │
├─────────────────────────────────────┤
│   Optimizers (3 modules)            │  ← Optimisation performances
│ • ExcelOptimizer (8 propriétés)     │
│ • ScreenOptimizer (3 propriétés)    │
│ • CalculationOptimizer (4 propriétés)
├─────────────────────────────────────┤
│    pywin32 COM Bridge               │  ← Accès Excel
│   (Dispatch/DispatchEx)             │
├─────────────────────────────────────┤
│    Excel.exe (out-of-process)       │  ← Serveur COM
└─────────────────────────────────────┘
```

**Patterns clés** :
- **RAII** : Context managers pour garantir la libération des ressources COM
- **Injection de dépendances** : Chaque manager reçoit une instance `ExcelManager`
- **Exceptions typées** : Chaque erreur a sa classe spécifique avec contexte métier
- **CLI mince** : Aucune logique métier dans `cli.py`

---

## 📊 État du projet (V1.0.0)

| Métrique | Statut |
|----------|--------|
| **Tests** | ✅ 581 passing, couverture 90%+ |
| **Modules** | ✅ 12 modules Python (100% conformes) |
| **Commandes CLI** | ✅ 21 commandes (100% fonctionnelles) |
| **Documentation** | ✅ Sphinx avec 6 pages + API docs |
| **Linting** | ✅ ruff (E, F, W codes clean) |
| **Type checking** | ✅ mypy (strict mode) |
| **Pre-commit hooks** | ✅ Git hooks configurés |

---

## 🧪 Tests

### Exécuter les tests
```bash
poetry run pytest --cov=src/ --cov-report=html
```

### Résultats
```
581 tests passed, 1 xfailed
Coverage: 90.05% (seuil: 90%)
Temps total: ~25s
```

### Framework & outils
- **pytest** : Framework de test principal
- **pytest-cov** : Couverture de code
- **pytest-mock** : Injection de mocks
- **pytest-timeout** : Timeout par test (60s)
- **unittest.mock** : Mocks COM (pas de COM réel)

---

## 🛠️ Développement

### Structure du projet
```
xlManage/
├── src/xlmanage/           # Code source principal
│   ├── __init__.py         # Exports publics
│   ├── cli.py              # Interface Typer
│   ├── exceptions.py       # Exceptions typées
│   ├── excel_manager.py    # Gestion cycle de vie
│   ├── workbook_manager.py # CRUD classeurs
│   ├── worksheet_manager.py# CRUD feuilles
│   ├── table_manager.py    # CRUD tables
│   ├── vba_manager.py      # Gestion modules VBA
│   ├── macro_runner.py     # Exécution macros
│   └── *_optimizer.py      # Optimisation (3 fichiers)
├── tests/                  # Tests unitaires (581 tests)
├── docs/                   # Documentation Sphinx
├── examples/               # Exemples d'utilisation
├── _dev/                   # Documentation de développement
│   ├── architecture.md     # Architecture détaillée (v1.0.0)
│   ├── stories/            # User stories par epic
│   └── reports/            # Rapports d'audit/tests
├── pyproject.toml          # Configuration Poetry + tools
└── README.md               # Ce fichier
```

### Contraintes de développement
- ✅ **[OBL-CHEMINS-001]** : Uniquement `/` ou `pathlib` pour les chemins
- ✅ **[INT-001..004]** : NE JAMAIS modifier `pyproject.toml` pour les dépendances (utiliser `poetry add/remove`)
- ✅ **[EXP-001]** : Modification manuelle OK pour `[tool.ruff]`, `[tool.pytest]`, etc.
- ✅ **Langue** : Code en anglais, CLI/docs en français
- ✅ **License** : GPL-3.0 (entête requis sur tous les fichiers Python)

---

## 🚨 Points importants - À lire absolument

### Gestion COM (Critical)
```python
# ❌ JAMAIS faire cela
app.Quit()  # → Provoque RPC error 0x800706BE!

# ✅ Toujours utiliser context manager
with ExcelManager() as mgr:
    # app.Quit() n'est JAMAIS appelé
    # Libération ordonnée: del ws, del wb, del app, gc.collect()
    pass
```

### Chemins et encodage
```python
# ❌ Incorrect
"examples\\vba_project\\modules"  # Backslash = caractère d'échappement!

# ✅ Correct
Path("examples/vba_project/modules")  # ou "examples/vba_project/modules"
```

### Dépendances
```bash
# ❌ Ne JAMAIS faire
# [Éditer manuellement pyproject.toml pour ajouter une dépendance]

# ✅ Toujours utiliser Poetry
poetry add package_name
poetry add --group dev package_name
```

---

## 📚 Documentation complète

La documentation détaillée est disponible dans :
- **`docs/_build/html/index.html`** : Documentation Sphinx générée
- **`_dev/architecture.md`** : Architecture v1.0.0 (détaillée, 1700+ lignes)
- **`_dev/stories/epic13/`** : User stories par epic (6 epics × 1-6 stories)

---

## 🐛 Signaler un bug

Les bugs peuvent être signalés via :
1. GitHub Issues : https://github.com/MaxImum2509/xlManage/issues
2. Description détaillée avec :
   - Version Python
   - Version Excel
   - Trace complète d'erreur
   - Étapes de reproduction

---

## 🤝 Contribution

Les contributions sont bienvenues! Avant de contribuer, lire :
- `_dev/CLAUDE.md` : Contraintes de développement
- `docs/contributing.rst` : Guide de contribution
- `_dev/architecture.md` : Architecture du projet

**Processus** :
1. Fork le repo
2. Créer une branche `feature/...` ou `fix/...`
3. Commit avec messages clairs
4. Tests et linting (`poetry run pytest`, `poetry run ruff check`)
5. Pull request vers `main`

---

## 📄 License

xlManage est publié sous la **GNU General Public License v3.0**.

Voir `LICENSE` pour les détails complets.

**En bref** : Vous pouvez utiliser, modifier et distribuer ce logiciel librement, mais vous devez :
- Inclure la license
- Publier le code modifié sous GPL-3.0
- Documenter les changements

---

## 🙏 Crédits

**Développement** : Claude (Anthropic)
**Version** : 1.0.0 (2026-02-07)
**Status** : Production-ready

---

## 🔗 Liens utiles

- **Repository** : https://github.com/MaxImum2509/xlManage
- **Documentation** : À venir sur GitHub Pages
- **Issues** : https://github.com/MaxImum2509/xlManage/issues
- **Releases** : https://github.com/MaxImum2509/xlManage/releases

---

## Version History

### v1.0.0 - 2026-02-07
- ✅ Cycle de vie Excel (start/stop/status)
- ✅ CRUD Workbooks, Worksheets, Tables
- ✅ Gestion VBA (import/export/delete modules)
- ✅ Exécution de macros avec arguments
- ✅ Optimisation de performances
- ✅ CLI complète (21 commandes)
- ✅ 581 tests avec 90%+ couverture
- ✅ Documentation Sphinx complète

---

**Faites de l'automatisation Excel simple. Utilisez xlManage.** 🚀
