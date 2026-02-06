# Informations sur le Projet `xlManage`

## Objectif du projet

`xlManage` est un utilitaire CLI implémenté en Python qui permet à un agent LLM de contrôler Excel via l'automatisation COM (package `pywin32`). Cette approche offre un contrôle plus complet sur Excel par rapport à `openpyxl`.

## Compétences nécessaires

**Compétences clés à utiliser dans ce projet :**

1. python-poetry-userguide - Gestion Poetry des dépendances
2. python-development-rules - Standards Python PEP 8 et clean code
3. python-com-automation - Automatisation COM générique pywin32
4. project-documentation - Documentation structurée du projet
5. git-best-practices - Bonnes pratiques Git
6. excel-python-tooling - Automatisation Excel spécifique
7. excel-development-rules - Standards développement VBA

## Stack technique

- **License** : GPL v3
- **Version Python** : 3.14
- **Framework de Test** : pytest, pytest-mock, pytest-cov, pytest-timeout
- **Linter** : ruff (linting + formatting)
- **Type Checking** : mypy avec pywin32-stubs
- **Documentation** : Sphinx 9.1.0 avec sphinx-rtd-theme
- **Automatisation Excel** : pywin32 (automatisation COM)
- **CLI** : typer avec rich pour le formatting
- **Gestion des Dépendances** : Poetry pour l'environnement virtuel et la gestion des dépendances
- **Pre-commit** : pre-commit pour les hooks git

## Contraintes de travail

### Construction des chemins de fichiers & répertoires

#### Obligations [OBL-CHEMINS]

⚠️ **OBLIGATION ABSOLUE : [OBL-CHEMINS-001] Utiliser uniquement `/` ou `pathlib`** pour former les chenins

- ✅ `mkdir -p examples/vba_project/modules`
- ✅ `Path("examples/vba_project/modules").mkdir(parents=True)`

**Justification** : Nous voulons être compatiblse cross-platform. Or les backslashes sont des caractères d'échappement sur beaucoup d’OS et de shell. L'utilisation (**INTERDIT**) de `\` crée des répertoires mal nommés, générant des erreurs.

**Application** : Cette règle s'applique à toutes les commandes shell (mkdir, cd, ls, etc.) et aux chaînes de caractères représentant des chemins.

#### Interdictions [INT-CHEMINS]

❌ **NE JAMAIS** utiliser de backslashes (`\`) dans les chemins :

- **[INT-CHEMINS-001]** Dans les commandes shell bash : `mkdir -p examples\vba_project\modules`
- **[INT-CHEMINS-002]** Dans les chaînes de caractères Python : `"examples\\vba_project\\modules"`
- **[INT-CHEMINS-003]** Dans les commandes mixtes shell/Python qui pourraient être interprétées par bash

### Poetry - Gestion des Dépendances (RÈGLES STRICTES)

⚠️ **OBLIGATION ABSOLUE** : Utiliser UNIQUEMENT la CLI Poetry pour gérer les packages et la configuration du projet.

#### Commandes obligatoires [OBL]

- **[OBL-001]** **Ajouter une dépendance** : `poetry add <package>`
- **[OBL-002]** **Ajouter une dépendance dev** : `poetry add --group dev <package>`
- **[OBL-003]** **Supprimer une dépendance** : `poetry remove <package>`
- **[OBL-004]** **Mettre à jour une dépendance** : `poetry update <package>`
- **[OBL-005]** **Synchroniser l'environnement** : `poetry install`

#### INTERDICTIONS ABSOLUES [INT]

❌ **NE JAMAIS** modifier directement le fichier `pyproject.toml` pour :

- **[INT-001]** Ajouter/supprimer des dépendances
- **[INT-002]** Modifier les versions des packages
- **[INT-003]** Changer la configuration Poetry
- **[INT-004]** Éditer les sections `[tool.poetry.dependencies]` ou `[tool.poetry.group.*.dependencies]`

#### Exception [EXP]

✅ La modification manuelle de `pyproject.toml` est UNIQUEMENT autorisée pour :

- **[EXP-001]** Les configurations d'outils (`[tool.ruff]`, `[tool.mypy]`, `[tool.pytest.*]`, etc.)
- **[EXP-002]** Les métadonnées du projet (`[project]`) lors de la création initiale
- **[EXP-003]** Les scripts d'entrée (`[project.scripts]`)

### Formalisation des obligations / interdictions

**Format d'indiçage :** `[TYPE-CATEGORIE-NNN]`

- **TYPE** : `OBL` (Obligation) | `INT` (Interdiction) | `EXP` (Exception)
- **CATEGORIE** : Domaine (`POETRY`, `CHEMINS`, etc.)
- **NNN** : Numéro à 3 chiffres

**Exemples :** `[OBL-POETRY-001]`, `[INT-CHEMINS-002]`

**Structure :**

1. Déclaration détaillée avec exemples ✅/❌ lors de l’énoncé des règles
2. Rappel condensé à la fin des instructionsr avec la référence aux règles

**Application :** Tous les fichiers d'instructions (guidelines Python, VBA, Git, ADR)

## Arborescence du Projet

- **Sources** : Enregistrées dans `src/`
- **Tests** : Créés dans `tests/`
- **Documentations** : Créés dans `docs/`
- **Scripts utilitaires** : Enregistrés dans `scripts/`
- **Développement** : Documentation et rapports dans `_dev/`
    - `_dev/architecture.md` : Documentation d'architecture globale
    - `_dev/stories/` : Stories des epics et des fonctionnalités
    - `_dev/reports/` : Rapports d'analyse, review, tests, etc.
    - `_dev/planning` : pour les plans d’actions ponctuels

## Langues de Communication

- **Langue principale** : Français
- **Documentation technique** : Français (avec termes techniques en anglais lorsque nécessaire)
- **Code source** : Anglais (suivant les conventions de nommage Python)
- **Communication d'équipe** : Français

## Fonctionnalités de `xlManage`

- Contrôle d'Excel via l'automatisation COM :
    - Lancement, arrêt, affichage, masquage et contrôle de propriété de MS Excel
    - CRUD de WorkBooks, WorkSheets, ListObjects
    - CRUD modules VBA standard, des modules de classe et de UserForms (création par import de fichiers .bas, .cls, .frm/frx)
    - Exécution de macros VBA (Sub et Function) avec passage de paramètres

## Tests

- **Framework** : pytest avec pytest-mock pour les mocks
- **Coverage** : pytest-cov pour la mesure de la couverture de code (objectif >90%)
- **Tests parallèles** : pytest-xdist pour exécuter les tests en parallèle
- **Timeout** : pytest-timeout pour éviter les tests bloqués (important pour COM)
- **Configuration** : `tests/conftest.py` pour les fixtures et hooks globaux pytest
- **Commande de test** : `pytest --cov=src/ --cov-report=html --cov-report=term --cov-fail-under=90`

### conftest.py

Le fichier `tests/conftest.py` contient les fixtures et hooks pytest globaux :

- **Fixtures partagées** : Excel app, workbooks, workbooks temporaires, etc.
- **Hooks pytest** : Nettoyage automatique des ressources COM (important pour éviter les processus Excel zombies)
- **Configuration timeout** : Timeout par défaut pour les tests (recommandé : 60s)
- **Configuration markers** : Markers pour catégoriser les tests (ex: @pytest.mark.com pour tests COM)

## Entête License

Tous les fichiers Python doivent inclure l'entête personnalisée suivante en haut du fichier :

```python
"""
[Description brève du module et de son rôle dans xlManage]

This file is part of xlManage.

xlManage is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

xlManage is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with xlManage.  If not, see <https://www.gnu.org/licenses/>.
"""
```

## Rappels des contraintes

### Obligations [OBL]

- Gérer les package par poetry (**[OBL-001]**, **[OBL-002]**, **[OBL-003]**, **[OBL-004]**, **[OBL-005]**)
- Utiliser uniquement `/` ou `pathlib` pour les chemins (**[OBL-CHEMINS-001]**)

### Interdictions [INT]

- ❌ **NE JAMAIS** modifier directement `pyproject.toml` (**[INT-001]**, **[INT-002]**, **[INT-003]**, **[INT-004]**)
- ❌ **NE JAMAIS** utiliser de backslashes (`\`) dans les chemins (**[INT-CHEMINS-001]**, **[INT-CHEMINS-002]**, **[INT-CHEMINS-003]**)

---
