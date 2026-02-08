# Règles Python

## Chemins

### OBLIGATIONS

- **OBL-001** : Utiliser le séparateur de chemins `/` ou `pathlib` uniquement.

### INTERDICTIONS

- **INT-001**: Jamais backslash (`\\`) en comme séparateur de chemin.

## Management du projet (utilitaire poetry)

### OBLIGATIONS

- **OBL-002** : Utiliser en PRIORITÉ `poetry`pour modifier le fichier `pyproject.toml` (initialisation du projet, installation de paquet, modification des propriétés du projet, etc)
- **OBL-003** : Appel de scripts et utilitaires par `poetry run`
- **OBL-004** : Respect STRICTE des PEP 518, PEP 621, PEP 668

### INTERDICTIONS

- **INT-002** : Ne JAMAIS modifier `pyproject.toml` directement s’il existe une fonction de `poetry` pour cela.
- **INT-003** : Ne JAMAIS utiliser `pip` pour l’installation de paquets

## Standards Code

### OBLIGATIONS

- **OBL-005** : Encode les fichiers en UTF-8 avec UNIQUEMENT LF (`\n`) pour passer à la ligne
- **OBL-006** : En-tête licence GPL v3
- **OBL-007** : Code en anglais (variables/fonctions/classes/docstrings/commits).
- **OBL-008** : Arborescence: src/ (code), tests/, docs/, scripts/ (tooling pour le développement)
- **OBL-009** : Tooling Python 3.14, git, Sphynx 9.1.x, pytest avec mock, ruff, mypy, bandit, pre-commit
- **OBL-010** : Suivre le Zen of Python (PEP 20)
- **OBL-011** : Respect STRICTE des règles de formatage PEP 8.
- **OBL-012** : Applique les règles du "clean code" (SRP (<20-30 lignes/fonction), DRY, KISS, YAGNI, SOLID, exceptions spécifiques)
- **OBL-013** : Documenter l’intégralité du projet en anglais avec les conventions de docstrings (PEP 257, PEP 287)
- **OBL-014** : Employer des "type hints" (PEP 484, PEP 585, PEP 586, PEP 589, PEP 604, PEP 612, PEP 544, PEP 591, PEP 613, PEP 695, PEP 561)
- **OBL-015** : Applique la gestion des paquets namespace implicites (PEP 420)
- **OBL-016** : Tests pytest avec mock pour coverage ≥90%.
- **OBL-017** : Valider la qualité du code à l’aide de ruff, mypy et bandit
- **OBL-018** : Utiliser pre-commit avant chaque commit
- **OBL-019** : Commiter en respectant les règles en vigueur

### INTERDICTIONS

- **INT-004**: NE JAMAIS utiliser d’Emojis, SAUF dans les fichiers markdown
- **INT-005**: AUCUN fonctions/classe/modules sans docstring ; `except:` générique.
