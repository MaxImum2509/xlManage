# Contribuer à xlManage

Merci de votre intérêt pour contribuer à xlManage !

## Introduction

xlManage est un utilitaire CLI en Python pour contrôler Excel via l'automatisation COM. Nous recherchons des contributions pour améliorer les fonctionnalités, corriger des bugs, et améliorer la documentation.

Voir [TODO.md](TODO.md) pour la liste des tâches à accomplir.

## Comment contribuer

### Signaler un bug
Si vous trouvez un bug, ouvrez une issue avec :
- Description détaillée du problème
- Étapes pour reproduire
- Comportement attendu vs réel
- Version de xlManage et de Python
- OS utilisé

### Proposer une nouvelle feature
Avant de proposer une feature :
1. Vérifier que la feature n'existe pas déjà
2. Consulter les ADR existants pour les décisions d'architecture
3. Ouvrir une issue pour discuter de la feature avant de l'implémenter
4. Si la feature est approuvée, un ADR peut être nécessaire

### Soumettre une PR
Pour soumettre une PR :
1. Fork le projet
2. Créer une branche `feature/titre-de-la-feature`
3. Suivre les conventions de code
4. Ajouter des tests (couverture >90%)
5. Passer les tests et les linters
6. Ouvrir une PR avec un descriptif clair

## Configuration de l'environnement de développement

### Prérequis
- Python 3.14+
- pipenv
- Microsoft Excel (pour les tests COM)

### Installation

```bash
# Cloner le dépôt
git clone https://github.com/user/xlmanage.git
cd xlmanage

# Installer les dépendances avec pipenv
pipenv install --dev

# Activer l'environnement virtuel
pipenv shell
```

### Exécuter les tests

```bash
# Exécuter tous les tests
pytest --cov=src/ --cov-report=html --cov-report=term --cov-fail-under=90

# Exécuter en parallèle (plus rapide)
pytest -n auto --cov=src/ --cov-report=html --cov-fail-under=90
```

## Conventions de code

### Style
- Suivre PEP 8
- Utiliser ruff pour le linting et le formatting
- Utiliser mypy pour le type checking
- Tous les fichiers Python doivent inclure l'entête license GPL v3

### Entête license

```python
"""
[Brief description of the module]

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

### Nommage
- Classes : `PascalCase`
- Fonctions/variables : `snake_case`
- Constantes : `UPPER_SNAKE_CASE`
- Modules : `snake_case`

### Tests
- Framework : pytest
- Mocks : pytest-mock
- Coverage objectif : >90%
- Tests parallèles : pytest-xdist
- Timeout : pytest-timeout (60s par défaut)

### Documentation
- Documentation inline : docstrings Google style
- Langue : Français pour les explications, anglais pour le code
- ADR pour les décisions d'architecture

## Processus de Pull Request

### Avant d'ouvrir une PR
1. S'assurer que le code suit les conventions
2. Exécuter les tests : `pytest --cov=src/ --cov-fail-under=90`
3. Exécuter le linter : `ruff check .`
4. Formatter le code : `ruff format .`
5. Exécuter le type checking : `mypy src/`
6. Mettre à jour la documentation si nécessaire

### Contenu de la PR
- Titre clair et descriptif
- Description détaillée des changements
- Liens vers les issues concernées
- Screenshot si applicable (pour les changements UI)
- Cocher les cases de checklist

### Checklist
- [ ] Code suit les conventions
- [ ] Tests ajoutés/mis à jour
- [ ] Coverage >90%
- [ ] Linter passe
- [ ] Type checking passe
- [ ] Documentation mise à jour
- [ ] Changelog mis à jour

### Review process
1. Mainteneurs revoient la PR
2. Comments sont adressés
3. Modifications si nécessaires
4. Approval d'au moins un mainteneur requis
5. Merge après approval

## Documentation

### Types de documentation
- **Code** : Docstrings dans les fichiers Python
- **ADR** : Décisions d'architecture dans `_dev/adr/`
- **PROGRESS** : Avancement dans `_dev/PROGRESS.md`
- **TODO** : Backlog dans `_dev/TODO.md`
- **CHANGELOG** : Historique dans `_dev/CHANGELOG.md`
- **User docs** : Documentation utilisateur dans `docs/`

### Style
- Docstrings Google style
- Français pour les explications
- Exemples de code si pertinent

## Questions et support

### Où poser des questions
- GitHub Issues pour les bugs et features
- GitHub Discussions pour les questions générales
- Email : contact@example.com

### Avant de demander
- Chercher dans les issues existantes
- Lire la documentation
- Consulter les ADR pour comprendre les décisions

## Conventions spécifiques à xlManage

### Code VBA
- Préfixe des procédures publiques avec `Public`
- Commentaires explicites pour les parties complexes
- Gestion d'erreurs avec `On Error`

### COM Automation
- Toujours libérer les objets COM avec `comtypes.CoTaskMemFree()` ou équivalent
- Utiliser des context managers pour la gestion des ressources
- Éviter les appels COM inutiles (mettre en cache)

### Tests COM
- Utiliser `@pytest.mark.com` pour les tests nécessitant Excel
- Nettoyage automatique via `conftest.py`
- Timeout augmenté pour les opérations lentes
