# Structure de répertoires du projet xlManage

Ce document décrit la structure de répertoires du projet xlManage, suivant les conventions Python et les bonnes pratiques de développement.

## Structure principale

```
/
├── src/
│   └── xlmanage/
│       ├── __pycache__/
│       ├── __init__.py
│       └── ... (fichiers source Python)
├── tests/
│   └── ... (fichiers de test)
├── docs/
│   └── ... (documentation)
├── examples/
│   ├── vba_project/
│   │   └── modules/
│   │       └── ... (exemples de modules VBA)
│   └── tbAffaires/
│       ├── data/
│       ├── extractions/
│       └── src/
├── _dev/
│   ├── reports/
│   ├── stories/
│   └── architecture.md
├── .gitignore
├── pyproject.toml
└── README.md
```

## Description des répertoires

### `src/xlmanage/`
Contient le code source principal du projet Python. Le sous-répertoire `__pycache__/` est utilisé pour le cache des fichiers .pyc.

### `tests/`
Contient les tests unitaires et d'intégration pour le projet.

### `docs/`
Contient la documentation du projet.

### `examples/`
Contient des exemples d'utilisation et démonstrations:
- `vba_project/modules/` : Exemples de modules VBA
- `tbAffaires/` : Cas d'utilisation spécifiques avec données et code source

### `_dev/`
Contient la documentation de développement:
- `reports/` : Rapports d'analyse, reviews, tests, etc.
- `stories/` : Stories des epics et des fonctionnalités
- `architecture.md` : Documentation d'architecture globale

## Conventions

- Noms de répertoires en minuscules
- Utilisation de tirets pour les noms composés (ex: `vba-project`)
- Structure suivant les conventions Python (PEP 517 et PEP 518)
- Chemins relatifs pour la portabilité

## Vérification

Tous les répertoires ont été créés avec les permissions appropriées et vérifiés.
