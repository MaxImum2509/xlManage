---
name: project-documentation
description: Gérer la documentation de projet : README (racine), PROGRESS (racine), TODO (racine), ADR (_dev/adr/), CHANGELOG (racine), CONTRIBUTING (racine), ainsi que Product Brief (_dev/), PRD (_dev/), Architecture (_dev/), et Epics (_dev/).
---

# Project Documentation Management

Méthodes pour utiliser, créer, mettre à jour ou valider la documentation du projet.

## Quick Reference

| Fichier                     | Usage                                       | Emplacement      | Référence                                                  |
| --------------------------- | ------------------------------------------- | ---------------- | ---------------------------------------------------------- |
| `README.md`                 | Présentation du projet                      | Racine du projet | [references/README.md](references/README.md)               |
| `PROGRESS.md`               | Avancement du projet                        | Racine du projet | [references/PROGRESS.md](references/PROGRESS.md)           |
| `TODO.md`                   | Idées et améliorations futures              | Racine du projet | [references/TODO.md](references/TODO.md)                 |
| `CHANGELOG.md`              | Historique des changements par version      | Racine du projet | [references/CHANGELOG.md](references/CHANGELOG.md)         |
| `CONTRIBUTING.md`           | Guide pour les contributeurs                | Racine du projet | [references/CONTRIBUTING.md](references/CONTRIBUTING.md)   |
| `product-brief.md`          | Vue d'ensemble et opportunité du produit    | `_dev/`          | [references/PRODUCT-BRIEF.md](references/PRODUCT-BRIEF.md) |
| `prd.md`                    | Spécification détaillée des fonctionnalités | `_dev/`          | [references/PRD.md](references/PRD.md)                     |
| `architecture.md`           | Architecture technique et composants        | `_dev/`          | [references/ARCHITECTURE.md](references/ARCHITECTURE.md)   |
| `epics.md`                  | Epics et user stories                       | `_dev/`          | [references/EPICS.md](references/EPICS.md)                 |
| `ADR-XXX-titre-aaaammjj.md` | Décisions d'architecture                    | `_dev/adr/`      | [references/ADR.md](references/ADR.md)                     |

## Core Principles

### Emplacement

- **À la racine du projet** : `README.md`, `PROGRESS.md`, `TODO.md`, `CHANGELOG.md`, `CONTRIBUTING.md`
- **Dans `_dev/`** : `product-brief.md`, `prd.md`, `architecture.md`, `epics.md`
- **Dans `_dev/adr/`** : Les fichiers ADR

### Format ADR

Les ADR utilisent le format de nommage : `ADR[NNN]-[Objet]-[aaaammjj].md` avec NNN à 3 chiffres (ex: `ADR001-choix-du-framework-20260201.md`).

### Ordre de création

1. **Product Brief** : Avant tout, pour valider l'opportunité
2. **PRD** : Pour spécifier les fonctionnalités
3. **Architecture** : Pour définir la structure technique
4. **Epics** : Pour détailler les stories par version
5. **ADR** : Au fur et à mesure des décisions techniques
6. **PROGRESS** : Mise à jour continue lors du développement
7. **TODO** : Backlog des idées et améliorations
8. **CHANGELOG** : Historique par version
9. **README** : Présentation publique (après MVP)

### Synchronisation

La documentation doit être tenue à jour en parallèle du développement :

- Chaque décision importante mérite un ADR
- Chaque feature terminée met à jour PROGRESS.md
- Chaque nouvelle idée va dans TODO.md
- Chaque release met à jour CHANGELOG.md

## Workflows

### Créer un nouveau projet

1. Créer `product-brief.md` avec `scripts/create_product_brief.py`
2. Créer `prd.md` avec `scripts/create_prd.py`
3. Créer `architecture.md` avec `scripts/create_architecture.py`
4. Créer `epics.md` avec `scripts/create_epics.py`
5. Initialiser `PROGRESS.md` avec `scripts/update_progress.py`

### Mettre à jour PROGRESS.md

Utiliser le script `scripts/update_progress.py` :

```bash
# Lister les epics
python scripts/update_progress.py list

# Ajouter un epic
python scripts/update_progress.py add-epic "Epic X - Nom" --objective "Description"

# Ajouter une story
python scripts/update_progress.py add-story "Epic X - Nom" "Story Y" --description "Desc"

# Ajouter une tâche
python scripts/update_progress.py add-task "Epic X - Nom" "Story Y" "Tâche"

# Marquer une tâche terminée
python scripts/update_progress.py complete-task "Epic X - Nom" "Story Y" "pattern"

# Ajouter à la section Terminées
python scripts/update_progress.py completed "Description"
```

Voir [references/PROGRESS.md](references/PROGRESS.md) pour le format détaillé.

### Créer un ADR

```bash
python scripts/create_adr.py "Titre de la décision"
```

Voir [references/ADR.md](references/ADR.md) pour le format complet.

### Ajouter dans TODO.md

```bash
python scripts/add_todo.py feature "Features Prioritaires" "Titre" "Description" Haute
```

Voir [references/TODO.md](references/TODO.md) pour les catégories et priorités.

### Mettre à jour CHANGELOG.md

```bash
# Ajouter une nouvelle version
python scripts/update_changelog.py --add-version 0.1.0 --added "Feature 1" "Feature 2"

# Release d'une version
python scripts/update_changelog.py --release 0.1.0
```

Voir [references/CHANGELOG.md](references/CHANGELOG.md) pour le format Keep a Changelog.

### Valider la documentation

```bash
python scripts/validate_docs.py
```

## Références

### Documentation produit
- **Product Brief** : [references/PRODUCT-BRIEF.md](references/PRODUCT-BRIEF.md)
- **PRD** : [references/PRD.md](references/PRD.md)
- **Architecture** : [references/ARCHITECTURE.md](references/ARCHITECTURE.md)
- **Epics** : [references/EPICS.md](references/EPICS.md)

### Documentation projet
- **PROGRESS.md** : [references/PROGRESS.md](references/PROGRESS.md)
- **TODO.md** : [references/TODO.md](references/TODO.md)
- **CHANGELOG.md** : [references/CHANGELOG.md](references/CHANGELOG.md)
- **CONTRIBUTING.md** : [references/CONTRIBUTING.md](references/CONTRIBUTING.md)
- **README.md** : [references/README.md](references/README.md)

### Documentation technique
- **ADR** : [references/ADR.md](references/ADR.md)
