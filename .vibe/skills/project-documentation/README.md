# project-documentation Skill

Skill pour gérer la documentation de projet.

## Emplacement

Fichiers de documentation :

- **À la racine** : `README.md`, `PROGRESS.md`, `TODO.md`, `CHANGELOG.md`, `CONTRIBUTING.md`
- **Dans `_dev/`** : `product-brief.md`, `prd.md`, `architecture.md`, `epics.md`
- **Dans `_dev/adr/`** : Architecture Decision Records

## Utilisation

### Créer les documents initiaux d'un projet

```bash
# Créer le Product Brief
python .opencode/skills/project-documentation/scripts/create_product_brief.py "MonProjet" --author "Moi"

# Créer le PRD
python .opencode/skills/project-documentation/scripts/create_prd.py "MonProjet" --author "Moi"

# Créer l'architecture
python .opencode/skills/project-documentation/scripts/create_architecture.py "MonProjet" --author "Moi"

# Créer les epics
python .opencode/skills/project-documentation/scripts/create_epics.py "MonProjet" --first-version 0.1.0
```

### Créer un ADR

```bash
python .opencode/skills/project-documentation/scripts/create_adr.py "Titre de la décision" --status accepté --decision-makers "Équipe"
```

### Mettre à jour PROGRESS.md

```bash
# Lister les epics
python .opencode/skills/project-documentation/scripts/update_progress.py list

# Ajouter un epic
python .opencode/skills/project-documentation/scripts/update_progress.py add-epic "Epic 1 - Fondations du projet" \
  --objective "Mettre en place la structure du projet et les outils de développement"

# Ajouter une story
python .opencode/skills/project-documentation/scripts/update_progress.py add-story "Epic 1 - Fondations du projet" \
  "Story 1 - Structure du projet" --description "Créer la structure des répertoires et fichiers"

# Ajouter une tâche
python .opencode/skills/project-documentation/scripts/update_progress.py add-task "Epic 1 - Fondations du projet" \
  "Story 1 - Structure du projet" "Créer le module principal"

# Marquer une tâche terminée
python .opencode/skills/project-documentation/scripts/update_progress.py complete-task "Epic 1 - Fondations du projet" \
  "Story 1 - Structure du projet" "module principal"

# Ajouter à la section Terminées
python .opencode/skills/project-documentation/scripts/update_progress.py completed "Feature name - 2026-02-01"
```

### Ajouter dans TODO.md

```bash
# Ajouter une feature
python .opencode/skills/project-documentation/scripts/add_todo.py feature "Features Prioritaires" \
  "Titre" "Description" Haute --adr oui

# Ajouter un bug
python .opencode/skills/project-documentation/scripts/add_todo.py bug "Titre" "module.py" \
  "Description" Haute "Step 1" "Step 2"

# Ajouter une entrée générique
python .opencode/skills/project-documentation/scripts/add_todo.py generic "Améliorations" \
  "Titre" --desc "Description" --module "module.py" --priority Moyenne
```

### Mettre à jour CHANGELOG.md

```bash
# Ajouter une nouvelle version
python .opencode/skills/project-documentation/scripts/update_changelog.py --add-version 0.1.0 --added "Feature 1" "Feature 2"

# Ajouter des changements dans Unreleased
python .opencode/skills/project-documentation/scripts/update_changelog.py --add-unreleased ajouté \
  --items "New feature" "Bug fix"

# Release d'une version
python .opencode/skills/project-documentation/scripts/update_changelog.py --release 0.1.0
```

### Valider la documentation

```bash
python .opencode/skills/project-documentation/scripts/validate_docs.py
```

## Réutilisation sur d'autres projets

Cette skill est conçue pour être réutilisable sur n'importe quel projet.

### Étapes

1. Copier le répertoire `.opencode/skills/project-documentation/` dans votre projet
2. Exécuter les scripts de création de documentation
3. Adapter les fichiers de documentation à votre projet
4. Utiliser les scripts pour mettre à jour la documentation

### Adapters les scripts

Les scripts sont génériques et utilisent des templates avec des marqueurs comme `[NOM_DU_PROJET]`. Remplacez ces marqueurs par votre nom de projet.

### Chemins par défaut

- Racine du projet : fichiers créés dans le répertoire courant
- `_dev/` : fichiers de documentation interne
- `_dev/adr/` : Architecture Decision Records

## Formatage

### Conventions de nommage

- **ADR** : `ADR-XXX-titre-aaaammjj.md` (XXX à 3 chiffres, date au format aaaammjj)
- **Epic** : `Epic X : Nom` (numérotation séquentielle)
- **Story** : `Story X.Y : Nom` (X = numéro d'epic, Y = numéro de story)

### Conventions de priorité

- **Must have** : Essentiel pour la version
- **Should have** : Important mais non bloquant
- **Could have** : Souhaité si possible
- **Won't have** : Explicitement exclu

### Conventions de complexité

- **Petite** : 0.5 - 1 jour
- **Moyenne** : 1 - 3 jours
- **Grande** : 3+ jours
