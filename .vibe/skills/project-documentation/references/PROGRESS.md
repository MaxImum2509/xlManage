# PROGRESS.md - Format et Structure

## Emplacement
`PROGRESS.md` (racine du projet)

## But
Tenir à jour l'avancement du projet avec les épics, stories et tâches.

## Structure du fichier

### En-tête
```markdown
# Avancement du Projet

**Dernière mise à jour**: YYYY-MM-DD
**Version cible**: X.Y.Z
```

### Sections

#### 0. En cours
État global du projet (tâches actives).

**Format**:
```markdown
## En cours

- [ ] Tâche en cours 1
- [ ] Tâche en cours 2
```

#### 1. Terminées
Tâches et stories terminées.

**Format**:
```markdown
## Terminées

- [x] Tâche terminée 1 - YYYY-MM-DD
- [x] Tâche terminée 2 - YYYY-MM-DD
```

#### 2. Epics
Structure hiérarchique pour organiser le développement.

**Structure** :
```markdown
## Epic X - [Nom de l'Epic]

**Objectif** : Description du but et périmètre

### Sections optionnelles de l'Epic
- **Critères d'acceptation** : Conditions de validation
- **Fonctionnalités PRD** : Fn01, Fn02, etc.
- **Dépendances** : Sur d'autres épics ou ressources
- **Risques** : Obstacles potentiels
- **Dates** : Début YYYY-MM-DD / Fin YYYY-MM-DD
- **Bilan** : Résumé après terminaison
- **Notes** : Remarques importantes

### Story Y

- [ ] Tâche 1
- [ ] Tâche 2
- [x] Tâche terminée - YYYY-MM-DD
```

**Exemple complet** :
```markdown
## Epic 1 - Mise en place de l'environnement

**Objectif** : Configurer l'environnement de développement et les tests

**Critères d'acceptation** :
- Tests exécutables avec >90% de couverture
- Linter et type checker configurés

**Dates** : Début 2026-02-01 / Fin 2026-02-05

### Story 1

- [ ] Créer environnement de test (conftest, coverage, mock)
- [ ] Créer le script de démarrage de l'utilitaire CLI
- [ ] Lancer les tests
```

#### 3. Sections optionnelles globales

**Tests et Qualité** :
```markdown
## Tests et Qualité

- **Couverture globale** : XX%
- **Tests unitaires** : XX tests
- **Tests d'intégration** : XX tests
- **Dernier run** : YYYY-MM-DD HH:MM
- **Lint/Format** : ✓ Pass / ✗ Fail
- **Type check** : ✓ Pass / ✗ Fail
```

**Bugs connus** :
```markdown
## Bugs Connus

### [Titre du bug] - Sévérité : [Haute/Moyenne/Basse]
- **Description** : Description du bug
- **Modules affectés** : src/module.py
- **Statut** : [Identifié/En investigation/En correction]
- **Assigné à** : (optionnel)
```

**Prochaine version** :
```markdown
## Prochaine Version : X.Y.Z (Planifiée pour YYYY-MM-DD)

**Objectifs** :
- Feature A
- Correction du bug B

**Tâches** :
- [ ] Implémenter Feature A
- [ ] Corriger bug B
- [ ] Mise à jour documentation
- [ ] Tests de régression
```

## Conventions

### Dates
Toutes les dates au format ISO : `YYYY-MM-DD`

### Numérotation des épics et stories
- Epics : Numérotation séquentielle (Epic 1, Epic 2, etc.)
- Stories : Numérotation séquentielle par epic (Story 1, Story 2, etc.)

### Liens
- Liens vers les ADR pour les décisions techniques
- Liens vers les fichiers source concernés
- Liens vers les tickets d'issues si applicable

### Mise à jour
Mettre à jour la date de dernière mise à jour dans l'en-tête à chaque modification.

## Script update_progress.py

Le script `scripts/update_progress.py` permet de mettre à jour PROGRESS.md avec la structure Epic/Story/Task.

### Commandes disponibles

| Commande | Description | Exemple |
|----------|-------------|---------|
| `list` | Lister tous les epics | `python scripts/update_progress.py list` |
| `add-epic` | Ajouter un nouvel epic | `python scripts/update_progress.py add-epic "Epic X - Nom" --objective "Description"` |
| `add-story` | Ajouter une story à un epic | `python scripts/update_progress.py add-story "Epic X - Nom" "Story Y" --description "Desc"` |
| `add-task` | Ajouter une tâche à une story | `python scripts/update_progress.py add-task "Epic X - Nom" "Story Y" "Tâche"` |
| `complete-task` | Marquer une tâche terminée | `python scripts/update_progress.py complete-task "Epic X - Nom" "Story Y" "pattern"` |
| `completed` | Ajouter à la section Terminées | `python scripts/update_progress.py completed "Description"` |

### Notes d'utilisation

- Le script doit être exécuté depuis la racine du projet
- Le fichier PROGRESS.md doit exister
- Les noms d'epic et de story doivent correspondre exactement aux sections existantes
- `complete-task` utilise un pattern de recherche pour identifier la tâche à marquer

## Exemple complet

```markdown
# Avancement du Projet

**Dernière mise à jour**: 2026-02-01
**Version cible**: 0.1.0

## En cours

- [ ] Implémenter delete_workbook()
- [ ] Ajouter tests de suppression

## Terminées

- [x] Contrôle basique d'Excel - 2026-01-15

## Epic 1 - Mise en place de l'environnement

**Objectif** : Configurer l'environnement de développement et les tests

**Critères d'acceptation** :
- Tests exécutables avec >90% de couverture
- Linter et type checker configurés

**Dates** : Début 2026-02-01 / Fin 2026-02-05

### Story 1

- [x] Créer environnement de test (conftest, coverage, mock) - 2026-02-02
- [ ] Créer le script de démarrage de l'utilitaire CLI
- [ ] Lancer les tests

## Tests et Qualité

- **Couverture globale** : 75%
- **Tests unitaires** : 42 tests
- **Tests d'intégration** : 8 tests
- **Dernier run** : 2026-02-01 14:30
- **Lint/Format** : ✓ Pass
- **Type check** : ✓ Pass
```
