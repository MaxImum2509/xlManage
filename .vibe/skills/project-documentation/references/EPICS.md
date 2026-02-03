# Epics - Format et Structure

## Emplacement
`_dev/epics.md`

## But
Document détaillant les epics du projet. Chaque epic regroupe des fonctionnalités liées avec leurs stories, estimations et dépendances.

## Structure du fichier

### En-tête
```markdown
# Epics - [NOM_DU_PROJET]

**Version** : 1.0
**Date de création** : YYYY-MM-DD
**Dernière mise à jour** : YYYY-MM-DD
```

### Sections

#### 1. Vue d'ensemble
Présentation des epics.

**Format** :
```markdown
## Vue d'ensemble

Ce document décrit les epics planifiés pour [NOM_DU_PROJET].

### Liste des epics
1. [Epic 1] - [Version]
2. [Epic 2] - [Version]
```

#### 2. Epic template
Template pour chaque epic.

**Format** :
```markdown
## Epic 1 : [Nom de l'epic]

**Version cible** : X.Y.Z
**Priorité** : [Haute/Moyenne/Basse]
**Statut** : [Non démarré/En cours/Terminé]
**Propriétaire** : [Nom]
**ADR associé** : [ADR-XXX-titre-aaaammjj.md](adr/ADR-XXX-titre-aaaammjj.md)

### Description
[Description détaillée de l'epic]

### Objectifs
- [Objectif 1]

### Stories

#### Story 1.1 : [Titre de la story]
- **User Story** : En tant que [persona], je veux [action], afin que [bénéfice]
- **Priorité** : [Must have/Should have/Could have]
- **Complexité** : [Petite/Moyenne/Grande]
- **Estimation** : X jours/heures
- **Critères d'acceptation** :
  - [ ] [Critère 1]
- **Tests** :
  - [ ] Test unitaire 1
- **Statut** : [Non démarré/En cours/Terminé]
```

## Conventions

### Statut
- **Non démarré** : Epic/story planifié mais pas commencé
- **En cours** : En cours de développement
- **Terminé** : Implémenté et testé

### Priorité
- **Must have** : Essentiel pour la version
- **Should have** : Important mais non bloquant
- **Could have** : Souhaité si possible

### Complexité
- **Petite** : 0.5 - 1 jour
- **Moyenne** : 1 - 3 jours
- **Grande** : 3+ jours

### User Stories
Format : "En tant que [persona], je veux [action], afin que [bénéfice]"

### Liens
- Liens vers les ADR pour les décisions techniques
- Liens vers le PRD pour plus de détails
- Liens vers PROGRESS.md pour l'avancement
