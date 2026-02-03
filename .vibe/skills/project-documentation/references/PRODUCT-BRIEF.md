# Product Brief - Format et Structure

## Emplacement
`_dev/product-brief.md`

## But
Document initial décrivant le problème, la solution proposée et l'opportunité du marché. Ce document est créé avant le PRD et sert de base de discussion.

## Structure du fichier

### En-tête
```markdown
# Product Brief - [NOM_DU_PROJET]

**Date de création** : YYYY-MM-DD
**Version** : 1.0
**Auteur** : [Nom de l'auteur]
```

### Sections

#### 1. Résumé exécutif
Vue d'ensemble du produit.

**Format** :
```markdown
## Résumé exécutif

[Paragraphe décrivant rapidement le produit, son objectif et son public cible]
```

#### 2. Problème
Description du problème à résoudre.

**Format** :
```markdown
## Problème

### Description
[Description détaillée du problème]

### Problème actuel
- [Problème 1]
- [Problème 2]

### Impact
- [Impact business]
- [Impact technique]
- [Impact utilisateur]
```

#### 3. Solution proposée
Description de la solution.

**Format** :
```markdown
## Solution proposée

### Concept
[Description du concept de solution]

### Avantages
- [Avantage 1]
- [Avantage 2]

### Différenciation
[Comment se différencie de solutions existantes]
```

#### 4. Public cible
Qui est le produit pour ?

**Format** :
```markdown
## Public cible

### Utilisateurs principaux
- [Type d'utilisateur 1] : [Description]
- [Type d'utilisateur 2] : [Description]

### Personas
#### Persona 1 : [Nom]
- **Rôle** : [Rôle]
- **Objectifs** : [Liste d'objectifs]
- **Douleurs** : [Liste de problèmes]
- **Motivations** : [Liste de motivations]

#### Persona 2 : [Nom]
[...]
```

#### 5. Cas d'usage
Scénarios d'utilisation.

**Format** :
```markdown
## Cas d'usage

### Cas d'usage 1 : [Titre]
- **Qui** : [Type d'utilisateur]
- **Contexte** : [Contexte]
- **Objectif** : [Objectif]
- **Scénario** :
  1. [Étape 1]
  2. [Étape 2]
  3. [Étape 3]
```

#### 6. Opportunité de marché
Pourquoi maintenant ?

**Format** :
```markdown
## Opportunité de marché

### Taille du marché
- [Estimation de la taille du marché]
- [Segmentation]

### Tendances
- [Tendance 1]
- [Tendance 2]

### Concurrents
| Concurrent | Forces | Faiblesses |
|-----------|--------|------------|
| [Nom] | [...] | [...] |
```

#### 7. Hypothèses clés
Hypothèses à valider.

**Format** :
```markdown
## Hypothèses clés

### Hypothèses de produit
- [Hypothèse 1] : Comment valider
- [Hypothèse 2] : Comment valider
```

#### 8. Succès
Comment mesurer le succès ?

**Format** :
```markdown
## Mesures de succès

### KPIs
- [KPI 1] : [Objectif]
- [KPI 2] : [Objectif]

### Indicateurs de succès
- [Indicateur 1]
- [Indicateur 2]
```

#### 9. Risques et atténuations
Risques potentiels.

**Format** :
```markdown
## Risques et atténuations

| Risque | Probabilité | Impact | Atténuation |
|--------|-------------|--------|-------------|
| [Risque 1] | [Haute/Moyenne/Basse] | [Impact] | [Atténuation] |
```

#### 10. Prochaines étapes
Actions à entreprendre.

**Format** :
```markdown
## Prochaines étapes

### Immédiat
- [ ] [Action 1]
- [ ] [Action 2]
```

## Conventions

### Date
Utiliser le format ISO : `YYYY-MM-DD`

### Mises à jour
Documenter les changements dans une section "Historique des versions" en bas de fichier.

### Liens
- Liens vers les PRD spécifiques
- Liens vers les ADR concernés
- Liens vers les ressources externes
