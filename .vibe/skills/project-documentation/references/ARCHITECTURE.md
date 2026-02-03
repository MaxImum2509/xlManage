# Architecture - Format et Structure

## Emplacement
`_dev/architecture.md`

## But
Document décrivant l'architecture technique du projet, les composants, les patterns de design, et les décisions techniques.

## Structure du fichier

### En-tête
```markdown
# Architecture - [NOM_DU_PROJET]

**Version** : 1.0
**Date de création** : YYYY-MM-DD
**Dernière mise à jour** : YYYY-MM-DD
**Auteur** : [Nom]
```

### Sections

#### 1. Vue d'ensemble
Présentation de l'architecture.

**Format** :
```markdown
## Vue d'ensemble

[NOM_DU_PROJET] est [type d'application] avec une architecture basée sur
[type d'architecture].

### Objectifs architecturaux
- [Objectif 1]
- [Objectif 2]
```

#### 2. Architecture de haut niveau
Vue macroscopique.

**Format** :
```markdown
## Architecture de haut niveau

### Couches

```
┌─────────────────────────────────────┐
│         Couche Application          │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│      Couche [NOM_DU_PROJET]       │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│      Couche d'abstraction           │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│      [Couche d'accès externe]      │
└──────────────────────────────────────┘
```
```

#### 3. Composants
Description détaillée des composants.

**Format** :
```markdown
## Composants

### [Composant 1]
**Responsabilité** : [Description]

**Attributs** :
- `attr1` : [Description]

**Méthodes principales** :
```python
def method1(self) -> ReturnType:
    """Description."""
    ...
```

**Responsabilités** :
- [Responsabilité 1]
```

#### 4. Patterns de design
Patterns utilisés.

**Format** :
```markdown
## Patterns de design

### [Pattern 1]
**Où** : [Composant]

**Pourquoi** : [Raison]

**Exemple** :
```python
# Exemple d'utilisation
...
```
```

#### 5. Gestion des erreurs
Stratégie de gestion d'erreurs.

**Format** :
```markdown
## Gestion des erreurs

### Exceptions personnalisées
```python
class BaseError(Exception):
    """Exception de base."""
    pass
```

### Stratégie de propagation
- [Description de la stratégie]
```

#### 6. Dépendances
Liste des dépendances externes.

**Format** :
```markdown
## Dépendances

### Dépendances principales
- **[Dépendance 1]** : [Description]
  - Version : [Version]
  - Raison : [Raison]
```

#### 7. Décisions techniques (ADR)
Liens vers les ADR.

**Format** :
```markdown
## Décisions techniques

### ADR-XXX : [Titre]
- **Status** : [Accepté]
- **Voir** : [ADR-XXX-titre-aaaammjj.md](adr/ADR-XXX-titre-aaaammjj.md)
```

## Conventions

### Diagrammes
Utiliser ASCII ou Mermaid pour les diagrammes.

### Liens
- Liens vers les ADR pour les décisions
- Liens vers le code source (si applicable)
