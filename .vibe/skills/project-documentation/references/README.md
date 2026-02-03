# README.md - Format et Structure

## Emplacement
`README.md` à la racine du projet

## But
Présentation générale du projet pour les nouveaux utilisateurs et contributeurs.

## Structure du fichier

### En-tête
```markdown
# [NOM_DU_PROJET]

[Badge License]
[Badge Version Python]
[Badge Status Build]
[Badge Coverage]

**[Description courte du projet en une phrase]**
```

### Sections recommandées

#### 1. Introduction
Présentation rapide du projet.

**Format** :
```markdown
## Qu'est-ce que [NOM_DU_PROJET] ?

[NOM_DU_PROJET] est [type de projet] qui [principe principal].
Il permet [principales fonctionnalités].

### Caractéristiques principales

- ✅ [Fonctionnalité 1]
- ✅ [Fonctionnalité 2]
- ✅ [Fonctionnalité 3]
```

#### 2. Installation
Instructions d'installation.

**Format** :
```markdown
## Installation

### Prérequis

- [Prérequis 1]
- [Prérequis 2]
- [Prérequis 3]

### Installation

```bash
git clone https://github.com/user/[projet].git
cd [projet]
[commandes d'installation]
```
```

#### 3. Usage rapide
Exemples de base.

**Format** :
```markdown
## Usage rapide

### [Exemple d'usage 1]

```python
[exemple de code]
```

### [Exemple d'usage 2]

```python
[exemple de code]
```
```

#### 4. Documentation
Liens vers la documentation détaillée.

**Format** :
```markdown
## Documentation

- [Documentation utilisateur](docs/index.md)
- [API Reference](docs/api.md)
- [Architecture](_dev/architecture.md)
- [Product Brief](_dev/product-brief.md)
- [PRD](_dev/prd.md)
```

#### 5. Développement
Informations pour les contributeurs.

**Format** :
```markdown
## Développement

Voir [CONTRIBUTING.md](CONTRIBUTING.md) pour plus d'informations sur la contribution.

### Exécuter les tests

```bash
[commande pour exécuter les tests]
```

### Linter et formatting

```bash
[commandes pour le linting et le formatting]
```
```

#### 6. Roadmap
Fonctionnalités à venir.

**Format** :
```markdown
## Roadmap

Voir [TODO.md](TODO.md) pour la liste complète des fonctionnalités planifiées.

### Prochaine version (X.Y.Z)

- [ ] [Fonctionnalité 1]
- [ ] [Fonctionnalité 2]
```

#### 7. License
Informations de license.

**Format** :
```markdown
## License

Ce projet est sous license [TYPE_DE_LICENSE]. Voir [LICENSE](LICENSE) pour plus de détails.
```

## Conventions

### Badges
Utiliser des badges standards : License, Version Python, Status Build, Coverage.

### Code snippets
Utiliser des blocs de code avec spécification du langage.

### Liens
- Liens relatifs vers les fichiers dans _dev/
- Liens absolus vers la documentation externe

### Images
Captures d'écran dans `docs/images/` avec des noms descriptifs.

## Exemple complet

```markdown
# [NOM_DU_PROJET]

[![License: TYPE](https://img.shields.io/badge/License-TYPE-blue.svg)](LICENSE)
[![Python 3.14+](https://img.shields.io/badge/python-3.14+-blue.svg)](https://www.python.org/)

**[Description courte du projet]**

## Qu'est-ce que [NOM_DU_PROJET] ?

[NOM_DU_PROJET] est [type de projet] qui [principe principal].
Il permet [principales fonctionnalités].

### Caractéristiques principales

- ✅ [Fonctionnalité 1]
- ✅ [Fonctionnalité 2]
- ✅ [Fonctionnalité 3]

## Installation

### Prérequis

- [Prérequis 1]
- [Prérequis 2]

### Installation

```bash
git clone https://github.com/user/[projet].git
cd [projet]
[commandes d'installation]
```

## Usage rapide

### [Exemple d'usage 1]

```python
[exemple de code]
```

## Documentation

- [Documentation utilisateur](docs/index.md)
- [Architecture](_dev/architecture.md)
- [PRD](_dev/prd.md)

## Développement

Voir [CONTRIBUTING.md](CONTRIBUTING.md) pour plus d'informations sur la contribution.

## License

Ce projet est sous license [TYPE_DE_LICENSE].
```
