# CHANGELOG.md - Format et Structure

## Emplacement
`_dev/CHANGELOG.md`

## But
Documenter les changements entre les différentes versions de xlManage en suivant le format [Keep a Changelog](https://keepachangelog.com/).

## Structure du fichier

### En-tête
```markdown
# Changelog

Tous les changements notables de ce projet sont documentés dans ce fichier.

Le format est basé sur [Keep a Changelog](https://keepachangelog.com/),
et ce projet adhère au [Versioning Sémantique](https://semver.org/lang/fr/).

## [X.Y.Z] - YYYY-MM-DD

### Ajouté
- Nouvelle feature 1
- Nouvelle feature 2

### Changé
- Modification 1

### Déprécié
- Fonctionnalité dépréciée 1

### Supprimé
- Fonctionnalité supprimée 1

### Corrigé
- Bug 1

### Sécurité
- Correction de sécurité 1
```

## Sections

### Ajouté (Added)
Nouvelles fonctionnalités, méthodes, API, etc.

### Changé (Changed)
Modifications dans la fonctionnalité existante.

### Déprécié (Deprecated)
Fonctionnalités qui seront supprimées dans les prochaines versions.

### Supprimé (Removed)
Fonctionnalités supprimées de cette version.

### Corrigé (Fixed)
Corrections de bugs.

### Sécurité (Security)
Corrections liées à la sécurité.

## Format de version

Les versions suivent le format **Versioning Sémantique** : `MAJEUR.MINEUR.PATCH`

- **MAJEUR** : Changements incompatibles avec l'API
- **MINEUR** : Nouvelles fonctionnalités compatibles en arrière
- **PATCH** : Corrections de bugs compatibles en arrière

## Conventions

### Ordre
1. La version la plus récente en premier
2. Dans chaque version : Ajouté → Changé → Déprécié → Supprimé → Corrigé → Sécurité

### Format des entrées
- Commencer par un tiret et un espace : `- `
- Utiliser le présent : "Ajouté" au lieu de "Ajouter"
- Être concis mais informatif
- Inclure des liens vers :
  - Les ADR concernés
  - Les issues/tickets
  - Les PR (Pull Requests)

### Dates
Date de sortie au format ISO : `YYYY-MM-DD`

### Liens
Utiliser la syntaxe Markdown pour les liens :
- `[ADR001](adr/ADR001-titre-20260201.md)`
- `[#42](https://github.com/user/repo/issues/42)`
- `[PR#12](https://github.com/user/repo/pull/12)`

### Changelog par feature
Pour les gros changements ou features majeures, créer un fichier de changelog détaillé dans `_dev/changelog/` avec le nom `feat-[feature-name]-[version].md`.

**Format** :
```markdown
# Changelog détaillé : [Feature Name] v[X.Y.Z]

## Contexte
[Description du contexte]

## Changements
[Changements détaillés]

## Migration guide
[Guide pour migrer depuis la version précédente]

## Breaking changes
[Liste des changements incompatibles]
```

Lier ce fichier depuis le CHANGELOG.md :
```markdown
### Ajouté
- [Nouvelle feature](changelog/feat-feature-name-0.1.0.md)
```

## Exemple complet

```markdown
# Changelog

Tous les changements notables de ce projet sont documentés dans ce fichier.

Le format est basé sur [Keep a Changelog](https://keepachangelog.com/),
et ce projet adhère au [Versioning Sémantique](https://semver.org/lang/fr/).

## [0.2.0] - 2026-02-15

### Ajouté
- Support des macros VBA ([ADR003](adr/ADR003-macros-vba-20260210.md))
- CRUD des modules de classe VBA
- Exécution de fonctions VBA avec retour de valeur

### Changé
- `ExcelManager.run_macro()` peut maintenant retourner des valeurs
- Amélioration des performances des accès COM ([ADR004](adr/ADR004-optimisation-com-20260212.md))

### Déprécié
- `ExcelManager.execute_vba()` sera remplacé par `run_macro()` dans v0.3.0

### Corrigé
- Fix : Processus Excel non libéré lors d'une exception COM ([#42](https://github.com/user/repo/issues/42))
- Fix : Crash lors de l'import de UserForms avec espaces dans le nom

## [0.1.0] - 2026-01-15

### Ajouté
- Contrôle basique d'Excel (lancement, arrêt, affichage, masquage)
- CRUD de WorkBooks et WorkSheets
- CRUD de ListObjects (tableaux Excel)
- CRUD de modules VBA standard (import de fichiers .bas)
- Support des UserForms (import de fichiers .frm/.frx)
- Exécution de macros VBA (Sub uniquement)
- Gestion automatique des ressources COM

### Changé
- Renommage de `ExcelApp` en `ExcelManager`
- Restructuration du code en modules

## [0.0.1] - 2025-12-01

### Ajouté
- Version initiale
- Structure de base du projet
- Configuration pytest et ruff
```

## Processus de mise à jour

### Avant chaque release
1. Créer une nouvelle section de version
2. Déplacer les entrées de "Unreleased" vers les sections appropriées
3. Mettre à jour la date de sortie

### Pendant le développement
1. Ajouter les changements dans une section `[Unreleased]`
2. Lier vers les ADR créés pour les décisions importantes

### Après chaque release
1. Tagger le commit de release avec `vX.Y.Z`
2. Créer un release GitHub avec les notes du changelog
3. Mettre à jour `PROGRESS.md` avec la version terminée
