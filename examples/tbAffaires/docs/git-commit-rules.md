# Règles Messages Commit (Conventional Commits 1.0)

## Objectif

Un historique lisible et exploitable (release notes, tri des changements, debug). Le message explique l’intention (_pourquoi_) et l’impact (_quoi au niveau métier_), pas l’inventaire des fichiers ni le détail du diff.

## Format

### Modèle

```
type[(scope)]: résumé (≤50 caractères)

Corps optionnel (≤72 caractères/ligne).

Footer optionnel :
BREAKING CHANGE: description
Refs: #123
```

### Types (exemples)

- `feat` : nouvelle feature (MINOR).
- `fix` : correction bug (PATCH).
- `docs` : documentation.
- `style` : formatage.
- `refactor` : restructuration.
- `test` : tests.
- `chore` : maintenance.
- `perf` : perf.
- `ci` : CI.
- `build` : build.
- `revert` : revert.

## Règles

### OBLIGATIONS

**OBJ-001** : Respecter le format `type[(scope)]: résumé` (≤50 caractères).

**OBJ-002** : Écrire le verbe du résumé à l'**impératif présent** (forme "action/commande") : `ajouter`, `corriger`, `documenter` ; **pas** au participe passé : `ajouté`, `corrigé`, `documenté`, ni au passé composé : `a corrigé`.

**OBJ-003** : Choisir un `type` qui reflète l'intention (parmi les types listés dans ce document).

**OBJ-004** : Écrire `type` et `scope` en minuscules (ex : `feat(vba)`), sans abréviations ambiguës.

**OBJ-005** : Utiliser un `scope` court et représentatif quand utile (un seul scope principal).

**OBJ-006** : Ajouter un corps uniquement si nécessaire ; wrap à ≤72 caractères/ligne ; expliquer le contexte, le pourquoi, l'impact, et/ou la migration.

**OBJ-007** : Déclarer une rupture avec `!` après le type (et/ou le scope) et/ou un footer `BREAKING CHANGE:` (décrire l'impact + la migration attendue).

**OBJ-008** : Mettre les références traçables en footer (ex : `Refs: #123`).

**OBJ-009** : Si le changement n'est pas évident, détailler dans le corps : motivation, impact métier, contraintes, risques, et (si besoin) consignes de migration.

### INTERDICTIONS

**INT-001** : Interdit d'écrire des sujets vagues : `wip`, `update`, `modifs`, `divers`, `stuff`.

**INT-002** : Interdit de lister l'inventaire du commit dans le message (liste de fichiers/chemins, "x fichiers modifiés", récapitulatif ligne à ligne du diff).

**INT-003** : Interdit de mélanger plusieurs intentions indépendantes dans un seul commit (scinder en commits distincts).

**INT-004** : Interdit d'utiliser des `type`/`scope` non standards ou trompeurs (ex : `hotfix`, `misc`, `changeset`).

## Exemples

**Exemple 1** :

```
feat(vba): ajouter CRUD modules

Implémente import/export .bas/.cls via xlManage.

Refs: #42
```

**Exemple 2** :

```
feat!(api): refactor ExcelManager

BREAKING CHANGE: supprime old API.
```

## Outils

commitlint, semantic-release (auto-changelogs/versions).
