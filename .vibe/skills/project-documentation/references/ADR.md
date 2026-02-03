# ADR - Architecture Decision Records

## Emplacement
`_dev/adr/ADR[NNN]-[Objet]-[aaaammjj].md`

## But
Documenter les décisions d'architecture importantes prises pendant le développement de xlManage. Chaque ADR capture le contexte, la décision, et les alternatives considérées.

## Format de nommage

Les fichiers ADR suivent ce format : `ADR[NNN]-[Objet]-[aaaammjj].md`

- **ADR** : Préfixe fixe
- **NNN** : Numéro séquentiel à 3 chiffres (001, 002, etc.)
- **Objet** : Titre court descriptif en kebab-case
- **aaaammjj** : Date de la décision en format ISO (ex: 20260201)

**Exemples** :
- `ADR001-choix-du-framework-20260201.md`
- `ADR002-strategie-de-gestion-erreurs-20260215.md`

## Structure du fichier ADR

### Template standard

```markdown
---
title: ADR[NNN]: [Titre de la décision]
status: [proposé/accepté/rejeté/obsolète/remplacé]
date: [aaaammjj]
decision-makers: [Liste des personnes ayant pris la décision]
technical-stories: [Tous tickets/liens concernés]
---

# Contexte et description du problème

[Description du problème à résoudre, du contexte et de la motivation pour cette décision.]

# Les options alternatives

## Option 1: [Titre de l'option]
[Description de l'option]
- **Avantages** : Liste des avantages
- **Désavantages** : Liste des désavantages

## Option 2: [Titre de l'option]
[Description de l'option]
- **Avantages** : Liste des avantages
- **Désavantages** : Liste des désavantages

# La décision

[Description de la décision choisie et sa justification.]

# Status des conséquences

[Description des conséquences positives et négatives de cette décision, et ce qui doit être fait ensuite.]

## Avantages
- [Liste des avantages]

## Désavantages
- [Liste des désavantages]

## ADRs concernés
- [ADR001-titre-aaaammjj.md](ADR001-titre-aaaammjj.md)
```

## Status possibles

| Status      | Description                            |
| ----------- | -------------------------------------- |
| proposé     | Décision proposée mais pas encore prise |
| accepté     | Décision acceptée et implémentée       |
| rejeté      | Décision rejetée                       |
| obsolète    | Décision obsolète mais pas remplacée   |
| remplacé    | Décision remplacée par un nouvel ADR   |

## Conventions

### Numérotation
Les numéros sont séquentiels et ne sont jamais réutilisés. Commencer à 001.

### Langue
Les ADR sont rédigés en **français** pour la documentation, mais le code et les termes techniques restent en anglais.

### Liens
- Lier les ADR entre eux lorsque pertinent (références, remplacements)
- Lier vers le code implémentant la décision si possible

### Mise à jour du status
- Quand une décision est rejetée, ne pas supprimer l'ADR : changer le status en "rejeté"
- Quand une décision est remplacée, créer un nouvel ADR et marquer l'ancien comme "remplacé" avec un lien vers le nouveau

## Quand créer un ADR ?

Créer un ADR pour :
- Choix de framework ou bibliothèque majeure
- Décisions d'architecture de haut niveau
- Changements importants dans la structure du code
- Choix de patterns de design
- Décisions sur la gestion des erreurs
- Choix de formats de données ou API
- Décisions impactant la performance ou la sécurité

Ne pas créer d'ADR pour :
- Implémentations mineures
- Choix triviaux (ex: nom d'une variable)
- Corrections de bugs
- Refactoring localisé

## Exemple complet

```markdown
---
title: ADR001: Choix de pywin32 pour l'automatisation Excel
status: accepté
date: 20260110
decision-makers: dev team
technical-stories: FEAT-001
---

# Contexte et description du problème

xlManage doit contrôler Microsoft Excel via Python. Plusieurs options existent pour l'automatisation Excel : pywin32, openpyxl, xlwings, etc. Nous avons besoin d'une solution qui permet un contrôle complet d'Excel, incluant l'accès au modèle objet VBA, l'exécution de macros, et la manipulation de UserForms.

# Les options alternatives

## Option 1: pywin32
Interface COM directe vers Excel. Permet un contrôle complet et l'accès à toutes les fonctionnalités d'Excel.

- **Avantages** :
  - Contrôle complet d'Excel
  - Accès au modèle objet VBA
  - Exécution de macros
  - Manipulation de modules VBA
  - Performance optimale (accès COM direct)
- **Désavantages** :
  - Plus complexe à utiliser
  - Windows uniquement
  - Gestion manuelle des ressources COM

## Option 2: openpyxl
Bibliothèque Python pure pour lire/écrire des fichiers Excel.

- **Avantages** :
  - Multi-plateforme
  - API simple
  - Pas besoin d'Excel installé
- **Désavantages** :
  - Pas de contrôle d'Excel (lecture/écriture de fichiers uniquement)
  - Pas d'accès aux macros
  - Pas d'exécution de VBA

## Option 3: xlwings
Wrapper autour de pywin32 avec une API plus simple.

- **Avantages** :
  - API plus simple que pywin32
  - Multi-plateforme (via AppleScript sur Mac)
  - Good documentation
- **Désavantages** :
  - Couche d'abstraction supplémentaire
  - Moins flexible que pywin32 direct
  - Dépendance supplémentaire

# La décision

Nous choisissons **pywin32** pour l'automatisation Excel.

**Justification** :
- Contrôle complet d'Excel est nécessaire (ouverture/fermeture, affichage/masquage)
- Accès au modèle objet VBA est requis (création de modules, UserForms)
- Exécution de macros VBA avec paramètres est indispensable
- Performance optimale pour les opérations COM intensives

La complexité de pywin32 sera mitigée par :
- Création de wrappers/helpers dans `ExcelManager`
- Documentation exhaustive de l'API COM Excel
- Tests complets de toutes les opérations

# Status des conséquences

## Avantages
- Contrôle complet d'Excel possible
- Accès à toutes les fonctionnalités VBA
- Performance optimale
- Pas de dépendance supplémentaire inutile

## Désavantages
- Windows uniquement
- Plus complexe à implémenter
- Gestion manuelle nécessaire des ressources COM (risque de processus Excel zombies)

## ADRs concernés
- Aucun (premier ADR)

## À faire
- [ ] Implémenter la classe `ExcelManager` avec gestion des ressources COM
- [ ] Créer des tests pour vérifier le nettoyage des processus Excel
- [ ] Documenter l'API COM Excel utilisée
```
