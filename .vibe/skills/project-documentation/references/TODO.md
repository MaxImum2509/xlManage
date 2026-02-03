# TODO.md - Format et Structure

## Emplacement
`_dev/TODO.md`

## But
Tenir à jour le backlog des idées de features, améliorations, bugs à corriger et refactoring à réaliser.

## Structure du fichier

### En-tête
```markdown
# TODO - xlManage

**Dernière mise à jour**: YYYY-MM-DD
```

### Sections

#### 1. Features Prioritaires (High Priority)
Features importantes à implémenter prochainement.

**Format**:
```markdown
## Features Prioritaires

### [Titre de la feature]
- **Priorité** : Haute
- **Description** : Description détaillée de la feature
- **ADR requis** : oui/non (si oui, créer ADR avant implémentation)
- **Estimation** : (optionnel) X jours/heures
- **Dépendances** : (optionnel) liste des dépendances
```

#### 2. Futures Features (Medium Priority)
Features intéressantes mais non urgentes.

**Format**:
```markdown
## Futures Features

### [Titre de la feature]
- **Priorité** : Moyenne
- **Description** : Description de la feature
- **ADR requis** : oui/non
```

#### 3. Améliorations (Enhancements)
Améliorations de fonctionnalités existantes.

**Format**:
```markdown
## Améliorations

### [Titre de l'amélioration]
- **Module affecté** : src/module.py
- **Description** : Description de l'amélioration
- **Avantages** : Liste des bénéfices
```

#### 4. Bugs à Corriger
Liste des bugs identifiés à corriger.

**Format**:
```markdown
## Bugs à Corriger

### [Titre du bug]
- **Module affecté** : src/module.py
- **Description** : Description du bug
- **Sévérité** : Haute/Moyenne/Basse
- **Étapes pour reproduire** :
  1. Étape 1
  2. Étape 2
```

#### 5. Refactoring
Code nécessitant du refactoring.

**Format**:
```markdown
## Refactoring

### [Titre du refactoring]
- **Module affecté** : src/module.py
- **Raison** : Pourquoi le refactoring est nécessaire
- **Approche suggérée** : Description de la solution
```

#### 6. Idées de Recherche
Idées exploratoires, concepts à investiguer.

**Format**:
```markdown
## Idées de Recherche

### [Titre de l'idée]
- **Description** : Description du concept à explorer
- **Questions** : Questions à investiguer
- **Ressources** : (optionnel) Liens vers articles, docs, etc.
```

#### 7. Documentation
Tâches de documentation à compléter.

**Format**:
```markdown
## Documentation

### [Titre de la tâche]
- **Description** : Documentation à créer ou mettre à jour
- **Fichiers concernés** : (optionnel)
```

## Conventions

### Priorités
- **Haute** : Bloquant ou critique pour la roadmap actuelle
- **Moyenne** : Important mais non urgent
- **Basse** : Nice-to-have, peut être reporté

### Tri
Dans chaque section, trier par priorité (haute en haut), puis par ordre d'ajout (récent en haut).

### Statut
Une fois qu'une tâche est commencée, la déplacer dans PROGRESS.md sous "En cours d'implémentation".

Une fois terminée, la déplacer dans PROGRESS.md sous "Features Terminées" ou supprimer si c'est un bug.

### Dates
Mettre à jour "Dernière mise à jour" dans l'en-tête à chaque modification.

## Exemple complet

```markdown
# TODO - xlManage

**Dernière mise à jour**: 2026-02-01

## Features Prioritaires

### Exécution de macros VBA
- **Priorité** : Haute
- **Description** : Permettre d'exécuter des macros VBA avec passage de paramètres
- **ADR requis** : oui
- **Estimation** : 3 jours
- **Dépendances** : Contrôle d'Excel terminé

## Futures Features

### Support des graphiques Excel
- **Priorité** : Moyenne
- **Description** : CRUD des graphiques dans les worksheets
- **ADR requis** : oui

## Améliorations

### Optimisation des accès COM
- **Module affecté** : src/excel_manager.py
- **Description** : Mettre en cache les objets COM fréquemment utilisés
- **Avantages** :
  - Meilleure performance
  - Réduction des appels COM

## Bugs à Corriger

### Mémoire non libérée à la fermeture
- **Module affecté** : src/excel_manager.py
- **Description** : Le processus Excel reste en mémoire après fermeture
- **Sévérité** : Haute
- **Étapes pour reproduire** :
  1. Ouvrir un classeur
  2. Fermer le classeur
  3. Vérifier le gestionnaire de tâches

## Refactoring

### Simplifier la gestion des erreurs COM
- **Module affecté** : src/excel_manager.py
- **Raison** : Code répétitif pour la gestion des exceptions pywin32
- **Approche suggérée** : Créer un décorateur @handle_com_errors

## Idées de Recherche

### Support des UserForms
- **Description** : Explorer la possibilité de manipuler les UserForms VBA
- **Questions** :
  - Est-ce possible via COM ?
  - Quelles sont les limites ?
- **Ressources** : Documentation pywin32, VBA object model

## Documentation

### Tutoriel de démarrage rapide
- **Description** : Créer un guide pour les nouveaux utilisateurs
- **Fichiers concernés** : README.md, docs/tutorial.md
```
