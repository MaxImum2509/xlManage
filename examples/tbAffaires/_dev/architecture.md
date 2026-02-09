# Architecture Technique - tbAffaires

> Ce document décrit **uniquement l'architecture technique** du projet.
> Pour les règles de codage : `docs/excel-development-rules.md` et `docs/python-development-rules.md`.
> Pour le processus de développement : `docs/excel-development-process.md`.
> Pour les règles d'implémentation : `project-context.md` (racine du projet).

## Glossaire

| Terme          | Signification                                                                                              |
| -------------- | ---------------------------------------------------------------------------------------------------------- |
| **ADV**        | Assistante De Vente (utilisatrice de l'application, 3 au total)                                            |
| **ERP**        | Progiciel de gestion intégré (source des données d'affaires)                                               |
| **RAII**       | Resource Acquisition Is Initialization : un objet acquiert des ressources à sa création et les libère automatiquement à sa destruction, même en cas d'erreur |
| **UPSERT**     | Update + Insert : met à jour les lignes existantes et insère les nouvelles en une seule opération           |
| **ListObject** | Tableau structuré Excel (créé via Insertion > Tableau). Permet de manipuler les données par colonnes nommées au lieu d'adresses de cellules |
| **Trigramme**  | Code à 3 lettres identifiant un ADV (ex : PAT, SOC, MAR)                                                  |
| **Fichier consolidé** | Fichier Excel de suivi hebdomadaire (ex : `Suivi affaires 2026-S05.xlsx`). Contient toutes les données consolidées des ADV, y compris les commentaires. Sert de livrable pour la direction et de source de commentaires historiques pour la semaine suivante. |

---

## Contexte du Projet

### Domaine et Complexité

- **Domaine** : Application desktop VBA/Excel
- **Complexité** : Faible-Moyenne
- **Composants** : 8-10 modules VBA

### Exigences Fonctionnelles (FR1-FR34)

Gestion session RAII, chargement données ERP, filtrage ADV, saisie commentaires, consolidation UPSERT, logging, configuration externe.

### Exigences Non-Fonctionnelles

| Exigence         | Cible                                      |
| ---------------- | ------------------------------------------ |
| Performance      | < 5 sec par opération, 800 affaires/ADV    |
| Fiabilité        | 100% disponibilité vendredi, 0% perte données |
| Maintenabilité   | Code compréhensible par non-experts VBA    |

### Contraintes Techniques

| Contrainte              | Valeur                                     |
| ----------------------- | ------------------------------------------ |
| Plateforme              | Windows + Excel 2016+                      |
| Infrastructure          | Active Directory uniquement (pas de cloud) |
| Persistance             | Fichiers Excel (pas de BDD)               |
| Concurrence             | Fichier unique partagé entre 3 ADV         |
| Budget                  | Pas d'investissement ERP                   |
| Outil développement VBA | xlManage (Python + pywin32)                |

---

## Règles Métier Immuables

> Ces règles sont **non-négociables**. Elles définissent le comportement attendu de l'application. Toute modification nécessite une validation métier.

### RÈGLE 1 : Unicité de l'Admin (CRITIQUE)

- **UN SEUL** utilisateur peut avoir `IsAdmin = Oui` dans tbADV
- Si deux admins détectés → **ERREUR BLOQUANTE** au démarrage (ERR-002)
- L'unicité est validée systématiquement par `modConfiguration`

### RÈGLE 2 : 1 Affaire = 1 ADV (CONCURRENCE)

- Chaque affaire appartient à **UN SEUL** ADV (plage exclusive)
- Pas de conflit de données possible (seulement conflit d'accès au fichier)
- Si un ADV est absent plus d'une semaine, ses affaires ne sont PAS mises à jour automatiquement
- L'admin doit consolider manuellement via Mode Admin pour les absences prolongées

### RÈGLE 3 : Validation Stricte du Mapping

- Toutes les colonnes du mapping doivent être présentes dans l'extraction ERP
- Vérification **AVANT** tout traitement
- Message d'erreur clair si colonne manquante (ERR-101)

### RÈGLE 4 : Extraction Repart à Zéro Chaque Année

- Le fichier d'extraction ERP repart à 0 affaires en début d'année
- Pas de problème de volume croissant à gérer
- Le fichier consolidé de l'année précédente n'est plus utilisé
- En début d'année, l'ADV peut ne pas avoir de fichier consolidé : les commentaires sont alors vides

---

## Architecture des Modules VBA

Chaque module a une responsabilité unique. Le code source est versionné dans `src/` sous forme de fichiers texte (`.bas` pour les modules standard, `.cls` pour les classes).

| Module                | Fichier source            | Responsabilité                                   |
| --------------------- | ------------------------- | ------------------------------------------------ |
| `clsOptimizer`        | `clsOptimizer.cls`        | Classe RAII (gestion état Excel)                 |
| `modUtils`            | `modUtils.bas`            | Helpers, constantes, gestion d'erreurs           |
| `modConfiguration`    | `modConfiguration.bas`    | Chargement configuration depuis data.xlsx (tbADV, tbParametres, tbMapping) |
| `modLogging`          | `modLogging.bas`          | Logging (INFO, ERREUR, SUCCES)                   |
| `modTimer`            | `modTimer.bas`            | Mesure de performance                            |
| `modExtraction`       | `modExtraction.bas`       | Sélection fichiers (consolidé + ERP) et chargement ERP |
| `modFiltrage`         | `modFiltrage.bas`         | Filtrage par trigramme ADV                       |
| `modConsolidation`    | `modConsolidation.bas`    | UPSERT + retry + sauvegardes (inclut colonne Commentaire) |
| `modCommentaires`     | `modCommentaires.bas`     | Extraction des commentaires depuis le fichier consolidé précédent |

---

## Structure de data.xlsx

> **RÈGLE CRITIQUE** : Chaque ListObject **DOIT** être isolé dans sa propre feuille. Un ListObject par feuille, pas plus.

| Feuille        | ListObject       | Rôle                                       |
| -------------- | ---------------- | ------------------------------------------ |
| ADV            | **tbADV**        | Utilisateurs et permissions                |
| Configuration  | **tbParametres** | Paramètres de l'application                |
| Mapping        | **tbMapping**    | Correspondance colonnes ERP / Suivi        |

> **Note :** Les commentaires historiques ne sont plus stockés dans data.xlsx. Ils sont désormais contenus dans le fichier consolidé précédent (voir section "Structure du Fichier Consolidé" ci-dessous).

### Détail des ListObjects

**tbADV** : `UserName | Nom | Prénom | Trigramme | IsAdmin`

**tbParametres** : `Parametre | Valeur | Description`

- CheminData, CheminExtraction, CheminConsolidation
- DelaiRetryMin (0), DelaiRetryMax (3), MaxTentatives (5)

**tbMapping** : `ColonneExtraction | ColonneSuivi | Type | Description`

- 16 colonnes mappées (Année, Mois, ADV, Affaire, CA prévu/réel, etc.)
- RepertoireConsolide (répertoire par défaut du dialogue de sélection du fichier consolidé)

---

## Structure du Fichier Consolidé

Le fichier consolidé joue un **double rôle** :

1. **Livrable direction** : fichier de suivi hebdomadaire transmis à la direction (ex : `Suivi affaires 2026-S05.xlsx`)
2. **Source de commentaires** : les commentaires saisis par les ADV sont stockés dans ce fichier et servent de source historique pour la semaine suivante

### Structure du ListObject principal

Le ListObject du fichier consolidé contient :
- Toutes les colonnes définies dans tbMapping (ColonneSuivi)
- Une colonne **Commentaire** (saisie ADV) qui contient les commentaires historiques

### Comportement au premier lancement

- Au premier lancement de l'année ou lors d'une première utilisation, l'ADV peut ne pas disposer d'un fichier consolidé précédent
- Dans ce cas, la sélection du fichier consolidé est ignorée (clic Annuler) et les commentaires sont vides pour toutes les affaires

### Validation du format

- Si le fichier consolidé sélectionné est invalide (structure incorrecte, colonnes manquantes) → **ERR-103** avec option de continuer sans commentaires ou de choisir un autre fichier
- La validation vérifie la présence du ListObject attendu et de la colonne Commentaire

---

## Authentification

L'application identifie l'utilisateur sans écran de connexion :

1. **Identification** : `Environ("USERNAME")` récupère le nom d'utilisateur Windows
2. **Vérification** : Recherche dans tbADV pour valider que l'utilisateur est autorisé
3. **Permissions** : Droits Active Directory restrictifs (répertoire `data/` uniquement)

Si l'utilisateur n'est pas trouvé dans tbADV → ERR-001.

---

## Gestion de la Concurrence

> **Contexte** : 3 ADV travaillent en parallèle sur un fichier partagé. Il n'y a pas de base de données, donc la concurrence est gérée au niveau du fichier Excel.

- **UPSERT incrémental** : Suppression des anciennes lignes de l'ADV puis ajout des nouvelles
- **Retry** : Si le fichier est verrouillé, l'application attend un délai aléatoire (0 à 3 secondes) puis réessaie, jusqu'à 5 tentatives maximum
- **Backup** : Sauvegarde automatique avant chaque consolidation dans `data/backups/`

---

## Format de Logging

Chaque action est tracée dans `tbAffaires.log` par `modLogging` :

```
DATE | USER | ACTION | RESULTAT
2026-01-23 14:32:15 | Patrick | Consolidation 50 affaires | SUCCES (0.8 sec)
```

---

## Codes d'Erreur

Chaque erreur a un code unique, un message explicite et des actions correctives pour l'utilisateur et l'admin.

| Code    | Description                            | Action Utilisateur             | Action Admin                     |
| ------- | -------------------------------------- | ------------------------------ | -------------------------------- |
| ERR-001 | Utilisateur non configuré              | Contacter Patrick              | Ajouter à tbADV                  |
| ERR-002 | Double admin détecté                   | Contacter Patrick              | Corriger tbADV                   |
| ERR-101 | Colonne mapping manquante              | Vérifier fichier               | Mettre à jour tbMapping          |
| ERR-102 | Fichier extraction introuvable         | Vérifier chemin                | Vérifier tbParametres            |
| ERR-103 | Format fichier consolidé invalide      | Continuer sans commentaires ou choisir un autre fichier | Vérifier la structure du fichier |
| ERR-201 | Fichier consolidation occupé           | Patienter/réessayer            | Vérifier qui a le fichier ouvert |
| ERR-202 | Échec consolidation après 5 tentatives | Ne pas fermer, appeler Patrick | Vérifier verrou fichier          |
| ERR-301 | Commentaire trop long                  | Raccourcir                     | -                                |
| ERR-401 | Mode Admin actif                       | Vérifier trigramme             | Confirmer usurpation             |

### Règles de Gestion des Erreurs

- Validation stricte du mapping **avant** chargement ERP (ERR-101)
- Validation unicité Admin **au démarrage** (ERR-002)
- Retry avec compteur visuel pour concurrence (ERR-201)
- Préservation des données saisies en cas d'échec (l'utilisateur ne perd jamais son travail)
- Log systématique de toutes les erreurs

---

## Flux de Données

Ce schéma montre le parcours des données du démarrage à la fin de session :

```
┌──────────────┐     ┌──────────────┐     ┌──────────────┐
│  data.xlsx   │     │ Consolidé    │     │ Extraction   │
│  (config)    │     │ précédent    │     │ ERP          │
│              │     │ (optionnel)  │     │              │
└──────┬───────┘     └──────┬───────┘     └──────┬───────┘
       │                    │                    │
       ▼                    │                    │
  1. modConfiguration       │                    │
     lit tbADV,             │                    │
     tbParametres,          │                    │
     tbMapping              │                    │
       │                    │                    │
       ▼                    │                    │
  2. Authentification       │                    │
     USERNAME →             │                    │
     vérif. tbADV           │                    │
                            ▼                    │
                       3. Dialogue 1             │
                          modExtraction          │
                          sélection consolidé    │
                          (OPTIONNEL)            │
                                                 ▼
                                            4. Dialogue 2
                                               modExtraction
                                               sélection ERP
                                               (OBLIGATOIRE)
                                                 │
                                                 ▼
                                            5. modExtraction
                                               charge ERP
                                               (lecture seule,
                                                validation mapping)
                                                 │
                            ┌────────────────────┘
                            ▼
                       6. modCommentaires
                          extrait commentaires
                          du consolidé (si fourni)
                          + modFiltrage
                          fusionne ERP + commentaires
                          filtre par trigramme ADV
                            │
                            ▼
                       7. L'ADV modifie
                          le ListObject
                          (Excel natif)
                            │
                            ▼                    ┌──────────────┐
                       8. modConsolidation       │ Fichier de   │
                          UPSERT (avec colonne   │ suivi partagé│
                          Commentaire)           └──────┬───────┘
                          + modLogging →                │
                            tbAffaires.log              │
                                                        ▼
```

### Comparaison Ancien / Nouveau Workflow

| Étape | Ancien workflow | Nouveau workflow |
| ----- | --------------- | ---------------- |
| Configuration | modConfiguration lit tbADV, tbParametres, tbMapping, **tbCommentaires** | modConfiguration lit tbADV, tbParametres, tbMapping (sans tbCommentaires) |
| Sélection fichiers | 1 dialogue : extraction ERP uniquement | 2 dialogues : consolidé précédent (optionnel) + extraction ERP |
| Source commentaires | tbCommentaires dans data.xlsx | Colonne Commentaire du fichier consolidé précédent |
| Consolidation | UPSERT + sauvegarde commentaires séparée | UPSERT avec colonne Commentaire incluse |
| Nombre d'étapes | 9 | 8 |

---

## Correspondance Exigences / Modules

| Catégorie     | Exigences | Modules impliqués                               |
| ------------- | --------- | ----------------------------------------------- |
| Session       | FR1-FR5   | clsOptimizer, modConfiguration, modUtils |
| Données       | FR6-FR11  | modExtraction (2 dialogues), modConfiguration   |
| Filtrage      | FR12-FR16 | modFiltrage                                     |
| Saisie        | FR17-FR19 | (Excel natif, pas de module VBA dédié)          |
| Consolidation | FR20-FR24 | modConsolidation, modUtils, modTimer            |
| Timer         | FR25-FR27 | modTimer                                        |
| Logging       | FR28-FR31 | modLogging                                      |
| Config        | FR32-FR34 | modConfiguration (via data.xlsx)                |
| Commentaires  | FR9, FR15 | modCommentaires (fichier consolidé précédent)   |

---

## Documentation Associée

### Guides Utilisateurs

| Document                              | Public               | Contenu                                                            |
| ------------------------------------- | -------------------- | ------------------------------------------------------------------ |
| `docs/guide-utilisateur.md`           | ADV (3 utilisateurs) | Procédure 5 étapes, problèmes courants, mode admin                 |
| `docs/guide-administrateur.md`        | Patrick (Admin)      | Configuration data.xlsx, points de vigilance, procédures d'urgence |
| `docs/points-vigilance-et-erreurs.md` | Dev + Admin          | Matrice des risques, codes erreur, stratégie de gestion d'erreurs  |

### Documentation Technique

| Document                                           | Contenu                                                   |
| -------------------------------------------------- | --------------------------------------------------------- |
| `docs/knowledge-base/guidelines/vba-guidelines.md` | Conventions de code VBA (Windows-1252, naming, structure) |
| `docs/knowledge-base/decisions/001-vba-toolkit.md` | Architecture du VBA Toolkit (post-développement)          |
| `project-context.md` (racine)                      | Règles d'implémentation pour SM et Dev                    |
