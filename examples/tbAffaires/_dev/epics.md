---
project_name: "tbAffaires"
date: "2026-02-08"
author: "Patrick"
source_documents:
    - "_dev/prd.md"
    - "_dev/architecture.md"
---

# tbAffaires - Epics & Stories

## Vue d'ensemble

Ce document décompose les exigences du PRD et de l'architecture en epics et stories implémentables. Chaque story est rattachée à un ou plusieurs modules VBA de l'architecture et à des exigences fonctionnelles (FR) ou non-fonctionnelles (NFR).

---

## Inventaire des Exigences

### Exigences Fonctionnelles

**Gestion de Session (FR1-FR5)**

- FR1: L'application s'initialise avec optimisation performances Excel (RAII)
- FR2: L'application identifie automatiquement l'utilisateur via username Windows
- FR3: L'application charge la configuration utilisateur depuis data.xlsx (tbADV)
- FR4: L'application affiche un message d'erreur si utilisateur non configuré
- FR5: L'application restaure l'état Excel à la fermeture (même en cas d'erreur)

**Chargement des Données (FR6-FR11)**

- FR6: L'ADV sélectionne le fichier d'extraction ERP via boîte de dialogue Windows
- FR7: L'application charge le fichier d'extraction en lecture seule
- FR8: L'application charge le mapping des colonnes depuis data.xlsx (tbMapping)
- FR9: L'application charge les commentaires historiques depuis le fichier consolidé précédent (colonne Commentaire)
- FR10: L'application crée automatiquement le fichier de suivi s'il n'existe pas
- FR11: L'application affiche un message d'erreur si colonne mappée introuvable

**Filtrage et Affichage (FR12-FR16)**

- FR12: L'application filtre les affaires par trigramme ADV de l'utilisateur connecté (ou usurpé en mode Admin)
- FR13: L'application affiche les affaires dans un ListObject temporaire
- FR14: L'application met en évidence les affaires en difficulté financière (rouge)
- FR15: L'application pré-remplit les commentaires existants de S-1
- FR16: L'ADV navigue avec fonctionnalités Excel natives (filtres, tri, Ctrl+F)

**Saisie des Commentaires (FR17-FR18)**

- FR17: L'ADV saisit de nouveaux commentaires directement dans le ListObject (colonne Commentaire déverrouillée, reste du classeur verrouillé)
- FR18: L'ADV modifie les commentaires existants

**Consolidation (FR20-FR24)**

- FR20: L'ADV déclenche la consolidation de ses données
- FR21: L'application supprime les anciennes données ADV avant ajout (UPSERT)
- FR22: L'application réessaie automatiquement si fichier verrouillé (retry 0-3s, 5 max)
- FR23: L'application affiche message d'erreur après 5 échecs consolidation
- FR24: L'application préserve les données saisies même en cas d'échec

**Mesure de Performance (FR25-FR27)**

- FR25: L'application mesure le temps des opérations critiques
- FR26: L'application affiche le temps écoulé dans le message de résultat
- FR27: L'application enregistre les temps dans les logs

**Logging et Traçabilité (FR28-FR31)**

- FR28: L'application enregistre chaque action dans un fichier de log
- FR29: L'application enregistre les erreurs avec contexte (qui, quand, quoi)
- FR30: L'application distingue les niveaux de log (INFO, ERREUR, SUCCES)
- FR31: L'Admin consulte le fichier de logs pour diagnostiquer les problèmes

**Configuration et Administration (FR32-FR34)**

- FR32: L'Admin modifie le mapping colonnes sans toucher au code VBA
- FR33: L'Admin ajoute/modifie des utilisateurs dans data.xlsx (tbADV)
- FR34: L'Admin configure les paramètres dans data.xlsx (tbParametres)

**Mode Admin (FR35-FR37)**

- FR35: L'application identifie les utilisateurs admin via la colonne IsAdmin dans tbADV
- FR36: L'Admin peut choisir de travailler au nom d'un autre ADV via une boîte de dialogue
- FR37: Le logging indique "Action par [Admin] au nom de [Utilisateur usurpé]"

### Exigences Non-Fonctionnelles

**Performance (NFR1-NFR4)**

- NFR1: Chargement extraction < 5 secondes
- NFR2: Chargement commentaires < 5 secondes
- NFR3: Consolidation UPSERT < 5 secondes
- NFR4: Interface réactive pendant opérations (pas de freeze > 1 sec)

**Fiabilité (NFR6-NFR9)**

- NFR6: Fonctionnement chaque vendredi sans échec (100% disponibilité hebdo)
- NFR7: Données saisies jamais perdues (0% perte de données)
- NFR8: État Excel restauré même en cas de crash (RAII)
- NFR9: Gestion conflits verrouillage fichier (5 tentatives max, délai 0-3s)

**Maintenabilité (NFR10-NFR13)**

- NFR10: Code compréhensible par non-expert VBA (fonctions nommées explicitement)
- NFR11: Modifications mapping sans toucher au code (100% via data.xlsx)
- NFR12: Messages d'erreur indiquent cause et action (Format: "Erreur + Solution")
- NFR13: Logs permettent diagnostic rapide (Format: Date, User, Action, Résultat)

**Sécurité (NFR14-NFR15)**

- NFR14: Seuls utilisateurs configurés peuvent utiliser (vérification tbADV au démarrage)
- NFR15: Permissions AD restreignent accès aux fichiers (ADV : data/ uniquement)

---

## Liste des Epics

| Epic | Titre                                  | Stories   | FRs couverts         | NFRs couverts          |
| ---- | -------------------------------------- | --------- | -------------------- | ---------------------- |
| 1    | Infrastructure et Session Utilisateur  | 1.1 - 1.6 | FR1-FR5, FR32-FR37   | NFR8, NFR10-NFR15      |
| 2    | Chargement et Préparation des Données  | 2.1 - 2.5 | FR6-FR11             | NFR1, NFR2             |
| 3    | Filtrage et Visualisation des Affaires | 3.1 - 3.4 | FR12-FR16            | NFR4                   |
| 4    | Saisie et Consolidation                | 4.1 - 4.4 | FR17-FR18, FR20-FR24 | NFR3, NFR6, NFR7, NFR9 |
| 5    | Logging et Observabilité               | 5.1 - 5.3 | FR25-FR31, FR37      | NFR13                  |

---

## Correspondance Modules VBA / Stories

| Module VBA            | Fichier source                | Stories           |
| --------------------- | ----------------------------- | ----------------- |
| `modUtils`            | `src/modUtils.bas`            | **1.6**           |
| `clsOptimizer`        | `src/clsOptimizer.cls`        | **1.2**           |
| `modConfiguration`    | `src/modConfiguration.bas`    | **1.1**, 1.3, 1.4 |
| `modLogging`          | `src/modLogging.bas`          | **5.1**           |
| `modTimer`            | `src/modTimer.bas`            | **5.2**           |
| `modExtraction`       | `src/modExtraction.bas`       | **2.1**, 2.2, 2.3 |
| `modCommentaires`     | `src/modCommentaires.bas`     | **2.4**           |
| `modFiltrage`         | `src/modFiltrage.bas`         | **3.1**           |
| `modConsolidation`    | `src/modConsolidation.bas`    | **4.2**, 4.3, 4.4 |

> **Gras** = story principale de création du module. Les autres stories étendent le module.

---

## Couverture FR / Stories

| FR   | Story | Description courte                  |
| ---- | ----- | ----------------------------------- |
| FR1  | 1.2   | Initialisation RAII                 |
| FR2  | 1.3   | Identification Windows              |
| FR3  | 1.3   | Chargement config utilisateur       |
| FR4  | 1.3   | Erreur si non configuré             |
| FR5  | 1.2   | Restauration état Excel             |
| FR6  | 2.1   | Sélection fichier consolidé + ERP   |
| FR7  | 2.3   | Chargement lecture seule            |
| FR8  | 2.2   | Chargement mapping                  |
| FR9  | 2.4   | Chargement commentaires historiques |
| FR10 | 2.5   | Création auto fichier suivi         |
| FR11 | 2.2   | Erreur colonne manquante            |
| FR12 | 3.1   | Filtrage par trigramme              |
| FR13 | 3.2   | Affichage ListObject                |
| FR14 | 3.3   | Mise en évidence difficultés        |
| FR15 | 3.4   | Pré-remplissage commentaires S-1    |
| FR16 | 3.2   | Navigation Excel native             |
| FR17 | 4.1   | Saisie commentaires                 |
| FR18 | 4.1   | Modification commentaires           |
| FR20 | 4.2   | Déclenchement consolidation         |
| FR21 | 4.2   | UPSERT suppression + ajout          |
| FR22 | 4.3   | Retry automatique                   |
| FR23 | 4.3   | Message erreur après 5 échecs       |
| FR24 | 4.4   | Préservation données                |
| FR25 | 5.2   | Mesure temps opérations             |
| FR26 | 5.2   | Affichage temps résultat            |
| FR27 | 5.2   | Enregistrement temps dans logs      |
| FR28 | 5.1   | Log des actions                     |
| FR29 | 5.1   | Log erreurs avec contexte           |
| FR30 | 5.1   | Distinction niveaux log             |
| FR31 | 5.3   | Consultation logs par Admin         |
| FR32 | 1.1   | Modification mapping sans code      |
| FR33 | 1.1   | Gestion utilisateurs data.xlsx      |
| FR34 | 1.1   | Configuration paramètres            |
| FR35 | 1.4   | Identification mode Admin           |
| FR36 | 1.5   | Usurpation utilisateur              |
| FR37 | 5.1   | Log spécifique mode Admin           |

---

## Ordre d'implémentation recommandé

L'ordre suit les dépendances entre modules. Chaque étape suppose les précédentes terminées.

```
Phase 1 - Fondations
- [X]  Story 1.6  modUtils (constantes, helpers, gestion d'erreurs)
- [X]  Story 5.1  modLogging (logging utilisé par tous les modules)
- [X]  Story 1.2  clsOptimizer (RAII)

Phase 2 - Configuration et Session
- [ ]    Story 1.1  Structure fichiers + data.xlsx (tbADV, tbParametres, tbMapping)
- [ ]    Story 1.3  modConfiguration + identification utilisateur
- [ ]    Story 1.4  Validation unicité Admin

Phase 3 - Chargement des Données
- [ ]    Story 2.1  Sélection fichiers (dialogues consolidé + ERP)
- [ ]    Story 2.2  Validation mapping colonnes
- [ ]    Story 2.3  Import données ERP en lecture seule
- [ ]    Story 2.4  Récupération commentaires historiques (modCommentaires)
- [ ]    Story 2.5  Création automatique fichier de suivi

Phase 4 - Filtrage et Affichage
- [ ]    Story 3.1  Filtrage par trigramme ADV (modFiltrage)
- [ ]    Story 3.2  Création ListObject temporaire
- [ ]    Story 3.3  Mise en évidence affaires en difficulté
- [ ]    Story 3.4  Pré-remplissage commentaires

Phase 5 - Saisie et Consolidation
- [ ]    Story 4.1  Protection classeur + saisie commentaires
- [ ]    Story 4.2  UPSERT incrémental (modConsolidation)
- [ ]    Story 4.3  Retry automatique + gestion verrous
- [ ]    Story 4.4  Backup + préservation données

Phase 6 - Finalisation
- [ ]    Story 5.2  Mesure de performance (modTimer)
- [ ]    Story 5.3  Consultation et diagnostic
- [ ]    Story 1.5  Mode Admin avec usurpation
```

---

## Epic 1: Infrastructure et Session Utilisateur

L'utilisateur peut démarrer l'application, être identifié automatiquement, et travailler dans un environnement Excel optimisé et sécurisé.

**FRs couverts :** FR1, FR2, FR3, FR4, FR5, FR32, FR33, FR34, FR35, FR36, FR37
**NFRs couverts :** NFR8 (RAII), NFR10-NFR13 (maintenabilité), NFR14-NFR15 (sécurité)

### Story 1.1: Créer la structure de fichiers et la configuration initiale

**En tant qu'** Admin,
**Je veux** créer la structure de fichiers et le fichier data.xlsx avec les tables de configuration,
**Afin de** pouvoir configurer l'application sans modifier le code VBA.

**Module VBA :** `modConfiguration` (lecture de la config au démarrage)

**Critères d'Acceptation :**

**Étant donné** que je suis sur le serveur AD
**Quand** je crée la structure dans `\\serveur-ad\FRV\AFFAIRES\01 SUIVI AFFAIRES\`
**Alors** les dossiers `data/`, `extractions/`, `data/backups/` existent

**Étant donné** que je crée le fichier data.xlsx
**Quand** j'ouvre le fichier
**Alors** il contient 3 feuilles (ADV, Configuration, Mapping) avec un ListObject par feuille

**Étant donné** que je configure tbADV
**Quand** j'ajoute les utilisateurs
**Alors** la table contient : UserName | Nom | Prénom | Trigramme | IsAdmin

**Étant donné** que je configure tbParametres
**Quand** j'ajoute les paramètres
**Alors** la table contient les entrées : CheminData, CheminExtraction, CheminConsolidation, RepertoireConsolide, DelaiRetryMin (0), DelaiRetryMax (3), MaxTentatives (5)

**Étant donné** que je configure tbMapping
**Quand** j'ajoute les 16 colonnes
**Alors** la table contient : ColonneExtraction | ColonneSuivi | Type | Description

---

### Story 1.2: Implémenter le pattern RAII avec clsOptimizer

**En tant qu'** ADV,
**Je veux** que l'application optimise Excel au démarrage et restaure l'état à la fermeture,
**Afin de** bénéficier de performances maximales et éviter les problèmes d'état Excel.

**Module VBA :** `clsOptimizer` (`src/clsOptimizer.cls`)

**Critères d'Acceptation :**

**Étant donné** que je démarre tbAffaires.xlsm
**Quand** l'ApplicationState s'initialise (`Class_Initialize`)
**Alors** les valeurs originales de ScreenUpdating, Calculation, EnableEvents et DisplayAlerts sont sauvegardées, puis ScreenUpdating, Calculation (xlCalculationManual) et Events sont désactivés

**Étant donné** que je ferme l'application (même en cas d'erreur)
**Quand** l'objet ApplicationState est détruit (`Class_Terminate`)
**Alors** l'état initial d'Excel est restauré (ScreenUpdating, Calculation, Events, DisplayAlerts)

**Étant donné** qu'une erreur survient pendant l'exécution
**Quand** l'erreur est interceptée
**Alors** l'état Excel est toujours restauré grâce au pattern RAII (destruction automatique de l'objet en fin de portée)

---

### Story 1.3: Identifier automatiquement l'utilisateur Windows

**En tant qu'** ADV,
**Je veux** être identifié automatiquement via mon username Windows,
**Afin de** ne pas avoir à saisir mes identifiants.

**Module VBA :** `modConfiguration` (`src/modConfiguration.bas`)

**Critères d'Acceptation :**

**Étant donné** que j'ouvre tbAffaires.xlsm
**Quand** l'application démarre
**Alors** `Environ("USERNAME")` est récupéré automatiquement

**Étant donné** que mon username existe dans tbADV
**Quand** l'identification réussit
**Alors** mon trigramme ADV, nom, prénom et statut admin sont chargés en mémoire (variables module)

**Étant donné** que mon username n'existe pas dans tbADV
**Quand** l'identification échoue
**Alors** un message d'erreur ERR-001 s'affiche : "Utilisateur non configuré. Contacter Patrick." et l'application se ferme proprement

---

### Story 1.4: Valider l'unicité de l'Admin

**En tant qu'** Admin,
**Je veux** que le système vérifie qu'il n'y a qu'un seul admin configuré,
**Afin d'** éviter les conflits de gestion.

**Module VBA :** `modConfiguration` (`src/modConfiguration.bas`)

**Critères d'Acceptation :**

**Étant donné** que le fichier data.xlsx est chargé
**Quand** l'application vérifie les admins
**Alors** elle compte les lignes avec IsAdmin = "Oui" dans tbADV

**Étant donné** qu'il y a exactement 1 admin
**Quand** la validation passe
**Alors** l'application continue normalement

**Étant donné** qu'il y a 0 ou 2+ admins
**Quand** la validation échoue
**Alors** une ERREUR BLOQUANTE ERR-002 s'affiche : "Double admin détecté. Contacter Patrick." et l'application se ferme proprement

---

### Story 1.5: Implémenter le mode Admin avec usurpation

**En tant qu'** Admin,
**Je veux** pouvoir travailler au nom d'un autre ADV,
**Afin de** gérer les absences ou problèmes utilisateurs.

**Module VBA :** `modConfiguration` (`src/modConfiguration.bas`)

**Critères d'Acceptation :**

**Étant donné** que je suis identifié comme admin (IsAdmin = "Oui")
**Quand** l'application démarre
**Alors** une boîte de dialogue me propose de choisir un ADV à usurper (liste des trigrammes depuis tbADV) ou de travailler sous mon propre trigramme

**Étant donné** que je choisis de travailler pour un autre ADV
**Quand** je sélectionne son trigramme
**Alors** le trigramme actif est remplacé par celui de l'ADV usurpé pour toute la session

**Étant donné** que je travaille en mode admin usurpé
**Quand** j'effectue une action
**Alors** le log indique "Action par [Admin] au nom de [Utilisateur usurpé]" (FR37)

**Étant donné** que je suis en mode admin usurpé
**Quand** je consulte l'interface
**Alors** une alerte visuelle permanente indique "Mode Admin actif - Trigramme : [XXX]" (ERR-401)

---

### Story 1.6: Implémenter le module utilitaires fondamentaux (modUtils)

**En tant que** développeur,
**Je veux** disposer d'un module utilitaire centralisé avec les constantes, helpers et gestion d'erreurs,
**Afin de** factoriser le code commun et standardiser le comportement des erreurs dans tous les modules.

**Module VBA :** `modUtils` (`src/modUtils.bas`)

> **Note :** Ce module est une **dépendance de tous les autres modules**. Il doit être implémenté en premier.

**Critères d'Acceptation :**

**Étant donné** que le projet démarre
**Quand** le module modUtils est créé
**Alors** il contient les constantes publiques des codes d'erreur (ERR_001 à ERR_401) avec leurs messages standardisés

**Étant donné** qu'un module a besoin de vérifier l'existence d'un fichier ou répertoire
**Quand** il appelle `FichierExiste()` ou `RepertoireExiste()`
**Alors** la fonction retourne True ou False sans erreur

**Étant donné** qu'un module a besoin de charger un Workbook, Worksheet ou ListObject
**Quand** il appelle `ChargerWorkbook()`, `ChargerWorksheet()` ou `ChargerListObject()`
**Alors** la fonction retourne l'objet ou Nothing avec un message d'erreur explicite (NFR12)

**Étant donné** qu'une erreur applicative survient
**Quand** `AfficherMessageErreur()` est appelé avec un code d'erreur
**Alors** un MsgBox s'affiche avec : le code d'erreur, le message explicite et l'action suggérée pour l'utilisateur

**Étant donné** qu'un message d'information doit être affiché
**Quand** `AfficherMessageInfo()` est appelé
**Alors** un MsgBox informatif s'affiche et l'action est loggée

---

## Epic 2: Chargement et Préparation des Données

L'ADV peut charger son fichier d'extraction ERP et retrouver automatiquement ses commentaires historiques.

**FRs couverts :** FR6, FR7, FR8, FR9, FR10, FR11
**NFRs couverts :** NFR1, NFR2 (performance < 5s)

### Story 2.1: Sélectionner le fichier consolidé et le fichier d'extraction ERP

**En tant qu'** ADV,
**Je veux** sélectionner le fichier consolidé précédent puis mon fichier d'extraction ERP via des boîtes de dialogue Windows,
**Afin de** charger mes commentaires historiques et mes données à traiter.

**Module VBA :** `modExtraction` (`src/modExtraction.bas`)

**Critères d'Acceptation :**

**Étant donné** que l'application est démarrée et que je suis identifié
**Quand** l'application me demande de sélectionner le fichier consolidé précédent
**Alors** une boîte de dialogue Windows (`GetOpenFilename`) s'ouvre sur le répertoire RepertoireConsolide configuré dans tbParametres

**Étant donné** que c'est ma première utilisation de l'année
**Quand** je clique sur Annuler dans le dialogue du fichier consolidé
**Alors** l'application continue sans commentaires historiques (colonne Commentaire vide)

**Étant donné** que j'ai terminé la sélection du fichier consolidé (ou annulé)
**Quand** l'application me demande de sélectionner l'extraction ERP
**Alors** une boîte de dialogue Windows s'ouvre sur le répertoire CheminExtraction configuré dans tbParametres

**Étant donné** que la boîte de dialogue ERP est ouverte
**Quand** je sélectionne un fichier .xlsx
**Alors** le chemin du fichier est mémorisé pour le chargement

**Étant donné** que je clique sur Annuler dans la boîte de dialogue ERP
**Quand** l'application détecte l'annulation
**Alors** un message d'erreur ERR-102 s'affiche et l'application se ferme proprement (l'extraction ERP est obligatoire)

---

### Story 2.2: Charger et valider le mapping des colonnes

**En tant qu'** ADV,
**Je veux** que l'application charge le mapping des colonnes depuis data.xlsx,
**Afin de** pouvoir adapter l'application si les colonnes ERP changent sans modifier le code.

**Module VBA :** `modExtraction` (`src/modExtraction.bas`) + `modConfiguration` (lecture tbMapping)

**Critères d'Acceptation :**

**Étant donné** que je déclenche le chargement de l'extraction
**Quand** l'application démarre le traitement
**Alors** elle charge d'abord tbMapping depuis data.xlsx (16 correspondances ColonneExtraction / ColonneSuivi)

**Étant donné** que le mapping est chargé
**Quand** l'application vérifie les colonnes du fichier ERP
**Alors** elle valide que **toutes** les colonnes définies dans tbMapping existent dans l'extraction (**AVANT** tout traitement - Règle Métier 3)

**Étant donné** qu'une colonne du mapping est introuvable dans l'extraction ERP
**Quand** la validation échoue
**Alors** un message d'erreur ERR-101 s'affiche : "Colonne '[NomColonne]' non trouvée dans l'extraction. Vérifiez le mapping dans data.xlsx." avec le nom exact de la colonne manquante

**Étant donné** que toutes les colonnes sont présentes
**Quand** la validation réussit
**Alors** l'application procède au chargement des données

---

### Story 2.3: Importer les données d'extraction en lecture seule

**En tant qu'** ADV,
**Je veux** que l'application charge mon fichier d'extraction ERP en lecture seule,
**Afin de** ne pas risquer de modifier les données source.

**Module VBA :** `modExtraction` (`src/modExtraction.bas`)

**Critères d'Acceptation :**

**Étant donné** que le mapping est validé
**Quand** l'application charge le fichier ERP
**Alors** elle l'ouvre en mode lecture seule (`ReadOnly:=True`)

**Étant donné** que le fichier est chargé
**Quand** les données sont transférées dans un tableau en mémoire
**Alors** l'application ferme le fichier ERP sans sauvegarder

**Étant donné** que le fichier contient ~800 affaires pour mon trigramme
**Quand** le chargement se termine
**Alors** le temps total est inférieur à 5 secondes (NFR1)

**Étant donné** que le fichier ERP est chargé
**Quand** les données sont importées
**Alors** seules les colonnes définies dans tbMapping sont conservées (renommées selon ColonneSuivi)

---

### Story 2.4: Récupérer les commentaires historiques

**En tant qu'** ADV,
**Je veux** retrouver automatiquement mes commentaires de la semaine précédente,
**Afin de** ne pas avoir à les recopier manuellement (~100 commentaires).

**Module VBA :** `modCommentaires` (`src/modCommentaires.bas`)

**Critères d'Acceptation :**

**Étant donné** que les données ERP sont chargées et qu'un fichier consolidé a été sélectionné
**Quand** l'application charge les commentaires historiques
**Alors** elle lit la colonne Commentaire du ListObject du fichier consolidé précédent

**Étant donné** que je suis identifié avec mon trigramme (ex: VC)
**Quand** les commentaires sont filtrés
**Alors** seuls les commentaires de mes affaires sont récupérés (correspondance par numéro d'affaire)

**Étant donné** qu'une affaire existe dans l'extraction avec un commentaire historique
**Quand** les données sont fusionnées
**Alors** le commentaire historique est associé à l'affaire correspondante

**Étant donné** qu'aucun fichier consolidé n'a été sélectionné (Annuler au dialogue 1)
**Quand** l'application prépare les commentaires
**Alors** la colonne Commentaire est vide pour toutes les affaires

**Étant donné** que le fichier consolidé a un format invalide (ListObject manquant, colonne Commentaire absente)
**Quand** la validation échoue
**Alors** ERR-103 s'affiche avec option "Continuer sans commentaires" ou "Choisir un autre fichier"

**Étant donné** que le chargement des commentaires est lancé
**Quand** il se termine
**Alors** le temps est inférieur à 5 secondes (NFR2)

---

### Story 2.5: Créer automatiquement le fichier de suivi

**En tant qu'** ADV,
**Je veux** que l'application crée automatiquement le fichier de suivi s'il n'existe pas,
**Afin de** ne pas avoir à le créer manuellement la première fois.

**Module VBA :** `modConsolidation` (`src/modConsolidation.bas`)

**Critères d'Acceptation :**

**Étant donné** que je déclenche la première consolidation de la semaine
**Quand** l'application vérifie l'existence du fichier de suivi
**Alors** elle cherche le fichier selon le chemin configuré dans tbParametres (CheminConsolidation)

**Étant donné** que le fichier de suivi n'existe pas
**Quand** la vérification échoue
**Alors** l'application crée un nouveau fichier basé sur `modèle.xltx` (format direction)

**Étant donné** que le fichier est créé depuis le modèle
**Quand** il est ouvert
**Alors** il contient un ListObject avec toutes les colonnes définies dans tbMapping (ColonneSuivi) + la colonne Commentaire

**Étant donné** que le fichier existe déjà
**Quand** la vérification réussit
**Alors** l'application utilise le fichier existant pour l'UPSERT

---

## Epic 3: Filtrage et Visualisation des Affaires

L'ADV visualise uniquement SES affaires avec les commentaires pré-remplis et les alertes visuelles.

**FRs couverts :** FR12, FR13, FR14, FR15, FR16
**NFRs couverts :** NFR4 (interface réactive)

### Story 3.1: Filtrer les affaires par trigramme ADV

**En tant qu'** ADV,
**Je veux** voir uniquement les affaires qui me sont assignées,
**Afin de** ne pas être distrait par les affaires des autres ADV.

**Module VBA :** `modFiltrage` (`src/modFiltrage.bas`)

**Critères d'Acceptation :**

**Étant donné** que les données ERP sont chargées (avec commentaires fusionnés)
**Quand** l'application filtre les affaires
**Alors** seules les lignes avec mon trigramme ADV (colonne ADV du mapping) sont conservées

**Étant donné** que je suis en mode admin usurpé (ex: pour l'ADV "HL")
**Quand** le filtrage s'applique
**Alors** seules les affaires de "HL" sont conservées, pas les miennes

**Étant donné** que ~800 affaires correspondent à mon trigramme
**Quand** le filtrage est appliqué
**Alors** le temps de traitement est inférieur à 1 seconde (NFR4)

**Étant donné** qu'aucune affaire ne correspond à mon trigramme
**Quand** le filtrage retourne 0 résultats
**Alors** un message s'affiche : "Aucune affaire trouvée pour votre trigramme [XXX]" et l'application continue (ListObject vide)

---

### Story 3.2: Créer et afficher le ListObject temporaire

**En tant qu'** ADV,
**Je veux** visualiser mes affaires dans un tableau Excel structuré (ListObject),
**Afin de** pouvoir utiliser les fonctionnalités natives d'Excel (filtres, tri, recherche).

**Module VBA :** `modFiltrage` (`src/modFiltrage.bas`)

**Critères d'Acceptation :**

**Étant donné** que les données sont filtrées par trigramme
**Quand** l'application crée l'affichage
**Alors** un ListObject temporaire est créé dans une feuille dédiée de tbAffaires.xlsm avec toutes les colonnes du mapping + Commentaire

**Étant donné** que le ListObject est créé
**Quand** je regarde l'affichage
**Alors** les colonnes sont ordonnées selon tbMapping : Année, Mois, ADV, Affaire, Client, CA Prévu, CA Réel, etc., Commentaire

**Étant donné** que le classeur est affiché
**Quand** je consulte mes affaires
**Alors** toutes les cellules sont verrouillées SAUF la colonne "Commentaire" (protection feuille activée)

**Étant donné** que le ListObject est affiché
**Quand** j'utilise Ctrl+F, les filtres auto ou le tri natifs Excel
**Alors** ces fonctionnalités fonctionnent normalement (FR16)

---

### Story 3.3: Mettre en évidence les affaires en difficulté financière

**En tant qu'** ADV,
**Je veux** identifier visuellement les affaires avec un résultat financier critique,
**Afin de** prioriser mon attention sur les problèmes urgents.

**Module VBA :** `modFiltrage` (`src/modFiltrage.bas`)

**Critères d'Acceptation :**

**Étant donné** que les affaires sont affichées dans le ListObject
**Quand** une affaire a un résultat financier négatif ou critique
**Alors** toute la ligne est mise en rouge (fond de cellule)

**Étant donné** que plusieurs affaires sont en difficulté
**Quand** je consulte le tableau
**Alors** elles sont visuellement distinctes des affaires sans alerte

**Étant donné** qu'une affaire n'est pas en difficulté
**Quand** elle s'affiche
**Alors** elle conserve la mise en forme standard (pas de rouge)

---

### Story 3.4: Pré-remplir les commentaires historiques

**En tant qu'** ADV,
**Je veux** retrouver mes commentaires de la semaine précédente déjà présents dans le tableau,
**Afin de** les modifier au lieu de les réécrire.

**Module VBA :** `modFiltrage` (`src/modFiltrage.bas`) + `modCommentaires`

**Critères d'Acceptation :**

**Étant donné** que les commentaires historiques sont chargés depuis le fichier consolidé précédent (Story 2.4)
**Quand** le ListObject est créé
**Alors** la colonne "Commentaire" contient les valeurs historiques pour chaque affaire

**Étant donné** qu'une affaire n'a pas de commentaire historique
**Quand** elle s'affiche
**Alors** la cellule Commentaire est vide (pas d'erreur, pas de placeholder)

**Étant donné** qu'une affaire a un commentaire historique
**Quand** je consulte la ligne
**Alors** je peux modifier directement le commentaire dans la cellule déverrouillée

**Étant donné** que le commentaire est pré-rempli
**Quand** je le modifie
**Alors** la nouvelle valeur remplace l'ancienne dans le ListObject (pas encore consolidée)

---

## Epic 4: Saisie et Consolidation

L'ADV saisit ses commentaires et les consolide de manière fiable dans le fichier de suivi partagé, même en cas de conflit d'accès entre ADV.

**FRs couverts :** FR17, FR18, FR20, FR21, FR22, FR23, FR24
**NFRs couverts :** NFR3 (consolidation < 5s), NFR6 (disponibilité), NFR7 (0% perte données), NFR9 (retry)

### Story 4.1: Protéger le classeur et permettre la saisie des commentaires

**En tant qu'** ADV,
**Je veux** pouvoir saisir et modifier mes commentaires dans la colonne dédiée tout en étant protégé contre les modifications accidentelles,
**Afin de** travailler sereinement sans risquer de corrompre les données ERP.

**Module VBA :** Pas de module dédié (utilise la protection native Excel configurée par modFiltrage lors de la création du ListObject)

**Critères d'Acceptation :**

**Étant donné** que le ListObject est affiché avec mes affaires
**Quand** je regarde le classeur
**Alors** la feuille est protégée : toutes les cellules sont verrouillées SAUF la colonne "Commentaire"

**Étant donné** que la colonne Commentaire est déverrouillée
**Quand** je clique sur une cellule Commentaire
**Alors** je peux saisir un nouveau commentaire ou modifier le commentaire existant (FR17, FR18)

**Étant donné** que je clique sur une cellule d'une autre colonne (Affaire, CA, ADV, etc.)
**Quand** j'essaie de modifier
**Alors** Excel affiche un message de protection et refuse la modification

**Étant donné** que je saisis un commentaire
**Quand** je valide (Entrée ou Tab)
**Alors** la valeur est stockée dans le ListObject en mémoire (pas encore consolidée dans le fichier partagé)

---

### Story 4.2: Implémenter l'UPSERT incrémental

**En tant qu'** ADV,
**Je veux** que mes données soient consolidées dans le fichier de suivi partagé via UPSERT,
**Afin de** mettre à jour mes affaires sans affecter celles des autres ADV.

**Module VBA :** `modConsolidation` (`src/modConsolidation.bas`)

**Critères d'Acceptation :**

**Étant donné** que je clique sur le bouton "Consolider"
**Quand** l'application ouvre le fichier de suivi partagé
**Alors** elle supprime d'abord toutes mes anciennes lignes dans le ListObject du fichier partagé (filtrées par mon trigramme ADV)

**Étant donné** que mes anciennes lignes sont supprimées
**Quand** l'application ajoute mes nouvelles données
**Alors** toutes mes affaires (données ERP + commentaires saisis) sont insérées à la fin du ListObject du fichier partagé

**Étant donné** que l'UPSERT est terminé
**Quand** le fichier est sauvegardé et fermé
**Alors** un message "Consolidation réussie - [N] affaires consolidées" s'affiche

**Étant donné** que la consolidation est en cours
**Quand** elle se termine
**Alors** le temps total (ouverture + suppression + insertion + sauvegarde) est inférieur à 5 secondes (NFR3)

**Étant donné** que je suis en mode admin usurpé pour l'ADV "HL"
**Quand** je consolide
**Alors** les lignes de "HL" sont supprimées/ajoutées dans le fichier partagé (pas les miennes)
**Et** une alerte visuelle confirme "Consolidation pour [HL] en mode Admin"

---

### Story 4.3: Implémenter le retry automatique et la gestion des verrous

**En tant qu'** ADV,
**Je veux** que l'application réessaie automatiquement si le fichier est verrouillé par un autre ADV,
**Afin de** ne pas avoir à relancer manuellement la consolidation.

**Module VBA :** `modConsolidation` (`src/modConsolidation.bas`)

**Critères d'Acceptation :**

**Étant donné** que je déclenche la consolidation
**Quand** le fichier de suivi est verrouillé par un autre ADV
**Alors** un message s'affiche avec un compteur visuel : "Fichier occupé. Tentative 1/5..." (ERR-201)

**Étant donné** qu'une tentative d'ouverture échoue
**Quand** l'application réessaie
**Alors** elle attend un délai aléatoire entre DelaiRetryMin (0s) et DelaiRetryMax (3s) configurés dans tbParametres

**Étant donné** que le fichier se libère entre les tentatives
**Quand** l'application réessaie avec succès
**Alors** la consolidation se poursuit normalement (UPSERT) et un message de succès s'affiche

**Étant donné** que MaxTentatives (5) tentatives ont toutes échoué
**Quand** la dernière tentative échoue
**Alors** un message d'erreur ERR-202 s'affiche : "Échec après 5 tentatives. Ne fermez pas l'application, vos données sont préservées. Contactez Patrick."

**Étant donné** que les tentatives sont en cours
**Quand** chaque tentative est effectuée
**Alors** le log enregistre chaque tentative avec horodatage et résultat

---

### Story 4.4: Sauvegarder automatiquement et préserver les données en cas d'échec

**En tant qu'** ADV,
**Je veux** que mes données saisies soient préservées et que le fichier soit sauvegardé avant toute modification,
**Afin de** ne jamais perdre mon travail.

**Module VBA :** `modConsolidation` (`src/modConsolidation.bas`)

**Critères d'Acceptation :**

**Étant donné** que je déclenche la consolidation
**Quand** l'application ouvre le fichier de suivi avec succès
**Alors** une copie de sauvegarde est créée dans `data/backups/` avec un nom horodaté (format `AAAAMMDD_HHMMSS_SuiviAffaires.xlsx`) **AVANT** toute modification

**Étant donné** que la consolidation échoue (après 5 tentatives de retry)
**Quand** l'erreur ERR-202 est affichée
**Alors** le ListObject temporaire avec mes données saisies reste affiché et modifiable dans tbAffaires.xlsm

**Étant donné** que la consolidation a échoué et que mes données sont préservées
**Quand** je retente la consolidation plus tard (re-clic sur "Consolider")
**Alors** les données saisies (y compris les commentaires modifiés) sont toujours présentes et consolidées

**Étant donné** qu'une erreur inattendue survient pendant la consolidation (erreur VBA, crash)
**Quand** la gestion d'erreur intercepte l'exception
**Alors** l'état Excel est restauré par le RAII (clsOptimizer) et le fichier de suivi partagé n'est pas corrompu (fermeture sans sauvegarde en cas d'erreur pendant l'UPSERT)

**Étant donné** que le backup a été créé avant la consolidation
**Quand** une corruption du fichier partagé est détectée
**Alors** l'Admin peut restaurer le fichier depuis `data/backups/`

---

## Epic 5: Logging et Observabilité

L'application trace toutes les actions et mesure les performances pour permettre un diagnostic rapide par l'Admin.

**FRs couverts :** FR25, FR26, FR27, FR28, FR29, FR30, FR31, FR37
**NFRs couverts :** NFR13 (logs pour diagnostic rapide)

### Story 5.1: Implémenter le module de logging

**En tant qu'** Admin,
**Je veux** que chaque action de l'application soit tracée dans un fichier de log,
**Afin de** pouvoir diagnostiquer rapidement tout problème.

**Module VBA :** `modLogging` (`src/modLogging.bas`)

> **Note :** Ce module est appelé par tous les autres modules. Il doit être implémenté en phase 1 (fondations).

**Critères d'Acceptation :**

**Étant donné** qu'une action est effectuée (connexion, chargement, filtrage, consolidation)
**Quand** la procédure `EnregistrerLog()` est appelée
**Alors** une ligne est ajoutée au fichier `tbAffaires.log` au format :
`DATE | USER | ACTION | RESULTAT`
Exemple : `2026-01-23 14:32:15 | Patrick | Consolidation 50 affaires | SUCCES (0.8 sec)`

**Étant donné** qu'une erreur survient
**Quand** le logging d'erreur est déclenché
**Alors** la ligne contient le contexte complet : date, trigramme utilisateur, action tentée, code d'erreur et message détaillé (FR29)

**Étant donné** que l'application supporte 3 niveaux de log
**Quand** une action est loggée
**Alors** le niveau est explicite : INFO (actions normales), ERREUR (échecs), SUCCES (opérations réussies) (FR30)

**Étant donné** que le fichier log n'existe pas au premier lancement
**Quand** la première écriture est tentée
**Alors** le fichier `tbAffaires.log` est créé automatiquement dans le répertoire `data/`

**Étant donné** qu'un utilisateur est en mode admin usurpé
**Quand** une action est loggée
**Alors** le champ USER contient "Patrick (pour HL)" au lieu de simplement "Patrick" (FR37)

**Étant donné** que le fichier de log est inaccessible (verrouillé, droits insuffisants)
**Quand** le module tente d'écrire
**Alors** l'erreur de logging est ignorée silencieusement (ne bloque pas l'application)

---

### Story 5.2: Implémenter le module de mesure de performance

**En tant que** développeur / Admin,
**Je veux** mesurer le temps d'exécution des opérations critiques,
**Afin de** valider les exigences de performance (< 5 secondes) et identifier les régressions.

**Module VBA :** `modTimer` (`src/modTimer.bas`)

**Critères d'Acceptation :**

**Étant donné** qu'une opération critique démarre (chargement ERP, chargement commentaires, consolidation UPSERT)
**Quand** `DemarrerTimer()` est appelé
**Alors** le temps de départ est capturé avec précision (via `Timer` VBA ou API Windows `GetTickCount`)

**Étant donné** que l'opération se termine
**Quand** `ArreterTimer()` est appelé
**Alors** la durée est calculée en secondes et retournée

**Étant donné** que la consolidation réussit en 0.8 secondes
**Quand** le résultat est affiché à l'utilisateur
**Alors** le message contient le temps : "Consolidation réussie - 50 affaires (0.8 sec)" (FR26)

**Étant donné** que le timer est utilisé pour une opération
**Quand** le résultat est loggé via modLogging
**Alors** le temps est inclus dans le champ RESULTAT : "SUCCES (0.8 sec)" (FR27)

**Étant donné** que la durée dépasse 5 secondes pour une opération critique
**Quand** le résultat est loggé
**Alors** un avertissement est enregistré : "WARNING - Opération lente (X.X sec)"

---

### Story 5.3: Permettre la consultation et le diagnostic par l'Admin

**En tant qu'** Admin,
**Je veux** consulter le fichier de logs pour diagnostiquer rapidement les problèmes signalés par les ADV,
**Afin de** résoudre les incidents sans perdre de temps.

**Module VBA :** Pas de module dédié (le fichier .log est consultable avec tout éditeur texte)

**Critères d'Acceptation :**

**Étant donné** qu'un ADV signale un problème (ex: "Ça ne veut pas consolider")
**Quand** j'ouvre le fichier `data/tbAffaires.log`
**Alors** les entrées sont lisibles, une par ligne, triées chronologiquement (plus récent en bas)

**Étant donné** que je cherche les erreurs récentes
**Quand** je recherche "ERREUR" dans le fichier
**Alors** je trouve le contexte complet : date, utilisateur, action, code d'erreur et message

**Étant donné** que l'erreur est identifiée (ex: ERR-101 Colonne mapping manquante)
**Quand** je consulte la table des codes d'erreur (dans architecture.md ou guide admin)
**Alors** l'action corrective Admin est documentée (ex: "Mettre à jour tbMapping dans data.xlsx")

**Étant donné** que je cherche les performances d'un ADV
**Quand** je filtre par trigramme dans les logs
**Alors** je vois les temps d'exécution de chaque opération et peux identifier les lenteurs

---

## Résumé Quantitatif

| Epic               | Stories | Modules VBA                                      | FRs        | NFRs        |
| ------------------ | ------- | ------------------------------------------------ | ---------- | ----------- |
| 1 - Infrastructure | 6       | clsOptimizer, modConfiguration, modUtils  | 11         | 8           |
| 2 - Chargement     | 5       | modExtraction, modCommentaires, modConsolidation | 6          | 2           |
| 3 - Filtrage       | 4       | modFiltrage                                      | 5          | 1           |
| 4 - Consolidation  | 4       | modConsolidation                                 | 7          | 4           |
| 5 - Logging        | 3       | modLogging, modTimer                             | 7          | 1           |
| **Total**          | **22**  | **9 modules**                                    | **36 FRs** | **16 NFRs** |
