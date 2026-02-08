---
stepsCompleted: [1, 2]
inputDocuments:
  - path: "_bmad-output/planning-artifacts/prd.md"
    type: "prd"
  - path: "_bmad-output/planning-artifacts/architecture.md"
    type: "architecture"
workflowType: "create-epics-and-stories"
project_name: "tbAffaires"
date: "2026-01-29"
author: "Patrick"
---

# tbAffaires - Epic Breakdown

## Overview

This document provides the complete epic and story breakdown for tbAffaires, decomposing the requirements from the PRD, UX Design if it exists, and Architecture requirements into implementable stories.

## Requirements Inventory

### Functional Requirements

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
- FR9: L'application charge les commentaires historiques depuis tbCommentaires dans data.xlsx
- FR10: L'application crée automatiquement le fichier de suivi s'il n'existe pas
- FR11: L'application affiche un message d'erreur si colonne mappée introuvable

**Filtrage et Affichage (FR12-FR16)**
- FR12: L'application filtre les affaires par trigramme ADV de l'utilisateur connecté (ou usurpé en mode Admin)
- FR13: L'application affiche les affaires dans un ListObject temporaire
- FR14: L'application met en évidence les affaires en difficulté financière (rouge)
- FR15: L'application pré-remplit les commentaires existants de S-1
- FR16: L'ADV navigue avec fonctionnalités Excel natives (filtres, tri, Ctrl+F)

**Saisie des Commentaires (FR17-FR19)**
- FR17: L'ADV saisit de nouveaux commentaires directement dans le ListObject (colonne Commentaire déverrouillée, reste du classeur verrouillé)
- FR18: L'ADV modifie les commentaires existants

**Consolidation (FR20-FR24)**
- FR20: L'ADV déclenche la consolidation de ses données
- FR21: L'application supprime les anciennes données ADV avant ajout (UPSERT)
- FR22: L'application réessaie automatiquement si fichier verrouillé (retry 0-3s, 5 max)
- FR23: L'application affiche message d'erreur après 5 échecs consolidation
- FR24: L'application préserve les données saisies même en cas d'échec

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

### NonFunctional Requirements

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
- NFR15: Permissions AD restreignent accès aux fichiers (ADV : data\ uniquement)

**Développement et Outils (NFR16-NFR19)**
- NFR16: Environnement Python obligatoire - Utiliser IMPÉRATIVEMENT pipenv (INTERDIT d'utiliser pip)
- NFR17: Pilotage Excel obligatoire - Utiliser OBLIGATOIREMENT le paquet pywin32 (INTERDIT d'utiliser openpyxl)
- NFR18: Localisation des scripts Python - Scripts Python OBLIGATOIREMENT enregistrés dans le répertoire scripts/
- NFR19: Automatisation via Python - Scripts pour création Excel, chargement VBA, tests automatisés

### Additional Requirements

**Règles Métier Immuables (Architecture)**
- UN SEUL utilisateur peut avoir IsAdmin = Oui dans tbADV (sinon ERREUR BLOQUANTE ERR-002)
- Chaque affaire appartient à UN SEUL ADV (plage exclusive) - pas de conflit de données possible
- Toutes les colonnes du mapping doivent être présentes dans l'extraction ERP avant tout traitement
- Le fichier d'extraction ERP repart à 0 affaires en début d'année (pas de problème de volume croissant)

**Structure des Données (Architecture)**
- Chaque ListObject DOIT être isolé dans sa propre feuille dans data.xlsx
- Feuille "ADV" → tbADV uniquement
- Feuille "Configuration" → tbParametres uniquement
- Feuille "Mapping" → tbMapping uniquement
- Feuille "Commentaires" → tbCommentaires uniquement

**Structure des Données tbADV**
- UserName | Nom | Prénom | Trigramme | IsAdmin

**Structure des Données tbParametres**
- Parametre | Valeur | Description
- CheminData, CheminExtraction, CheminConsolidation
- DelaiRetryMin (0), DelaiRetryMax (3), MaxTentatives (5)

**Structure des Données tbMapping**
- ColonneExtraction | ColonneSuivi | Type | Description
- 16 colonnes mappées (Année, Mois, ADV, Affaire, CA prévu/réel, etc.)

**Structure des Données tbCommentaires**
- NumeroAffaire | TrigrammeADV | Commentaire | DateModification

**Codes d'Erreur Standardisés (Architecture)**
- ERR-001: Utilisateur non configuré
- ERR-002: Double admin détecté
- ERR-101: Colonne mapping manquante
- ERR-102: Fichier extraction introuvable
- ERR-201: Fichier consolidation occupé
- ERR-202: Échec consolidation après 5 tentatives
- ERR-301: Commentaire trop long
- ERR-401: Mode Admin actif

**Logging Format (Architecture)**
- DATE | USER | ACTION | RESULTAT
- Exemple: 2026-01-23 14:32:15 | Patrick | Consolidation 50 affaires | SUCCES (0.8 sec)

**Naming Conventions (Architecture)**
- Modules VBA: Préfixe mod (ex: modConfiguration)
- Classes VBA: Préfixe cls (ex: clsApplicationState)
- Fonctions VBA: PascalCase français (Verbe+Nom) ex: ChargerDonneesExtraction()
- Constantes VBA: SCREAMING_SNAKE_CASE
- Fichiers horodatés: AAAAMMDD_HHMMSS

**Gestion Concurrence (Architecture)**
- UPSERT incrémental: Suppression ancien ADV + ajout nouveau
- Retry: Délai aléatoire 0-3s, max 5 tentatives
- Backup: Avant chaque consolidation dans data\backups\

**Workflow de Développement VBA (Architecture)**
- Code VBA source dans src/ (Git-friendly)
- Synchronisation via VBA Toolkit (Python + pywin32)
- Éditer src/ → Git commit → Import Excel → Tester

### FR Coverage Map

| FR   | Epic | Description courte                   |
| ---- | ---- | ------------------------------------ |
| FR1  | 1    | Initialisation RAII                  |
| FR2  | 1    | Identification Windows               |
| FR3  | 1    | Chargement config utilisateur        |
| FR4  | 1    | Erreur si non configuré              |
| FR5  | 1    | Restauration état Excel              |
| FR6  | 2    | Sélection fichier ERP                |
| FR7  | 2    | Chargement lecture seule             |
| FR8  | 2    | Chargement mapping                   |
| FR9  | 2    | Chargement commentaires historiques  |
| FR10 | 2    | Création auto fichier suivi          |
| FR11 | 2    | Erreur colonne manquante             |
| FR12 | 3    | Filtrage par trigramme               |
| FR13 | 3    | Affichage ListObject                 |
| FR14 | 3    | Mise en évidence difficultés         |
| FR15 | 3    | Pré-remplissage commentaires S-1     |
| FR16 | 3    | Navigation Excel native              |
| FR17 | 4    | Saisie commentaires                  |
| FR18 | 4    | Modification commentaires            |
| FR20 | 4    | Déclenchement consolidation          |
| FR21 | 4    | UPSERT suppression anciennes données |
| FR22 | 4    | Retry automatique                    |
| FR23 | 4    | Message erreur après 5 échecs        |
| FR24 | 4    | Préservation données en cas d'échec  |
| FR28 | 5    | Log des actions                      |
| FR29 | 5    | Log erreurs avec contexte            |
| FR30 | 5    | Distinction niveaux log              |
| FR31 | 5    | Consultation logs par Admin          |
| FR32 | 1    | Modification mapping sans code       |
| FR33 | 1    | Gestion utilisateurs data.xlsx       |
| FR34 | 1    | Configuration paramètres             |
| FR35 | 1    | Identification mode Admin            |
| FR36 | 1    | Usurpation utilisateur               |
| FR37 | 1    | Log spécifique mode Admin            |

## Epic List

### Epic 1: Infrastructure et Session Utilisateur
L'utilisateur peut démarrer l'application, être identifié automatiquement, et travailler dans un environnement Excel optimisé et sécurisé.

**FRs couverts:** FR1, FR2, FR3, FR4, FR5, FR32, FR33, FR34, FR35, FR36, FR37
**NFRs couverts:** NFR8 (RAII), NFR10-NFR13 (maintenabilité), NFR14-NFR15 (sécurité)

### Epic 2: Chargement et Préparation des Données
L'ADV peut charger son fichier d'extraction ERP et retrouver automatiquement ses commentaires historiques.

**FRs couverts:** FR6, FR7, FR8, FR9, FR10, FR11
**NFRs couverts:** NFR1, NFR2 (performance < 5s)

### Epic 3: Filtrage et Visualisation des Affaires
L'ADV visualise uniquement SES affaires avec les commentaires pré-remplis et les alertes visuelles.

**FRs couverts:** FR12, FR13, FR14, FR15, FR16
**NFRs couverts:** NFR4 (interface réactive)

### Epic 4: Saisie et Consolidation des Commentaires
L'ADV saisit ses commentaires et les consolide de manière fiable, même en cas de conflit d'accès.

**FRs couverts:** FR17, FR18, FR20, FR21, FR22, FR23, FR24
**NFRs couverts:** NFR3 (consolidation < 5s), NFR6, NFR7, NFR9 (fiabilité, retry)

### Epic 5: Logging et Observabilité
L'Admin peut consulter les logs pour diagnostiquer rapidement tout problème.

**FRs couverts:** FR28, FR29, FR30, FR31
**NFRs couverts:** NFR13 (logs pour diagnostic rapide)

---

## Epic 1: Infrastructure et Session Utilisateur

L'utilisateur peut démarrer l'application, être identifié automatiquement, et travailler dans un environnement Excel optimisé et sécurisé.

### Story 1.1: Créer la structure de fichiers et la configuration initiale

**En tant qu'** Admin,
**Je veux** créer la structure de fichiers et le fichier data.xlsx avec les tables de configuration,
**Afin de** pouvoir configurer l'application sans modifier le code VBA.

**Critères d'Acceptation :**

**Étant donné** que je suis sur le serveur AD
**Quand** je crée la structure dans `\\serveur-ad\FRV\AFFAIRES\01 SUIVI AFFAIRES\`
**Alors** les dossiers `data\`, `extractions\`, `backups\` existent

**Étant donné** que je crée le fichier data.xlsx
**Quand** j'ouvre le fichier
**Alors** il contient 4 feuilles (ADV, Configuration, Mapping, Commentaires) avec les ListObjects correspondants

**Étant donné** que je configure tbADV
**Quand** j'ajoute les utilisateurs
**Alors** la table contient : UserName, Nom, Prénom, Trigramme, IsAdmin

### Story 1.2: Implémenter le pattern RAII avec clsApplicationState

**En tant qu'** ADV,
**Je veux** que l'application optimise Excel au démarrage et restaure l'état à la fermeture,
**Afin de** bénéficier de performances maximales et éviter les problèmes d'état Excel.

**Critères d'Acceptation :**

**Étant donné** que je démarre tbAffaires.xlsm
**Quand** l'ApplicationState s'initialise
**Alors** ScreenUpdating, Calculation et Events sont désactivés

**Étant donné** que je ferme l'application (même en cas d'erreur)
**Quand** l'objet ApplicationState est détruit
**Alors** l'état initial d'Excel est restauré (ScreenUpdating, Calculation, Events)

**Étant donné** qu'une erreur survient pendant l'exécution
**Quand** l'erreur est interceptée
**Alors** l'état Excel est toujours restauré grâce au pattern RAII

### Story 1.3: Identifier automatiquement l'utilisateur Windows

**En tant qu'** ADV,
**Je veux** être identifié automatiquement via mon username Windows,
**Afin de** ne pas avoir à saisir mes identifiants.

**Critères d'Acceptation :**

**Étant donné** que j'ouvre tbAffaires.xlsm
**Quand** l'application démarre
**Alors** `Environ("USERNAME")` est récupéré automatiquement

**Étant donné** que mon username existe dans tbADV
**Quand** l'identification réussit
**Alors** mon trigramme ADV est chargé en mémoire

**Étant donné** que mon username n'existe pas dans tbADV
**Quand** l'identification échoue
**Alors** un message d'erreur ERR-001 s'affiche : "Utilisateur non configuré. Contacter Patrick."

### Story 1.4: Valider l'unicité de l'Admin

**En tant qu'** Admin,
**Je veux** que le système vérifie qu'il n'y a qu'un seul admin configuré,
**Afin d'** éviter les conflits de gestion.

**Critères d'Acceptation :**

**Étant donné** que le fichier data.xlsx est chargé
**Quand** l'application vérifie les admins
**Alors** elle compte les lignes avec IsAdmin = "Oui"

**Étant donné** qu'il y a exactement 1 admin
**Quand** la validation passe
**Alors** l'application continue normalement

**Étant donné** qu'il y a 0 ou 2+ admins
**Quand** la validation échoue
**Alors** une ERREUR BLOQUANTE ERR-002 s'affiche : "Double admin détecté. Contacter Patrick."

### Story 1.5: Implémenter le mode Admin avec usurpation

**En tant qu'** Admin,
**Je veux** pouvoir travailler au nom d'un autre ADV,
**Afin de** gérer les absences ou problèmes utilisateurs.

**Critères d'Acceptation :**

**Étant donné** que je suis identifié comme admin (IsAdmin = "Oui")
**Quand** l'application démarre
**Alors** une boîte de dialogue me propose de choisir un ADV à usurper

**Étant donné** que je choisis de travailler pour un autre ADV
**Quand** je sélectionne son trigramme
**Alors** l'application filtre sur ses affaires (pas les miennes)

**Étant donné** que je travaille en mode admin usurpé
**Quand** j'effectue une action
**Alors** le log indique "Action par [Admin] au nom de [Utilisateur usurpé]"

**Étant donné** que je suis en mode admin usurpé
**Quand** je consolide
**Alors** une alerte visuelle indique "Mode Admin actif - Consolidation pour [Trigramme]"

---

## Epic 2: Chargement et Préparation des Données

L'ADV peut charger son fichier d'extraction ERP et retrouver automatiquement ses commentaires historiques.

### Story 2.1: Sélectionner le fichier d'extraction ERP

**En tant qu'** ADV,
**Je veux** sélectionner mon fichier d'extraction ERP via une boîte de dialogue Windows,
**Afin de** charger mes données à traiter.

**Critères d'Acceptation :**

**Étant donné** que l'application est démarrée et que je suis identifié
**Quand** je clique sur "Charger extraction"
**Alors** une boîte de dialogue Windows s'ouvre avec le répertoire par défaut configuré dans tbParametres

**Étant donné** que la boîte de dialogue est ouverte
**Quand** je sélectionne un fichier .xlsx
**Alors** le chemin du fichier est mémorisé pour le chargement

**Étant donné** que je sélectionne un fichier
**Quand** le fichier est verrouillé par un autre processus
**Alors** un message d'erreur ERR-102 s'affiche : "Fichier extraction introuvable ou verrouillé. Vérifiez le chemin."

### Story 2.2: Charger et valider le mapping des colonnes

**En tant qu'** ADV,
**Je veux** que l'application charge le mapping des colonnes depuis data.xlsx,
**Afin de** pouvoir adapter l'application si les colonnes ERP changent sans modifier le code.

**Critères d'Acceptation :**

**Étant donné** que je déclenche le chargement de l'extraction
**Quand** l'application démarre le traitement
**Alors** elle charge d'abord tbMapping depuis data.xlsx

**Étant donné** que le mapping est chargé
**Quand** l'application vérifie les colonnes du fichier ERP
**Alors** elle valide que toutes les colonnes définies dans tbMapping existent dans l'extraction

**Étant donné** qu'une colonne du mapping est introuvable dans l'extraction ERP
**Quand** la validation échoue
**Alors** un message d'erreur ERR-101 s'affiche : "Colonne 'NomColonne' non trouvée. Vérifiez le mapping dans data.xlsx."

**Étant donné** que toutes les colonnes sont présentes
**Quand** la validation réussit
**Alors** l'application procède au chargement des données

### Story 2.3: Importer les données d'extraction en lecture seule

**En tant qu'** ADV,
**Je veux** que l'application charge mon fichier d'extraction ERP en lecture seule,
**Afin de** ne pas risquer de modifier les données source.

**Critères d'Acceptation :**

**Étant donné** que le mapping est validé
**Quand** l'application charge le fichier ERP
**Alors** elle l'ouvre en mode lecture seule (ReadOnly = True)

**Étant donné** que le fichier est chargé
**Quand** les données sont transférées
**Alors** l'application ferme le fichier ERP sans sauvegarder

**Étant donné** que le fichier contient ~800 affaires pour mon trigramme
**Quand** le chargement se termine
**Alors** le temps total est inférieur à 5 secondes (NFR1)

**Étant donné** que le fichier ERP est chargé
**Quand** les données sont importées
**Alors** elles sont stockées dans un ListObject temporaire en mémoire

### Story 2.4: Récupérer les commentaires historiques

**En tant qu'** ADV,
**Je veux** retrouver automatiquement mes commentaires de la semaine précédente,
**Afin de** ne pas avoir à les recopier manuellement (~100 commentaires).

**Critères d'Acceptation :**

**Étant donné** que les données ERP sont chargées
**Quand** l'application charge les commentaires historiques
**Alors** elle lit tbCommentaires dans data.xlsx

**Étant donné** que je suis identifié avec mon trigramme (ex: VC)
**Quand** les commentaires sont filtrés
**Alors** seuls les commentaires de mes affaires sont récupérés

**Étant donné** qu'une affaire existe dans l'extraction avec un commentaire historique
**Quand** les données sont fusionnées
**Alors** le commentaire historique est associé à l'affaire correspondante

**Étant donné** que le chargement des commentaires est lancé
**Quand** il se termine
**Alors** le temps est inférieur à 5 secondes (NFR2)

### Story 2.5: Créer automatiquement le fichier de suivi

**En tant qu'** ADV,
**Je veux** que l'application crée automatiquement le fichier de suivi s'il n'existe pas,
**Afin de** ne pas avoir à le créer manuellement la première fois.

**Critères d'Acceptation :**

**Étant donné** que je déclenche la première consolidation
**Quand** l'application vérifie l'existence du fichier de suivi
**Alors** elle cherche le fichier selon le chemin configuré dans tbParametres

**Étant donné** que le fichier de suivi n'existe pas
**Quand** la vérification échoue
**Alors** l'application crée un nouveau fichier basé sur consolidation.xltx

**Étant donné** que le fichier est créé
**Quand** il est ouvert
**Alors** il contient la structure standard avec les colonnes définies dans tbMapping

**Étant donné** que le fichier existe déjà
**Quand** la vérification réussit
**Alors** l'application utilise le fichier existant pour l'UPSERT

---

## Epic 3: Filtrage et Visualisation des Affaires

L'ADV visualise uniquement SES affaires avec les commentaires pré-remplis et les alertes visuelles.

### Story 3.1: Filtrer les affaires par trigramme ADV

**En tant qu'** ADV,
**Je veux** voir uniquement les affaires qui me sont assignées,
**Afin de** ne pas être distrait par les affaires des autres ADV.

**Critères d'Acceptation :**

**Étant donné** que les données ERP sont chargées
**Quand** l'application filtre les affaires
**Alors** seules les lignes avec mon trigramme ADV sont conservées

**Étant donné** que je suis en mode admin usurpé (ex: pour l'ADV "HL")
**Quand** le filtrage s'applique
**Alors** seules les affaires de "HL" sont affichées, pas les miennes

**Étant donné** que ~800 affaires correspondent à mon trigramme
**Quand** le filtrage est appliqué
**Alors** le temps de traitement est inférieur à 1 seconde (NFR4)

**Étant donné** qu'aucune affaire ne correspond à mon trigramme
**Quand** le filtrage retourne 0 résultats
**Alors** un message s'affiche : "Aucune affaire trouvée pour votre trigramme"

### Story 3.2: Créer et afficher le ListObject temporaire

**En tant qu'** ADV,
**Je veux** visualiser mes affaires dans un tableau Excel structuré (ListObject),
**Afin de** pouvoir utiliser les fonctionnalités natives d'Excel (filtres, tri, recherche).

**Critères d'Acceptation :**

**Étant donné** que les données sont filtrées par trigramme
**Quand** l'application crée l'affichage
**Alors** un ListObject temporaire est créé avec toutes les colonnes du mapping

**Étant donné** que le ListObject est créé
**Quand** je regarde l'affichage
**Alors** il contient : Numéro Affaire, Client, CA Prévu, CA Réel, Résultat Financier, Commentaire...

**Étant donné** que le classeur est affiché
**Quand** je consulte mes affaires
**Alors** toutes les cellules sont verrouillées SAUF la colonne "Commentaire"

**Étant donné** que le ListObject est affiché
**Quand** j'utilise Ctrl+F, les filtres ou le tri natifs Excel
**Alors** ces fonctionnalités fonctionnent normalement (FR16)

### Story 3.3: Mettre en évidence les affaires en difficulté financière

**En tant qu'** ADV,
**Je veux** identifier visuellement les affaires avec un résultat financier critique,
**Afin de** prioriser mon attention sur les problèmes urgents.

**Critères d'Acceptation :**

**Étant donné** que les affaires sont affichées dans le ListObject
**Quand** une affaire a un résultat financier négatif ou critique (selon seuil dans tbParametres)
**Alors** toute la ligne est mise en rouge (fond ou texte)

**Étant donné** que plusieurs affaires sont en difficulté
**Quand** je consulte le tableau
**Alors** elles apparaissent en haut du ListObject (tri automatique par priorité)

**Étant donné** qu'une affaire n'est pas en difficulté
**Quand** elle s'affiche
**Alors** elle conserve la mise en forme standard (pas de rouge)

**Étant donné** que le seuil de criticité est configurable dans tbParametres
**Quand** l'Admin modifie la valeur
**Alors** la mise en évidence s'adapte automatiquement sans changer le code

### Story 3.4: Pré-remplir les commentaires historiques

**En tant qu'** ADV,
**Je veux** retrouver mes commentaires de la semaine précédente déjà présents dans le tableau,
**Afin de** les modifier au lieu de les réécrire.

**Critères d'Acceptation :**

**Étant donné** que les commentaires historiques sont chargés depuis tbCommentaires
**Quand** le ListObject est créé
**Alors** la colonne "Commentaire" contient les valeurs historiques pour chaque affaire

**Étant donné** qu'une affaire n'a pas de commentaire historique
**Quand** elle s'affiche
**Alors** la cellule Commentaire est vide (pas d'erreur)

**Étant donné** qu'une affaire a un commentaire historique
**Quand** je consulte la ligne
**Alors** je peux modifier directement le commentaire dans la cellule déverrouillée

**Étant donné** que le commentaire est pré-rempli
**Quand** je le modifie
**Alors** la nouvelle valeur remplace l'ancienne en mémoire (pas encore sauvegardée)
