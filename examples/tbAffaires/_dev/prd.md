---
stepsCompleted: [1, 2, 3, 4, 5-skipped, 6-skipped, 7, 8, 9, 10, 11]
inputDocuments:
  - path: "_bmad-output/planning-artifacts/product-brief-tbAffaires-2026-01-23.md"
    type: "product-brief"
    loaded: true
workflowType: "prd"
documentCounts:
  briefs: 1
  research: 0
  brainstorming: 0
  projectDocs: 0
projectType: "greenfield"
classification:
  projectType: "Desktop App (VBA/Excel)"
  domain: "General - Business Process Automation"
  complexity: "Low-Medium"
  projectContext: "greenfield"
date: 2026-01-23
author: Patrick
---

# Product Requirements Document - tbAffaires

**Author:** Patrick
**Date:** 2026-01-23
**Version:** 1.0

---

## Executive Summary

### Vision

tbAffaires est une solution VBA/Excel qui transforme le processus de reporting hebdomadaire des ADV (Responsables d'Affaires) : de 20 minutes de travail séquentiel fastidieux à moins de 10 minutes de travail parallèle fluide.

### Problème

3 ADV doivent produire chaque vendredi un rapport sur ~2500 affaires. Actuellement :

- Travail **séquentiel** (verrouillage Excel)
- **Recopie manuelle** de ~100 commentaires historiques par ADV
- Direction reçoit le fichier **lundi matin** au lieu de vendredi soir

### Solution

Une macro VBA sur serveur Active Directory permettant :

- **Travail parallèle** (1 affaire = 1 ADV = pas de conflit)
- **Récupération automatique** des commentaires historiques
- **Consolidation incrémentale** (UPSERT) avec gestion des conflits
- **Format conforme** au modèle.xltx de la direction

### Différenciateur

Architecture **data-driven** : tout est configurable sans toucher au code (mapping colonnes, utilisateurs, paramètres). Connaissance métier profonde intégrée par l'équipe qui vit le problème quotidiennement.

### Impact Mesurable

| Métrique        | Avant       | Après                  |
| --------------- | ----------- | ---------------------- |
| Temps/ADV       | 20 min      | < 10 min (-50%)        |
| Erreurs copie   | Fréquentes  | 0%                     |
| Délai direction | Lundi matin | Vendredi soir          |
| ROI annuel      | -           | ~52 heures économisées |

---

## Classification Projet

| Critère    | Valeur                                |
| ---------- | ------------------------------------- |
| Type       | Desktop App (VBA/Excel)               |
| Domaine    | General - Business Process Automation |
| Complexité | Low-Medium                            |
| Contexte   | Greenfield                            |
| Plateforme | Windows uniquement                    |
| Connexion  | Online uniquement (serveur AD requis) |

---

## Critères de Succès

### Succès Utilisateur

| Critère          | Mesure                                    | Objectif            |
| ---------------- | ----------------------------------------- | ------------------- |
| Temps de travail | Mesure manuelle ou perception utilisateur | < 10 min/semaine    |
| Satisfaction     | Rétro informelle mensuelle avec l'admin   | 3/3 ADV "Satisfait" |
| Adoption         | Usage effectif chaque vendredi            | 100% des ADV        |

### Succès Business

| Critère         | Mesure                               | Objectif                 |
| --------------- | ------------------------------------ | ------------------------ |
| Délai direction | Heure de réception fichier consolidé | Vendredi soir (vs lundi) |
| Qualité données | Logs + feedback direction            | 0% commentaires perdus   |
| ROI             | Temps économisé × 52 semaines        | ~52 heures/an            |

### Succès Technique

| Critère       | Mesure                                    | Objectif             |
| ------------- | ----------------------------------------- | -------------------- |
| Chargement    | Temps ouverture extraction + commentaires | < 5 secondes         |
| Consolidation | Temps sauvegarde UPSERT                   | < 5 secondes         |
| Stabilité     | Bugs critiques/semaine                    | 0                    |
| Traçabilité   | Logging des actions                       | 100% actions loggées |

---

## Scope Produit

### MVP - Minimum Viable Product

**Fonctionnalités essentielles :**

- Chargement extraction ERP via boîte de dialogue Windows
- Identification automatique utilisateur (username système)
- Mode Admin pour usurpation d'utilisateur (V1)
- Filtrage automatique par trigramme ADV
- Récupération commentaires depuis le fichier consolidé de la semaine précédente (optionnel au premier lancement)
- Saisie dans ListObject (colonne Commentaire déverrouillée, reste verrouillé)
- Mise en évidence affaires en difficulté (résultat financier critique)
- Consolidation incrémentale UPSERT avec retry (0-3s, 5 tentatives)
- Logging détaillé (qui, quand, quoi, résultat)
- RAII (ApplicationState) pour stabilité
- Messages d'erreur clairs (cause + action suggérée)
- Mapping colonnes flexible via data.xlsx

**Livrables MVP :**

- `tbAffaires.xlsm` - Application principale
- `data.xlsx` - Configuration (tbADV, tbParametres, tbMapping)
- Guide utilisateur ADV (1 page)

**Estimation :** 45-65 heures → 5-7 semaines à 10h/semaine

### Post-MVP (V2)

| Fonctionnalité         | Effort | Déclencheur              |
| ---------------------- | ------ | ------------------------ |
| Sauvegardes horodatées | 3-4h   | Premier incident données |
| Guide gestionnaire     | 2-3h   | Questions récurrentes    |

### Vision Future

- Notification automatique assistant projet (email/Teams)
- Dashboard de suivi des métriques (temps/ADV, adoption)
- Historique multi-semaines consultable
- Export automatique vers direction

---

## Parcours Utilisateur

### Parcours 1 : Vincent, ADV - Happy Path

**Contexte :** Vendredi 16h00, Vincent gère ~800 affaires, pressé de finir sa semaine.

**Déroulement :**

1. Double-clic sur `tbAffaires.xlsm` → identification automatique (username → trigramme VC)
2. Sélection du fichier consolidé précédent (optionnel, clic Annuler si première semaine), puis sélection du fichier d'extraction ERP via boîte de dialogue → chargement < 5s
3. Affichage ListObject filtré sur SES affaires uniquement (classeur verrouillé sauf colonne Commentaire)
4. Affaires en difficulté en rouge, commentaires historiques pré-remplis depuis le fichier consolidé de S-1
5. Saisie des nouveaux commentaires dans la colonne déverrouillée
6. Clic "Consolider" → **"Consolidation réussie"**
7. Notification orale à l'assistant projet → départ week-end

**Capacités révélées :** FR1-FR24, FR28-FR37

---

### Parcours 2 : Vincent, ADV - Fichier Verrouillé

**Contexte :** Vincent et Hélène consolident en même temps.

**Déroulement :**

1. Vincent clique "Consolider" → **"Fichier occupé. Tentative 1/5..."**
2. Retry automatique (délai aléatoire 0-3s)
3. Tentative 3 : Hélène a terminé → consolidation réussit
4. Si 5 échecs : message d'erreur clair, données saisies préservées

**Capacités révélées :** FR22-FR24

---

### Parcours 3 : Patrick, Admin - Diagnostic Temps Réel

**Contexte :** Vendredi 16h20, Najoi a un message d'erreur.

**Déroulement :**

1. Najoi vient voir Patrick : "Ça ne veut pas consolider"
2. Patrick voit à l'écran : **"Colonne 'Trigramme_ADV' non trouvée. Vérifiez le mapping."**
3. Patrick ouvre les logs → `ERREUR | Colonne manquante: Trigramme_ADV`
4. Patrick ouvre `data.xlsx` → met à jour tbMapping (colonne renommée par l'ERP)
5. Najoi relance → chargement OK → consolidation réussie à 16h35

**Capacités révélées :** FR11, FR28-FR32

---

### Parcours 4 : Le Directeur - Analyse Vendredi Soir

**Contexte :** Vendredi 17h15, réception fichier consolidé.

**Déroulement :**

1. Email de l'assistant projet avec fichier "Suivi affaires 2026-S04.xlsx"
2. Ouverture → format modèle.xltx avec mises en forme conditionnelles
3. 35 affaires rouges (difficulté) sur 2400 totales, triées en premier
4. Lecture commentaires → appel ADV pour clarification si nécessaire
5. Décisions et email de cadrage envoyés avant le week-end

**Contraintes sur le livrable :** Format modèle.xltx, mise en forme préservée, commentaires exploitables.

---

### Résumé Capacités par Parcours

| Parcours         | Capacités Clés                                                                    |
| ---------------- | --------------------------------------------------------------------------------- |
| ADV Happy Path   | Identification auto, chargement < 5s, filtrage, récup commentaires, timer, UPSERT |
| ADV Erreur       | Retry automatique, messages clairs, logging, préservation données                 |
| Admin Diagnostic | Logging détaillé, mapping modifiable, architecture data-driven                    |
| Direction        | Format modèle.xltx, mise en forme, tri, commentaires exploitables                 |

---

## Exigences Techniques Desktop App

### Support Plateforme

- **OS :** Windows uniquement
- **Excel :** 2016+ (compatible ListObjects et VBA)
- **Réseau :** Accès serveur Active Directory obligatoire
- **Permissions :** Lecture/écriture sur répertoire `data\`

### Structure Fichiers Serveur AD

```
\\serveur-ad\FRV\AFFAIRES\01 SUIVI AFFAIRES\
├── tbAffaires.xlsm                       # Application
└── data\
|   ├── data.xlsx                         # Config (tbADV, tbParametres, tbMapping)
|   ├── modèle.xltx                       # Modèle de fichier de consolidation à transmettre à la direction
|   └── tbAffaires.log                    # Fichier de logs
└── extractions\                          # Répertoires où se trouvent les fichiers d'extraction de l'ERP
|   ├── extraction1.xlsx
|   ├── extraction2.xlsx
|   ├── ...
|   └── extractionN.xlsx
|── Suivi affaires 2026-S04.xlsx          # Consolidation 2026 semaine 04 (contient les commentaires)
|── ...                                   # Ensemble des consolidations de l'année
└── Suivi affaires 2026-S52.xlsx          # Consolidation 2026 semaine 52
```

### Intégrations Système

| Intégration       | Méthode                  | Usage                      |
| ----------------- | ------------------------ | -------------------------- |
| Username Windows  | `Environ("USERNAME")`    | Identification automatique |
| Fichiers réseau   | Chemins UNC `\\serveur\` | Data, commentaires, logs   |
| Sélection fichier | `GetOpenFilename`        | Choix fichier consolidé (opt.) + extraction ERP |

### Stratégie de Déploiement

1. Admin modifie `tbAffaires.xlsm` sur son poste
2. Copie sur serveur AD (emplacement central)
3. ADV ouvrent toujours depuis emplacement réseau
4. Pas de copie locale → toujours dernière version

---

## Exigences Fonctionnelles

### Gestion de Session (FR1-FR5)

- **FR1:** L'application s'initialise avec optimisation performances Excel (RAII)
- **FR2:** L'application identifie automatiquement l'utilisateur via username Windows
- **FR3:** L'application charge la configuration utilisateur depuis data.xlsx (tbADV)
- **FR4:** L'application affiche un message d'erreur si utilisateur non configuré
- **FR5:** L'application restaure l'état Excel à la fermeture (même en cas d'erreur)

### Chargement des Données (FR6-FR11)

- **FR6:** L'ADV sélectionne d'abord le fichier consolidé précédent (optionnel), puis le fichier d'extraction ERP via boîte de dialogue Windows
- **FR7:** L'application charge le fichier d'extraction en lecture seule
- **FR8:** L'application charge le mapping des colonnes depuis data.xlsx (tbMapping)
- **FR9:** L'application charge les commentaires historiques depuis le fichier consolidé de la semaine précédente (colonne Commentaire)
- **FR10:** L'application crée automatiquement le fichier de suivi s'il n'existe pas
- **FR11:** L'application affiche un message d'erreur si colonne mappée introuvable

### Filtrage et Affichage (FR12-FR16)

- **FR12:** L'application filtre les affaires par trigramme ADV de l'utilisateur connecté (ou usurpé en mode Admin)
- **FR13:** L'application affiche les affaires dans un ListObject temporaire
- **FR14:** L'application met en évidence les affaires en difficulté financière (rouge)
- **FR15:** L'application pré-remplit les commentaires existants de S-1
- **FR16:** L'ADV navigue avec fonctionnalités Excel natives (filtres, tri, Ctrl+F)

### Saisie des Commentaires (FR17-FR19)

- **FR17:** L'ADV saisit de nouveaux commentaires directement dans le ListObject (colonne Commentaire déverrouillée, reste du classeur verrouillé)
- **FR18:** L'ADV modifie les commentaires existants

### Consolidation (FR20-FR24)

- **FR20:** L'ADV déclenche la consolidation de ses données
- **FR21:** L'application supprime les anciennes données ADV avant ajout (UPSERT)
- **FR22:** L'application réessaie automatiquement si fichier verrouillé (retry 0-3s, 5 max)
- **FR23:** L'application affiche message d'erreur après 5 échecs consolidation
- **FR24:** L'application préserve les données saisies même en cas d'échec

### Logging et Traçabilité (FR28-FR31)

- **FR28:** L'application enregistre chaque action dans un fichier de log
- **FR29:** L'application enregistre les erreurs avec contexte (qui, quand, quoi)
- **FR30:** L'application distingue les niveaux de log (INFO, ERREUR, SUCCES)
- **FR31:** L'Admin consulte le fichier de logs pour diagnostiquer les problèmes

### Configuration et Administration (FR32-FR34)

- **FR32:** L'Admin modifie le mapping colonnes sans toucher au code VBA
- **FR33:** L'Admin ajoute/modifie des utilisateurs dans data.xlsx (tbADV)
- **FR34:** L'Admin configure les paramètres dans data.xlsx (tbParametres)

### Mode Admin (FR35-FR37)

- **FR35:** L'application identifie les utilisateurs admin via la colonne `IsAdmin` dans tbADV
- **FR36:** L'Admin peut choisir de travailler au nom d'un autre ADV via une boîte de dialogue
- **FR37:** Le logging indique "Action par [Admin] au nom de [Utilisateur usurpé]"

---

## Exigences Non-Fonctionnelles

### Performance (NFR1-NFR4)

| NFR  | Exigence                              | Mesure                |
| ---- | ------------------------------------- | --------------------- |
| NFR1 | Chargement extraction                 | < 5 secondes          |
| NFR2 | Chargement commentaires               | < 5 secondes          |
| NFR3 | Consolidation UPSERT                  | < 5 secondes          |
| NFR4 | Interface réactive pendant opérations | Pas de freeze > 1 sec |

### Fiabilité (NFR6-NFR9)

| NFR  | Exigence                                        | Mesure                       |
| ---- | ----------------------------------------------- | ---------------------------- |
| NFR6 | Fonctionnement chaque vendredi sans échec       | 100% disponibilité hebdo     |
| NFR7 | Données saisies jamais perdues                  | 0% perte de données          |
| NFR8 | État Excel restauré même en cas de crash (RAII) | Restauration automatique     |
| NFR9 | Gestion conflits verrouillage fichier           | 5 tentatives max, délai 0-3s |

### Maintenabilité (NFR10-NFR13)

| NFR   | Exigence                                    | Mesure                                |
| ----- | ------------------------------------------- | ------------------------------------- |
| NFR10 | Code compréhensible par non-expert VBA      | Fonctions nommées explicitement       |
| NFR11 | Modifications mapping sans toucher au code  | 100% via data.xlsx                    |
| NFR12 | Messages d'erreur indiquent cause et action | Format : "Erreur + Solution"          |
| NFR13 | Logs permettent diagnostic rapide           | Format : Date, User, Action, Résultat |

### Sécurité (NFR14-NFR15)

| NFR   | Exigence                                       | Mesure                          |
| ----- | ---------------------------------------------- | ------------------------------- |
| NFR14 | Seuls utilisateurs configurés peuvent utiliser | Vérification tbADV au démarrage |
| NFR15 | Permissions AD restreignent accès aux fichiers | ADV : data\ uniquement          |

### Développement et Outils (NFR16-NFR19)

| NFR   | Exigence                         | Contrainte                                                                |
| ----- | -------------------------------- | ------------------------------------------------------------------------- |
| NFR16 | Environnement Python obligatoire | Utiliser IMPÉRATIVEMENT pipenv (INTERDIT d'utiliser pip)                  |
| NFR17 | Pilotage Excel obligatoire       | Utiliser OBLIGATOIREMENT le paquet pywin32 (INTERDIT d'utiliser openpyxl) |
| NFR18 | Localisation des scripts Python  | Scripts Python OBLIGATOIREMENT enregistrés dans le répertoire scripts/    |
| NFR19 | Automatisation via Python        | Scripts pour création Excel, chargement VBA, tests automatisés            |

---

## Contraintes de Développement Python

### Règles IMPÉRATIVES

⚠️ **RÈGLES STRICTES À RESPECTER :**

🚫 **INTERDIT :**

- Utiliser Python dans l’application finale
- Utiliser `pip` directement pour installer des dépendances Python
- Utiliser le paquet `openpyxl` pour manipuler des fichiers Excel
- Enregistrer des scripts Python en dehors du répertoire `scripts/`
- Insérer des émojis dans les chaînes de caractères

✅ **OBLIGATOIRE :**

- Utiliser Python comme outils de développement
- Utiliser `pipenv` pour l'environnement virtuel Python
- Utiliser le paquet `pywin32` pour piloter Excel via COM
- Enregistrer tous les scripts Python dans le répertoire `scripts/` à la racine du projet
- Utiliser des caractères textes UNIQUEMENT dans les chaînes de caractères

### Rationnel technique

1. **pipenv obligatoire** : Assure l'isolement des dépendances et la reproductibilité de l'environnement
2. **pywin32 obligatoire** : Contrôle natif d'Excel via COM, compatible avec les fichiers xlsm et les macro VBA
3. **openpyxl interdit** : Ne peut pas manipuler les macro VBA et ne fournit pas les mêmes fonctionnalités COM
4. **scripts/ obligatoire** : Centralisation et structure claire pour la maintenance

### Cas d'utilisation des scripts Python

Les scripts Python sont utilisés **UNIQUEMENT durant la phase de développement** pour automatiser :

1. **Création de fichiers Excel**
   - Génération de fichiers `data.xlsx` avec les ListObjects tbADV, tbParametres, tbMapping (3 feuilles)
   - Mise à jour automatique des structures de tables et des données dans les fichiers excel

2. **Gestion des modules VBA**
   - CRUD de modules VBA dans les fichiers Excel
   - Import de fichiers (\*.bas) de modules VBA dans les fichiers Excel
   - Import de fichiers (\*.cls) de modules de classe dans les fichiers Excel
   - Import de fichiers (\*.frm/frx) de UserForms dans les fichiers Excel
   - Automatisation du déploiement des mises à jour VBA

3. **Tests automatisés**
   - Exécution de tests unitaires sur les fonctions VBA via COM
   - Tests d'intégration du flux de données complet
   - Tests de performance (chargement < 5 secondes, consolidation < 5 secondes)
   - Tests de gestion des erreurs et retry
   - Validation de la structure des ListObjects

### Workflow de développement Python

```bash
# 1. Installation de l'environnement virtuel (une seule fois)
pipenv install

# 2. Activation de l'environnement virtuel
pipenv shell

# 3. Ajout d'une dépendance (jamais pip install !)
pipenv install <paquet>

# 4. Exécution d'un script
pipenv python  scripts/nom_du_script.py

# 5. Lancement des tests automatisés
pipenv python -m pytest tests/
```

### Exemple de script Python correct

```python
# scripts/create_data_xlsx.py
import win32com.client as win32
import os

def create_excel_file(filepath):
    """
    Crée un fichier Excel avec pywin32 (OBLIGATOIRE)
    """
    # Ouvrir Excel via pywin32
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Créer un nouveau classeur
        workbook = excel.Workbooks.Add()

        # Créer les ListObjects
        worksheet = workbook.Worksheets(1)
        list_objects = worksheet.ListObjects.Add(1, worksheet.Range("A1:D1"))
        list_objects.Name = "tbADV"

        # Sauvegarder
        workbook.SaveAs(os.path.abspath(filepath))
        print(f"[OK] Fichier créé : {filepath}")

    except Exception as e:
        print(f"[FAIL] Erreur : {e}")
        raise

    finally:
        # Fermer Excel proprement (RAII pattern)
        workbook.Close(False)
        excel.Quit()
        print("[OK] Excel fermé proprement")

if __name__ == "__main__":
    create_excel_file("data/data.xlsx")
```

---

## Analyse des Risques

### Risques Techniques

| Risque                              | Probabilité | Impact | Mitigation                         |
| ----------------------------------- | ----------- | ------ | ---------------------------------- |
| Conflit verrouillage fichier        | Moyenne     | Moyen  | Retry aléatoire 0-3s, 5 tentatives |
| Performance dégradée (800 affaires) | Faible      | Moyen  | RAII + optimisations Excel         |
| Format extraction ERP change        | Moyenne     | Faible | Mapping colonnes flexible          |

### Risques Projet

| Risque                  | Probabilité | Impact | Mitigation                                |
| ----------------------- | ----------- | ------ | ----------------------------------------- |
| Temps disponible réduit | Moyenne     | Moyen  | Pas de deadline, progression incrémentale |
| Blocage technique VBA   | Faible      | Élevé  | Architecture simple, patterns éprouvés    |
| Adoption ADV difficile  | Faible      | Moyen  | Guide 1 page, période transition          |

### Plan de Contingence

- **Blocage majeur :** Rollback vers processus manuel
- **Temps réduit :** MVP reste viable, V2 reportée
- **ADV absent (V1) :** Processus manuel temporaire jusqu'à V2

---

## Annexe : Ressources Projet

### Équipe

| Rôle         | Personne | Disponibilité    |
| ------------ | -------- | ---------------- |
| Développeur  | Patrick  | 10h/semaine      |
| Admin        | Patrick  | Support continu  |
| Utilisateurs | 3 ADV    | Vendredi 16h-17h |

### Documents Liés

- Product Brief : `product-brief-tbAffaires-2026-01-23.md`
- Modèle direction : `modèle.xltx` (à obtenir)

---

**Document généré le :** 2026-01-29
**Workflow :** PRD Create Mode
**Status :** ✅ Complet
