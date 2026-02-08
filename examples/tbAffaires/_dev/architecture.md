---
stepsCompleted: [1, 2, 3, 4, 5, 6]
inputDocuments:
  - path: '_bmad-output/planning-artifacts/product-brief-tbAffaires-2026-01-23.md'
    type: 'product-brief'
  - path: '_bmad-output/planning-artifacts/prd.md'
    type: 'prd'
workflowType: 'architecture'
project_name: 'tbAffaires'
date: '2026-01-23'
last_updated: '2026-01-29'
author: 'Patrick'
update_reason: 'Alignement avec PRD - Arbitrages Patrick intégrés + Documentation créée'
---

# Architecture Decision Document

> **📖 RÈGLES D'IMPLÉMENTATION** : Voir `project-context.md` à la racine du projet (LA BIBLE pour SM et Dev)

## Project Context

### Domaine et Complexité

- **Domaine** : Desktop App (VBA/Excel)
- **Complexité** : Faible-Moyenne
- **Composants** : 8-10 modules VBA

### Exigences Clés

**Fonctionnelles (FR1-FR34)** : Gestion session RAII, chargement données ERP, filtrage ADV, saisie commentaires, consolidation UPSERT, logging, configuration externe.

**Non-Fonctionnelles** :
- Performance : < 5 sec par opération, 800 affaires/ADV
- Fiabilité : 100% disponibilité vendredi, 0% perte données
- Maintenabilité : Code compréhensible par non-experts VBA

### Contraintes Techniques

| Contrainte | Valeur |
|------------|--------|
| Plateforme | Windows + Excel 2016+ |
| Infrastructure | Active Directory uniquement (pas de cloud) |
| Persistance | Fichiers Excel (pas de BDD) |
| Concurrence | Fichier unique partagé entre 3 ADV |
| Budget | Pas d'investissement ERP |
| Outil développement VBA | VBA Toolkit (Python + pywin32) dans scripts/ - **V2 uniquement** |

### Règles Métier Immuables

**RÈGLE 1 : Unicité de l'Admin (CRITIQUE)**
- UN SEUL utilisateur peut avoir `IsAdmin = Oui` dans tbADV
- Si deux admins détectés → ERREUR BLOQUANTE au démarrage (ERR-002)
- L'unicité est validée systématiquement par `modConfiguration`

**RÈGLE 2 : 1 Affaire = 1 ADV (CONCURRENCE)**
- Chaque affaire appartient à UN SEUL ADV (plage exclusive)
- Pas de conflit de données possible (seulement conflit de fichier)
- Si un ADV est absent n semaines (n > 1), ses affaires ne sont PAS mises à jour automatiquement
- L'admin doit consolider manuellement via Mode Admin pour les absences prolongées

**RÈGLE 3 : Validation Stricte du Mapping**
- Toutes les colonnes du mapping doivent être présentes dans l'extraction ERP
- Vérification AVANT tout traitement
- Message d'erreur clair si colonne manquante (ERR-101)

**RÈGLE 4 : Extraction Repart à Zéro Chaque Année**
- Le fichier d'extraction ERP repart à 0 affaires en début d'année
- Pas de problème de volume croissant à gérer
- Simplification de l'architecture

---

## Core Architectural Decisions

### Structure des Modules VBA

| Module | Responsabilité |
|--------|----------------|
| `clsApplicationState` | Classe RAII (gestion état Excel) - préfixe cls pour les classes |
| `modUtils` | Helpers, constantes, error handlers |
| `modConfiguration` | Chargement data.xlsx |
| `modLogging` | Logging (INFO, ERREUR, SUCCES) |
| `modTimer` | Mesure performance |
| `modExtraction` | Chargement fichier ERP |
| `modFiltrage` | Filtrage par trigramme ADV |
| `modConsolidation` | UPSERT + retry + sauvegardes |
| `modCommentaires` | Gestion historique commentaires (chargement/sauvegarde tbCommentaires) |

### Structure data.xlsx

**RÈGLE CRITIQUE** : Chaque ListObject **DOIT** être isolé dans sa propre feuille.

- Feuille "ADV" → **tbADV** (uniquement)
- Feuille "Configuration" → **tbParametres** (uniquement)
- Feuille "Mapping" → **tbMapping** (uniquement)
- Feuille "Commentaires" → **tbCommentaires** (uniquement) - *Historique centralisé des commentaires*

**tbADV** : `UserName | Nom | Prénom | Trigramme | IsAdmin`

**tbParametres** : `Parametre | Valeur | Description`
- CheminData, CheminExtraction, CheminConsolidation
- DelaiRetryMin (0), DelaiRetryMax (3), MaxTentatives (5)

**tbMapping** : `ColonneExtraction | ColonneSuivi | Type | Description`
- 16 colonnes mappées (Année, Mois, ADV, Affaire, CA prévu/réel, etc.)

**tbCommentaires** : `NumeroAffaire | TrigrammeADV | Commentaire | DateModification`
- Historique centralisé des commentaires (remplace commentaires_2026.xlsx)

### Authentification

- Identification : `Environ("USERNAME")` Windows
- Vérification : Lookup dans tbADV
- Permissions : AD restrictives (data\ uniquement)

### Gestion Concurrence

- **UPSERT incrémental** : Suppression ancien ADV + ajout nouveau
- **Retry** : Délai aléatoire 0-3s, max 5 tentatives
- **Backup** : Avant chaque consolidation dans `data\backups\`

### Format Logging

```
DATE | USER | ACTION | RESULTAT
2026-01-23 14:32:15 | Patrick | Consolidation 50 affaires | SUCCES (0.8 sec)
```

---

## Implementation Patterns

### Naming Conventions

| Élément | Convention | Exemple |
|---------|------------|---------|
| Modules VBA | Préfixe `mod` | `modConfiguration` |
| Fonctions VBA | PascalCase français (Verbe+Nom) | `ChargerDonneesExtraction()` |
| Constantes VBA | SCREAMING_SNAKE_CASE | `MAX_TENTATIVES` |
| Fichiers horodatés | AAAAMMDD_HHMMSS | `backup_20260123_143022.xlsx` |

### Error Handling

- **Format message** : "Erreur + Solution"
- **Centralisation** : Error handlers dans `modUtils`
- **Exemple** : `"Colonne Trigramme non trouvée. Vérifiez le mapping dans data.xlsx."`

### RAII Pattern (ApplicationState)

```vba
' Class_Initialize : Optimise (désactive ScreenUpdating, Calculation, Events)
' Class_Terminate : Restaure état initial (même en crash)
```

### Error Handling Strategy

**Principe : "Fail Fast, Fail Clear"**

Toutes les erreurs suivent le même format :
```
[TYPE ERREUR] : [Description courte]
[Explication contextuelle]
[SOLUTION]
[Contact]
```

**Codes d'Erreur Standardisés :**

| Code | Description | Action Utilisateur | Action Admin |
|------|-------------|-------------------|--------------|
| ERR-001 | Utilisateur non configuré | Contacter Patrick | Ajouter à tbADV |
| ERR-002 | Double admin détecté | Contacter Patrick | Corriger tbADV |
| ERR-101 | Colonne mapping manquante | Vérifier fichier | Mettre à jour tbMapping |
| ERR-102 | Fichier extraction introuvable | Vérifier chemin | Vérifier tbParametres |
| ERR-201 | Fichier consolidation occupé | Patienter/réessayer | Vérifier qui a le fichier ouvert |
| ERR-202 | Échec consolidation après 5 tentatives | Ne pas fermer, appeler Patrick | Vérifier verrou fichier |
| ERR-301 | Commentaire trop long | Raccourcir | - |
| ERR-401 | Mode Admin actif | Vérifier trigramme | Confirmer usurpation |

**Règles de Gestion :**
- Validation stricte du mapping avant chargement ERP (ERR-101)
- Validation unicité Admin au démarrage (ERR-002)
- Retry avec compteur visuel pour concurrence (ERR-201)
- Préservation des données saisies en cas d'échec
- Log systématique de toutes les erreurs

---

## Project Structure

```
\\serveur-ad\FRV\AFFAIRES\01 SUIVI AFFAIRES\
├── tbAffaires.xlsm                       # Application principale
├── data\
│   ├── data.xlsx                         # Config (tbADV, tbParametres, tbMapping, tbCommentaires)
│   ├── consolidation.xltx                # Modèle de fichier pour la direction
│   ├── backups｜                         # Sauvegardes horodatées (V2)
│   └── tbAffaires.log                    # Fichier de logs
├── extractions｜                          # Répertoire des fichiers ERP (paramétrable)
├── Suivi affaires 2026-S04.xlsx          # Consolidation semaine 04
├── ...                                   # Autres consolidations
└── Suivi affaires 2026-S52.xlsx          # Consolidation semaine 52

# Structure développement (hors production)
├── Pipfile                      # Dépendances Python
├── scripts/                     # Scripts Python (voir python-guidelines.md)
│   ├── vba_toolkit/             # API Python pour développement VBA
│   │   ├── __init__.py          # API publique
│   │   ├── excel_manager.py     # RAII pour piloter Excel
│   │   ├── vba_exporter.py      # Export VBA → fichiers
│   │   ├── vba_importer.py      # Import fichiers → VBA
│   │   ├── vba_sync.py          # Synchronisation bidirectionnelle
│   │   ├── vba_validator.py     # Validation cohérence
│   │   └── backup_manager.py    # Gestion des backups
│   ├── export_vba_modules.py    # Script export manuel
│   ├── import_vba_modules.py    # Script import manuel
│   └── tests/                   # Tests unitaires
└── src/                         # Code VBA source (Git-friendly)
    ├── clsApplicationState.cls
    ├── modUtils.bas
    ├── modConfiguration.bas
    ├── modLogging.bas
    ├── modTimer.bas
    ├── modExtraction.bas
    ├── modFiltrage.bas
    ├── modConsolidation.bas
    └── modCommentaires.bas
```

**Note:** Le code VBA est enregistré dans `src/` pour permettre le versioning Git et le refactoring. Le VBA Toolkit synchronise `src/` avec `tbAffaires.xlsm`.

### Python Guidelines

**Règles critiques** : Voir `project-context.md` (section Python)
**Détails complets** : Voir `docs/knowledge-base/guidelines/python-guidelines.md`

Contraintes clés :
- `pipenv` obligatoire (pas `pip`)
- `pywin32` obligatoire (pas `openpyxl`)
- Scripts Python dans le répertoire `scripts/`

---

## Data Flow

1. `modConfiguration` lit data.xlsx (tbADV, tbParametres, tbMapping, tbCommentaires)
2. `Environ("USERNAME")` → vérification tbADV
3. Boîte dialogue Windows → chargement extraction (lecture seule), s'ouvre sur le répertoire configuré dans tbParametres
4. `modFiltrage` → ListObject temporaire filtré par trigramme
5. `modCommentaires` → lecture commentaires historiques depuis tbCommentaires
6. ADV modifie ListObject (Excel natif)
7. `modConsolidation` → UPSERT dans fichier de suivi (racine du partage)
8. `modCommentaires` → sauvegarde commentaires mis à jour dans tbCommentaires
9. `modLogging` → append tbAffaires.log

---

## VBA Development Workflow

Le workflow de développement VBA utilise le VBA Toolkit pour synchroniser le code entre les fichiers source (`src/`) et le classeur Excel (`tbAffaires.xlsm`).

### Structure VBA Source

Le code VBA est enregistré dans `src/` sous forme de fichiers texte :

```
src/
├── clsApplicationState.cls   # Classe RAII (préfixe cls pour les classes)
├── modUtils.bas              # Helpers, constantes, gestion erreurs
├── modConfiguration.bas      # Chargement configuration
├── modLogging.bas            # Logging
├── modTimer.bas              # Mesure performance
├── modExtraction.bas         # Import ERP
├── modFiltrage.bas           # Filtrage ADV
├── modConsolidation.bas      # UPSERT + retry
└── modCommentaires.bas       # Gestion historique commentaires
```

### Workflow Développeur

```
┌─────────────────────────────────────────────────────────────┐
│                     DÉVELOPPEMENT VBA                        │
└─────────────────────────────────────────────────────────────┘

1. ÉDITION DU CODE
   ├── Éditer fichiers dans src/ (IDE texte, Git...)
   ├── Refactoriser, formater, documenter
   └── Git commit/pull/push (manuel)

2. IMPORT DANS EXCEL
   ├── Script: python scripts/import_vba_modules.py
   ├── VBA Toolkit: VBAImporter.import_all("src/")
   ├── Backup automatique avant import
   └── tbAffaires.xlsm mis à jour

3. TESTS DANS EXCEL
   ├── Ouvrir tbAffaires.xlsm
   ├── Tester fonctionnalités (manuels ou automatisés)
   └── Debug VBA si nécessaire

4. [Optionnel] EXPORT POUR SAUVEGARDER
   ├── Script: python scripts/export_vba_modules.py
   └── VBA Toolkit: VBAExporter.export_all("src/")
```

### API VBA Toolkit

```python
from vba_toolkit import VBAExporter, VBAImporter, VBASync

# Exporter tous les modules VBA du classeur
with VBAExporter("tbAffaires.xlsm") as exporter:
    modules = exporter.export_all("src/")
    print(f"{len(modules)} modules exportés")

# Importer les modules depuis src/ vers classeur
with VBAImporter("tbAffaires.xlsm") as importer:
    importer.import_all("src/")
    print("Modules importés avec succès")

# Synchroniser bidirectionnellement
with VBASync("tbAffaires.xlsm", "src/") as sync:
    report = sync.compare()
    if report.has_conflicts:
        sync.resolve_conflicts()
    sync.apply_changes()
```

### Avantages du VBA Toolkit

| Avantage | Description |
|----------|-------------|
| **Versioning Git** | Code VBA versionnable dans src/ |
| **Refactoring** | Refactoriser dans IDE texte moderne |
| **Travail équipe** | Git merge/pull sur fichiers VBA |
| **Backup auto** | Snapshots avant chaque import |
| **Productivité** | Import/Export rapide et fiable |
| **Validation** | Vérification cohérence automatique |

### Scénarios d'Utilisation

**Scénario 1: Nouvelle fonctionnalité**
```python
# 1. Éditer src/modExtraction.bas (nouvelle fonction)
# 2. Git commit
# 3. Importer pour tester
from vba_toolkit import VBAImporter
with VBAImporter("tbAffaires.xlsm") as importer:
    importer.import_module("src/modExtraction.bas")
```

**Scénario 2: Résolution de conflits Git**
```python
# 1. Git merge sur src/modUtils.bas
# 2. Résoudre conflits dans IDE
# 3. Importer version résolue
from vba_toolkit import VBAImporter
with VBAImporter("tbAffaires.xlsm") as importer:
    importer.import_module("src/modUtils.bas")
```

**Scénario 3: Comparaison versions**
```python
# Comparer classeur vs src/
from vba_toolkit import VBASync
with VBASync("tbAffaires.xlsm", "src/") as sync:
    report = sync.compare()
    for diff in report.differences:
        print(f"{diff.module}: {diff.status}")
```

**Référence:** Voir ADR-005 pour les détails complets du VBA Toolkit

---

## Requirements Mapping (FR → Modules)

| Catégorie | FR | Modules |
|-----------|-----|---------|
| Session | FR1-FR5 | clsApplicationState, modConfiguration, modUtils |
| Données | FR6-FR11 | modExtraction, modConfiguration |
| Filtrage | FR12-FR16 | modFiltrage |
| Saisie | FR17-FR19 | (Excel natif) |
| Consolidation | FR20-FR24 | modConsolidation, modUtils, modTimer |
| Timer | FR25-FR27 | modTimer |
| Logging | FR28-FR31 | modLogging |
| Config | FR32-FR34 | modConfiguration (via data.xlsx) |
| Commentaires | FR9, FR15 | modCommentaires (tbCommentaires dans data.xlsx) |

---

## Documentation Associée

### Guides Utilisateurs

| Document | Public | Contenu |
|----------|--------|---------|
| `docs/guide-utilisateur.md` | ADV (3 utilisateurs) | Procédure 5 étapes, problèmes courants, mode admin |
| `docs/guide-administrateur.md` | Patrick (Admin) | Configuration data.xlsx, points de vigilance, procédures d'urgence |
| `docs/points-vigilance-et-erreurs.md` | Dev + Admin | Matrice des risques, codes erreur, stratégie de gestion d'erreurs |

### Documentation Technique

| Document | Contenu |
|----------|---------|
| `docs/knowledge-base/guidelines/vba-guidelines.md` | Conventions de code VBA (Windows-1252, naming, structure) |
| `docs/knowledge-base/decisions/001-vba-toolkit.md` | Architecture du VBA Toolkit (post-développement) |
| `project-context.md` (racine) | Règles d'implémentation pour SM et Dev |

---

## Development Sequence

### Phase 1: Infrastructure (Étape 0)

1. Structure fichiers (tbAffaires.xlsm + data.xlsx + src/)
2. **VBA Toolkit** (scripts/vba_toolkit/) - *V2*
   - ExcelManager (RAII pour Excel)
   - VBAExporter, VBAImporter, VBASync, VBAValidator
   - BackupManager (sauvegardes horodatées)
   - Scripts utilitaires (import/export)
   - Tests unitaires
3. **Configuration Git** (.gitignore pour *.xlsm, src/ inclus)

### Phase 2: Modules VBA (Étapes 1-9)

4. clsApplicationState (RAII - Classe)
5. modUtils (fondation)
6. modConfiguration
7. modLogging
8. modTimer
9. modExtraction
10. modFiltrage
11. modConsolidation
12. modCommentaires (gestion historique commentaires)

**Note:** Chaque module est développé dans `src/` puis importé dans `tbAffaires.xlsm` via VBA Toolkit.

### Phase 3: Tests & Documentation (Étapes 10-12)

13. Tests manuels (5 scénarios)
14. Documentation (guide utilisateur, gestionnaire, FAQ)
15. Documentation développeur (VBA Toolkit usage)

### Phase 4: Déploiement (Étape 13)

16. Déploiement serveur AD

**Workflow cyclique:**
```
Éditer src/ → Git commit → Import Excel → Tester → [Modifier src/] → Répéter
```
