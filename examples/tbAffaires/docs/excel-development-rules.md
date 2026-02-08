# Règles VBA Excel

## Encodage & Chemins

### OBLIGATIONS

- **OBL-001** : Fichiers .bas/.cls/.frm en **Windows-1252** + CRLF (`\r\n`) - transcrire à partir d'UTF-8 si nécessaire.
- **OBL-002** : Chemins POSIX `/` partout (portable Win/Mac).

### INTERDICTIONS

- **INT-001** : Emojis code/strings.
- **INT-002** : `Option Explicit` absent.

## Naming & Structure

### OBLIGATIONS

- **OBL-003** : Français PascalCase.
- **OBL-004** : `modXXX.bas` / `clsYYY.cls` / `frmZZZ.frm`.
- **OBL-005** : **Option Explicit** tête module.
- **OBL-006** :
    - pour les modules de classe ou les modules normaux repecter l’ordre : en-tête module, Option Explicit, consts, events, procs.
    - pour les UserForms : en-tête, Option Explicit, variables de module (avec WithEvents), constantes, événements du formulaire (Initialize, Activate, Deactivate), événements des contrôles (Click, Change, etc.), procédures publiques, procédures privées, événements liés à Excel (si applicable).
- **OBL-007** : En-tête module :

```
' mod[nom module].bas - [nom module] tbAffaires
'
'-------------------------------------------------------------------------------
' GPL v3 - LICENSE
'-------------------------------------------------------------------------------
'
' Auteur : ...
' Date : ...
' Objet : ...
' Liste des éléments public (variables/procédures/fonctions)
```

- **OBL-008** : Headers procs :

```
'-------------------------------------------------------------------------------
' ProcedureName - Brève desc.
' Parameters : param - Type - Desc.
' Return     : Type - Desc.
'-------------------------------------------------------------------------------
```

## Patterns Code

### OBLIGATIONS

- **OBL-009** : ListObject tables ; Range cellules.
- **OBL-010** : Arrays bulk R/W (jusqu'à 100000 lignes, 50 colonnes).
- **OBL-011** : RAII ExcelOptimizer (Set opt = New clsExcelOptimizer en début de procédure et Set opt = Nothing en fin de procédure, utilisé même pour les procédures courtes).
- **OBL-012** : UserForm MVVM.
- **OBL-013** : Error : On Error GoTo + CleanUp + Log.
- **OBL-014** : Cache refs.

### INTERDICTIONS

- **INT-003** : Ne JAMAIS utiliser `Select` ou `Activate` SAUF pour indiquer position initiale utilisateur (utiliser `Range("X").Select` avant de rendre la main).

## Commentaires & Clean Code

### OBLIGATIONS

- **OBL-015** : **Minimiser commentaires** : code self-explanatory (noms clairs) ; commenter WHY (pas WHAT).
- **OBL-016** : Clean code : SRP (≤20 lignes/proc), DRY, KISS, YAGNI, SOLID ; exceptions spécifiques.

## Tooling

### OBLIGATIONS

- **OBL-017** : xlManage (`poetry run xlmanage`) : CRUD/sync/tests.

## Licence & Docs

### OBLIGATIONS

- **OBL-018** : Header GPL v3 (conserver le modèle mentionné dans AGENTS.md).
- **OBL-019** : Docstrings tous procs/classes (format OBL-008).
