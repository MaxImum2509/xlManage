# AGENTS.md - Instructions pour agents sur tbAffaires

## Project overview

`tbAffaires` est une application VBA/Excel sur Active Directory pour reporting hebdomadaire (3 ADV, ~2500 affaires). Travail parallèle (plages exclusives), commentaires historiques auto (data.xlsx), UPSERT incrémental + retry. xlManage pour dev VBA: sync src/ ↔ tbAffaires.xlsm, data.xlsx, tests.

## Arborescence

- app/ : application finale (tbAffaires.xlsm) et ses données,
- src/ : sources du projet VBA,
- tests/ : tests en VBA de l’application finals,
- docs/ : documentation utilisateur et administrateur,
- scripts/ : tooling python et shell (bash) pour l’import des modules de code VBA dans le fichier application (app/tbAffaires.xlsm),
- \_dev/ : architecture.md, stories/, reports/, planning/

## Langues

- Documentation : Français
- VBA : Nommage code Français PascalCase / Commentaires : Français
- Python : Anglais
- Shell scripts : Anglais

**Fonctionnalités:**

- ID: Environ("USERNAME") → tbADV
- RAII: clsApplicationState
- ERP load (dialog, read-only)
- Filtre ADV ListObject
- Commentaires auto
- UPSERT retry 5x
- Log: DATE|USER|ACTION|RESULTAT
- Admin mode

## Code style guidelines

### VBA

**Lire et respecter** : @docs/excel-development-rules.md.

### Python

**Lire et respecter** : @docs/python-development-rules.md.

### Git commits (Conventional)

**Lire et respecter** : @docs/git-commit-rules.md.

## Processus de dévelopment

**Lire et respecter** : @docs/excel-development-process.md.
