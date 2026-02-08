# AGENTS.md - Instructions pour agents sur tbAffaires

## Project overview

`tbAffaires` est une application VBA/Excel sur Active Directory pour reporting hebdomadaire (3 ADV, ~2500 affaires). Travail parallèle (plages exclusives), commentaires historiques auto (data.xlsx), UPSERT incrémental + retry. xlManage pour dev VBA: sync src/ ↔ tbAffaires.xlsm, data.xlsx, tests.

**Stack:**
- VBA Excel 2016+
- tbAffaires.xlsm (modConfiguration.bas, clsApplicationState.cls...)
- data.xlsx (tbADV, tbParametres, tbMapping, tbCommentaires)
- xlManage (pywin32) via src/xlmanage/
- Sources: src/ (.bas/.cls)
- Licence: GPL v3

**Arborescence:**
- src/, tests/, docs/, scripts/, _dev/ (architecture.md, stories/, reports/, planning)

**Langues:**
- Docs/comm: Français
- VBA: Français PascalCase
- Python: Anglais

**Fonctionnalités:**
- ID: Environ("USERNAME") → tbADV
- RAII: clsApplicationState
- ERP load (dialog, read-only)
- Filtre ADV ListObject
- Commentaires auto
- UPSERT retry 5x
- Log: DATE|USER|ACTION|RESULTAT
- Admin mode

## Build and test commands

- `poetry install`
- `poetry run pytest scripts/tests/`
- `poetry run xlmanage sync src/ tbAffaires.xlsm`
- pre-commit

## Code style guidelines

### VBA
**OBL:**
- OBL-001: Windows-1252 + CRLF
- OBL-002: Chemins /
- OBL-003: PascalCase FR
- OBL-004: modXXX.bas/clsYYY.cls
- OBL-005: Option Explicit
- OBL-006: Ordre en-tête/Option/consts/events/procs
- OBL-007: Header GPL
- OBL-008: Proc headers
- OBL-009+: ListObject/arrays/RAII/MVVM/error GoTo+Log/cache
- OBL-015+: Min comments WHY/clean SRP≤20l/DRY/KISS/YAGNI/SOLID

**INT:**
- INT-001: No emojis
- INT-002: No Select/Activate sans Explicit

### Python
**OBL:**
- OBL-001: / ou pathlib
- OBL-002+: Poetry CLI/run/PEP518+
- OBL-005+: UTF8 LF/GPL/anglais/arbo/tooling3.14/PEP8/clean/docstrings/type hints/≥90%cov/ruff/mypy/bandit/pre-commit/commits

**INT:**
- INT-001+: No \ /direct pyproject.toml/pip/emojis(no docless)/except:

### Chemins/Poetry
- OBL-CHEMINS-001: / ou pathlib
- INT-CHEMINS-001-003: No \
- OBL-POETRY-001-005: Poetry CLI
- INT-POETRY-001-004: No direct edit
- EXP-001-003: tool/project/scripts only

### Git commits (Conventional)
`type[(scope)]: résumé ≤50c`

Types: feat/fix/docs/style/refactor/test/chore/perf/ci/build/revert

Impératif, breaking !/BREAKING CHANGE.

Ex: `feat(vba): ajouter CRUD`

## Testing instructions

- VBA: Manuels 5 scénarios (_dev/tests-playbook.md): nominal/conflit/mapping/admin/edges
- Python: pytest scripts/tests/ (mock/cov90+/timeout/xdist), conftest.py anti-zombie
- Valider: ruff/mypy/bandit/pre-commit

## Security considerations

- No secrets/repo (.env .gitignore)
- Bandit Python scans
- VBA: read-only ERP, safe macros
- COM: RAII no zombies
- Review changes, no force push main
