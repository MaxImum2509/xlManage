---
name: python-development-rules
description: Python coding standards for xlManage project. PEP 8 naming, clean code (SRP/DRY/KISS/YAGNI/SOLID), Sphinx docstrings, file organization, testing with pytest (90%+ coverage), project constraints (English only, no emojis in code strings, pathlib for paths, UTF-8 encoding). Use when writing Python code, reviewing code quality, naming variables/functions/classes, writing docstrings, or organizing project files.
---

# Python Development Rules

Coding standards and conventions for the xlManage Python project.

## Quick Reference

| Topic | See |
|-------|-----|
| Naming conventions | [references/NAMING.md](references/NAMING.md) |
| Clean code & SOLID | [references/CLEAN-CODE.md](references/CLEAN-CODE.md) |
| License management | [references/LICENSE-MANAGEMENT.md](references/LICENSE-MANAGEMENT.md) |

## Project Constraints

**Language**: English only for all code (variables, functions, classes, comments, docstrings, commit messages).

**No emojis** in executable code strings (encoding issues). Allowed only in markdown docs.

**Path portability**: Always use `/` or `pathlib`, never backslashes.

**File organization**: `src/` for application code, `tests/` for tests, `scripts/` for utilities, `docs/` for documentation (Sphinx). One class per file, file name matches class name in `snake_case`.

**Testing**: pytest, target 90%+ coverage. Run: `pytest --cov=src/ --cov-report=term --cov-fail-under=90`

**Documentation**: Sphinx format required for all public functions:

```python
def create_workbook(name: str) -> Workbook:
    """
    Create new Excel workbook.

    Args:
        name: Workbook name.

    Returns:
        Workbook: New workbook object.

    Raises:
        ValueError: If name is empty.
    """
```

**License**: Detect and apply the project's license. See [references/LICENSE-MANAGEMENT.md](references/LICENSE-MANAGEMENT.md).

**Python version**: 3.14

## Clean Code Principles

- **SRP**: Functions do one thing (<20 lines target)
- **DRY**: Factorize common code
- **KISS**: Prefer simple solutions; use guard clauses over nested conditions
- **YAGNI**: Don't implement unused features
- **Error handling**: Specific exceptions over generic ones, always log errors
- **SOLID**: Single Responsibility, Open/Closed, Liskov, Interface Segregation, Dependency Inversion

See [references/CLEAN-CODE.md](references/CLEAN-CODE.md) for examples.

## Naming Conventions (PEP 8)

| Element | Convention | Example |
|---------|-----------|---------|
| Files/modules | `snake_case` | `excel_manager.py` |
| Variables | `snake_case` | `cell_value` |
| Constants | `UPPER_CASE` | `MAX_RETRIES` |
| Functions | `snake_case` (verb) | `create_workbook()` |
| Classes | `PascalCase` (noun) | `ExcelManager` |
| Exceptions | `PascalCase` + `Error` | `ValidationError` |
| Private | `_prefix` | `_internal_state` |
| Tests | `test_<name>` | `test_create_workbook()` |

See [references/NAMING.md](references/NAMING.md) for detailed rules.

## Anti-Patterns

| Never | Alternative |
|-------|-------------|
| Emojis in strings | Text equivalents |
| Backslash paths | `pathlib` or `os.path` |
| Multiple classes/file | One class per file |
| French code | English only |
| Function without docstring | Sphinx docstring |
| Generic `except:` | Specific exceptions |

## Code Quality

```bash
poetry run ruff check .
poetry run ruff format .
poetry run mypy src/
```
