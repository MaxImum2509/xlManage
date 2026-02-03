---
name: python-poetry-userguide
description: Poetry package manager guide for xlManage project. Covers mandatory Poetry CLI commands for dependency management (add, remove, update, install), virtual environment setup, forbidden actions (never edit pyproject.toml dependencies directly), allowed exceptions (tool configs only), development workflow (tests, linting, type checking, docs). Use when managing Python packages, adding/removing dependencies, setting up the development environment, or running project commands.
---

# Poetry Package Management

Mandatory rules and workflow for dependency management in xlManage.

## Obligation: Poetry CLI Only

**NEVER** edit `pyproject.toml` directly for dependencies. **NEVER** use pip.

### Required Commands [OBL]

| Command | Purpose | Ref |
|---------|---------|-----|
| `poetry add <package>` | Add production dependency | [OBL-001] |
| `poetry add --group dev <package>` | Add dev dependency | [OBL-002] |
| `poetry remove <package>` | Remove dependency | [OBL-003] |
| `poetry update <package>` | Update specific dependency | [OBL-004] |
| `poetry install` | Install from lock file | [OBL-005] |

### Forbidden Actions [INT]

**NEVER** modify `pyproject.toml` directly for:

| Action | Ref |
|--------|-----|
| Adding/removing dependencies | [INT-001] |
| Changing package versions | [INT-002] |
| Modifying Poetry configuration | [INT-003] |
| Editing `[tool.poetry.dependencies]` or `[tool.poetry.group.*.dependencies]` | [INT-004] |

### Allowed Exceptions [EXP]

Manual `pyproject.toml` editing is ONLY allowed for:

| Context | Ref |
|---------|-----|
| Tool configs (`[tool.ruff]`, `[tool.mypy]`, `[tool.pytest.*]`) | [EXP-001] |
| Project metadata (`[project]`) during initial setup | [EXP-002] |
| Entry point scripts (`[project.scripts]`) | [EXP-003] |

## Setup

```bash
# Install all dependencies (including dev)
poetry install --with dev

# Enter virtual environment
poetry shell

# Or run commands directly
poetry run <command>
```

## Adding Packages

```bash
# Production
poetry add requests
poetry add "requests>=2.28.0"

# Development
poetry add --group dev pytest
```

## Removing / Updating

```bash
poetry remove requests
poetry update requests
poetry update  # All packages
```

## Development Workflow

### Run Application

```bash
poetry run python -m xlmanage --help
```

### Run Tests

```bash
poetry run pytest
poetry run pytest --cov=src/ --cov-report=html --cov-report=term --cov-fail-under=90
```

### Code Quality

```bash
poetry run ruff check .
poetry run ruff format .
poetry run mypy src/
```

### Build Documentation

```bash
poetry run sphinx-build -b html docs docs/_build
```

### Run Utility Scripts

```bash
poetry run python scripts/script_name.py
```

## Environment Info

```bash
poetry env info --path    # Virtual env path
poetry show               # Installed packages
poetry show --tree        # Dependency tree
```

## Anti-Patterns

| Never | Alternative |
|-------|-------------|
| `pip install pkg` | `poetry add pkg` |
| Edit pyproject.toml deps | `poetry add/remove` |
| Manual version changes | `poetry update pkg` |
| `pip freeze > requirements.txt` | `poetry export` |

---

## Constraint Reminders

### Obligations [OBL]
- Manage packages with Poetry (**[OBL-001]** to **[OBL-005]**)

### Prohibitions [INT]
- **NEVER** modify `pyproject.toml` directly (**[INT-001]** to **[INT-004]**)
