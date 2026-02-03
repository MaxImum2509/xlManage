# Python Naming Conventions (PEP 8)

Detailed naming rules for Python code in xlManage project.

## File Naming

- Pattern: `snake_case`
- One class per file, file name matches class name

```
excel_manager.py     # class ExcelManager
data_processor.py    # class DataProcessor
```

## Variable Naming

```python
# Regular: snake_case, descriptive
workbook_name, cell_value, max_rows

# Constants: UPPER_CASE
MAX_RETRIES = 3
DEFAULT_TIMEOUT = 30

# Private: _prefix
_internal_state, _workbook_cache
```

## Function Naming

```python
# Regular: snake_case, verb form
def create_workbook(): pass
def validate_cell_value(): pass

# Private: _prefix
def _validate_input(): pass
```

## Class Naming

```python
# Regular: PascalCase, noun
class ExcelManager: pass
class WorkbookProcessor: pass

# Exceptions: PascalCase + Error
class ExcelError(Exception): pass
class ValidationError(ValueError): pass
```

## Testing Naming

```python
# Files: test_<module>.py
# Functions: test_<function>
# Classes: Test<Class>

class TestExcelManager:
    def test_create_workbook(self): pass
    def test_create_workbook_with_empty_name(self): pass
```

## Guidelines

- Be descriptive, avoid abbreviations (except `id`, `url`, `config`, `max`, `min`)
- Avoid numbers in names (`process1` -> `process_input`)
- Avoid type prefixes (`str_name` -> `user_name`)
- Avoid negated names (`is_not_valid` -> `is_valid`)
- Use consistent verb patterns (`get_*`, `create_*`, `validate_*`)
