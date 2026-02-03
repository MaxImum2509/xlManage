# Clean Code Principles

Core principles for maintainable, readable Python code.

## SRP - Single Responsibility

### Functions: one thing, well

```python
# Bad: multiple responsibilities
def process_excel_file(file_path):
    workbook = read_workbook(file_path)
    validate_workbook(workbook)
    format_workbook(workbook)
    save_workbook(workbook, "output.xlsx")

# Good: single responsibility per function
def process_excel_file(file_path):
    workbook = read_workbook(file_path)
    validate_workbook(workbook)
    format_workbook(workbook)
    return workbook
```

### Classes: one reason to change

```python
# Bad
class ExcelManager:
    def read_file(self): pass
    def send_email(self): pass
    def log_to_database(self): pass

# Good
class ExcelManager:
    def read_file(self): pass

class EmailSender:
    def send_email(self): pass
```

## DRY - Don't Repeat Yourself

```python
# Bad: repeated validation
def process_user(user):
    if user is None: raise ValueError("User required")
    if user.name is None or user.name == "": raise ValueError("User name required")

# Good: extracted
def validate_required(value, field_name):
    if value is None or value == "":
        raise ValueError(f"{field_name} required")

def process_user(user):
    validate_required(user, "User")
    validate_required(user.name, "User name")
```

## KISS - Keep It Simple

```python
# Bad: deeply nested
def process_data(data):
    if data is not None:
        if len(data) > 0:
            if data[0] is not None:
                return data[0].value

# Good: guard clauses
def process_data(data):
    if data is None: return None
    if len(data) == 0: return None
    if data[0] is None: return None
    return data[0].value
```

## YAGNI - You Aren't Gonna Need It

```python
# Bad: over-engineered
class ExcelManager:
    def __init__(self):
        self.supported_formats = ["xlsx", "xls", "csv", "ods", "pdf"]
        self.optimization_levels = ["none", "low", "medium", "high", "extreme"]

# Good: minimal
class ExcelManager:
    def __init__(self):
        self.supported_formats = ["xlsx"]
```

## Error Handling

```python
# Bad
try:
    workbook = open_excel(file_path)
except:
    print("Error occurred")

# Good
try:
    workbook = open_excel(file_path)
except FileNotFoundError:
    logger.error(f"File not found: {file_path}")
except PermissionError:
    logger.error(f"Permission denied: {file_path}")
except Exception as e:
    logger.error(f"Unexpected error: {e}")
    raise
```

## Function Length: < 20 lines

```python
# Good: split into smaller functions
def process_workbook_data(workbook):
    all_cells = get_all_cell_values(workbook)
    valid_numbers = filter_valid_numbers(all_cells)
    return filter_positive_numbers(valid_numbers)
```

## SOLID Principles

### O - Open/Closed
Open for extension, closed for modification.

### L - Liskov Substitution
Subtypes must be substitutable for base types.

### I - Interface Segregation
Specific interfaces over general ones.

### D - Dependency Inversion

```python
# Bad
class ExcelManager:
    def __init__(self):
        self.storage = FileSystemStorage()  # Concrete

# Good
class ExcelManager:
    def __init__(self, storage: Storage):
        self.storage = storage  # Abstract
```

## Comments: Why, Not What

```python
# Bad
counter += 1  # Increment counter by 1

# Good
counter = atomic_increment(counter)  # Ensure thread safety

# Best: self-documenting code
distance_in_km = distance_in_miles * 1.609
```
