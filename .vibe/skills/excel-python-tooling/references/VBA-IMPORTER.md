# VBAImporter

Unified interface for importing/exporting all VBA module types with automatic dependency detection.

## Quick Start

```python
from VBAImporter import VBAImporter

with VBAImporter(r"C:\path\to\workbook.xlsm") as importer:
    importer.import_module(r"C:\path\to\modUtils.bas")
    importer.import_module(r"C:\path\to\clsOptimizer.cls")
    importer.import_module(r"C:\path\to\frmDialog.frm")
    # Auto-save on exit
```

## Batch Import with Dependencies

```python
with VBAImporter(r"C:\path\to\workbook.xlsm") as importer:
    count = importer.import_directory(
        r"C:\vba\modules",
        pattern="*.bas *.cls *.frm",
        overwrite=True,
        auto_dependencies=True
    )
```

Dependencies detected by scanning for `Dim obj As New cls...` patterns. Import order: classes first, standard modules, UserForms last.

## Export

```python
with VBAImporter(r"C:\path\to\workbook.xlsm") as importer:
    importer.export_module("modUtils", r"C:\backup\modules")
    count = importer.export_all_modules(r"C:\backup\modules", pattern="*")

    for module in importer.list_modules():
        print(f"{module['name']} - {module['type']}")
```

## Module Type Support

| Type | Extension | Handling |
|------|-----------|----------|
| Standard | `.bas` | Direct import via VBComponents.Import |
| Class | `.cls` | Header parsing, VB_PredeclaredId handling, code injection |
| UserForm | `.frm`+`.frx` | Import via VBComponents.Import with binary UI data |

## Class Module Import Details

**CRITICAL**: Excel doesn't handle class module import headers correctly. VBAImporter:

1. Reads file with `encoding='windows-1252'`
2. Extracts `VB_Name` from `Attribute VB_Name = "..."` line
3. Extracts `VB_PredeclaredId` (True/False)
4. Strips header (everything before `Option Explicit`)
5. Creates component: `VBComponents.Add(2)` (vbext_ct_ClassModule)
6. Sets name and `Properties("PredeclaredId")` **before** adding code
7. Clears auto-generated content, then `AddFromString(clean_code)`

**Set PredeclaredId before injecting code.** Wrong order causes "Object variable not set" errors.

## VB_PredeclaredId

| Value | Behavior | Use Case |
|-------|----------|----------|
| `True` | Singleton (accessible without `New`) | Forms, global state |
| `False` | Standard class (requires `New`) | Services, helpers |

## Generating VBA from Python

```python
def generate_vba_module(module_name: str, code: str, output_path: Path):
    header = f'Attribute VB_Name = "{module_name}"\nOption Explicit\n\n'
    with open(output_path, "w", encoding="windows-1252", newline="\r\n") as f:
        f.write(header + code)
```

## Best Practices

1. **Use VBAImporter class** for all production work
2. **Windows-1252 encoding** for all VBA files
3. **Extract VB_Name** from file, never hardcode
4. **Respect PredeclaredId** - critical for singleton patterns
5. **Never call `excel.Quit()`**
6. **Use context managers** (`with VBAImporter(...)`)
