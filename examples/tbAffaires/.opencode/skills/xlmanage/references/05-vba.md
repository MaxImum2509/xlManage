# VBAManager et MacroRunner - Gestion VBA

## Imports

```python
from xlmanage import VBAManager, VBAModuleInfo, MacroRunner, MacroResult
```

## VBAManager - Import/Export Modules

### Instanciation

```python
mgr = ExcelManager()
mgr.start()

vba_mgr = VBAManager(mgr.app)
```

### Méthodes CRUD

| Méthode | Paramètres | Retour | Description |
|----------|------------|--------|-------------|
| `list_modules()` | - | `List[VBAModuleInfo]` | Liste tous les modules VBA |
| `import_module(file_path)` | `file_path: Path` | `None` | Importe module (.bas/.cls/.frm) |
| `export_module(module_name, output_path)` | `module_name: str, output_path: Path` | `None` | Exporte module vers fichier |
| `delete_module(module_name)` | `module_name: str` | `None` | Supprime module |

### Types de Modules VBA

| Type | Extension | COM Type | Description |
|------|-----------|-----------|-------------|
| Standard | `.bas` | `vbext_ct_StdModule` (1) | Module standard (Sub/Function) |
| Class | `.cls` | `vbext_ct_ClassModule` (2) | Module de classe |
| UserForm | `.frm` (+`.frx`) | `vbext_ct_MSForm` (3) | Formulaire utilisateur |
| Document | - | `vbext_ct_Document` (100) | Feuille/ThisWorkbook |

## Usage Patterns

### Lister les Modules VBA

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    vba_mgr = VBAManager(mgr.app)

    modules = vba_mgr.list_modules()

    for mod in modules:
        print(f"{mod.name}: {mod.module_type}, "
              f"{mod.lines_count} lines, "
              f"predeclared={mod.has_predeclared_id}")
```

### Importer un Module Standard (.bas)

```python
from pathlib import Path

with ExcelManager(visible=False) as mgr:
    mgr.start()
    vba_mgr = VBAManager(mgr.app)

    # Importer module standard
    vba_mgr.import_module(Path("vba/modUtils.bas"))
    print("Imported modUtils")
```

### Importer un Module de Classe (.cls)

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    vba_mgr = VBAManager(mgr.app)

    # Importer module de classe
    vba_mgr.import_module(Path("vba/clsOptimizer.cls"))
    print("Imported clsOptimizer")
```

### Importer un UserForm (.frm + .frx)

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    vba_mgr = VBAManager(mgr.app)

    # .frx doit être dans le même répertoire que .frm
    vba_mgr.import_module(Path("vba/frmMain.frm"))
    print("Imported frmMain UserForm")
```

### Exporter un Module

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    vba_mgr = VBAManager(mgr.app)

    # Exporter module standard
    vba_mgr.export_module("modUtils", Path("backup/modUtils.bas"))

    # Exporter module de classe
    vba_mgr.export_module("clsOptimizer", Path("backup/clsOptimizer.cls"))

    # Exporter UserForm (génère .frm et .frx)
    vba_mgr.export_module("frmMain", Path("backup/frmMain.frm"))
```

### Supprimer un Module

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    vba_mgr = VBAManager(mgr.app)

    # Supprimer module standard ou de classe
    vba_mgr.delete_module("modUtils")
    vba_mgr.delete_module("clsOptimizer")

    # NOTE: Les modules Document (ThisWorkbook, Sheet1...) ne peuvent pas être supprimés
```

## VBAModuleInfo Structure

```python
dataclass VBAModuleInfo:
    name: str                  # Nom du module
    module_type: str           # Type (1=Standard, 2=Class, 3=Form, 100=Document)
    lines_count: int           # Nombre de lignes de code
    has_predeclared_id: bool   # Pour classes: PredeclaredId (singleton pattern)
```

## MacroRunner - Exécution Macros

### Instanciation

```python
mgr = ExcelManager()
mgr.start()

runner = MacroRunner(mgr)
```

### Méthode `run()`

| Paramètre | Type | Description |
|-----------|-------|-------------|
| `macro_name` | `str` | Nom macro ("Module1.MySub" ou "MySub") |
| `workbook` | `Path|None` | Classeur contenant la macro (None=actif ou PERSONAL.XLSB) |
| `args` | `str|None` | Arguments CSV (`'"hello"',42,true`) |

**Retour:** `MacroResult`

### Exécuter une Macro Sans Arguments

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    runner = MacroRunner(mgr)

    result = runner.run("Module1.ProcessData")

    if result.success:
        print(f"Macro executed successfully")
        if result.return_value is not None:
            print(f"Return: {result.return_value} ({result.return_type})")
    else:
        print(f"Error: {result.error_message}")
```

### Exécuter une Macro Avec Arguments

```python
with ExcelManager(visible=False) as mgr:
    mgr.start()
    runner = MacroRunner(mgr)

    # Arguments CSV : guillemets pour strings, séparés par virgules
    result = runner.run("Module1.CalculateTotal", args='"Product A",10,25.50')

    if result.success:
        print(f"Total: {result.return_value}")
```

### Exécuter une Macro d'un Classeur Spécifique

```python
from pathlib import Path

with ExcelManager(visible=False) as mgr:
    mgr.start()
    runner = MacroRunner(mgr)

    result = runner.run(
        "Module1.GetData",
        workbook=Path("macros/personal.xlsm")
    )
```

## MacroResult Structure

```python
dataclass MacroResult:
    macro_name: str           # Nom complet de la macro
    return_value: Any|None    # Valeur retournée (None pour Sub)
    return_type: str          # Type Python ("str", "int", "float", "NoneType")
    success: bool             # True si exécution sans erreur VBA
    error_message: str|None   # Message erreur VBA si échec
```

## Exceptions

| Exception | Condition | Contexte |
|-----------|-----------|-----------|
| `VBAProjectAccessError` | Trust Center bloque accès VBA | `list_modules()`, `import_module()`, etc. |
| `VBAWorkbookFormatError` | Classeur .xlsx (sans macros) | `list_modules()`, `import_module()`, etc. |
| `VBAModuleNotFoundError` | Module introuvable | `export_module()`, `delete_module()` |
| `VBAModuleAlreadyExistsError` | Module existe déjà | `import_module()` |
| `VBAImportError` | Échec import (fichier invalide, encodage) | `import_module()` |
| `VBAExportError` | Échec export (permissions, path) | `export_module()` |
| `VBAMacroError` | Échec exécution/parsing macro | `MacroRunner.run()` |

## Workflow Complet

```python
from pathlib import Path
from xlmanage import (
    ExcelManager, VBAManager, MacroRunner,
    VBAProjectAccessError, VBAWorkbookFormatError, VBAMacroError
)

def sync_vba_modules(workbook_path: Path, vba_dir: Path):
    """Synchronise les modules VBA depuis un répertoire"""
    with ExcelManager(visible=False) as mgr:
        mgr.start()
        vba_mgr = VBAManager(mgr.app)
        runner = MacroRunner(mgr)

        try:
            # Lister modules existants
            existing = vba_mgr.list_modules()
            existing_names = {m.name for m in existing}
            print(f"Existing modules: {existing_names}")

            # Importer nouveaux modules
            for vba_file in vba_dir.glob("*.bas"):
                module_name = vba_file.stem

                if module_name not in existing_names:
                    print(f"Importing {module_name}...")
                    vba_mgr.import_module(vba_file)
                else:
                    print(f"Skipping {module_name} (exists)")

            # Exécuter macro d'initialisation si elle existe
            try:
                result = runner.run("modMain.Initialize")
                if result.success:
                    print("Initialization macro executed")
            except VBAMacroError:
                print("No initialization macro found")

        except VBAProjectAccessError:
            print("ERROR: Trust Center blocks VBA access. "
                  "Enable 'Trust access to the VBA project object model'")
        except VBAWorkbookFormatError:
            print(f"ERROR: {workbook_path} is .xlsx format. "
                  "VBA requires .xlsm")

# Usage
sync_vba_modules(
    Path("data/myapp.xlsm"),
    Path("vba/modules")
)
```

## VBA File Encoding

**CRITICAL:** Tous les fichiers VBA doivent utiliser **Windows-1252** encodage avec CRLF.

```python
# Lire fichier VBA
with open("module.bas", "r", encoding="windows-1252") as f:
    content = f.read()

# Écrire fichier VBA
with open("module.bas", "w", encoding="windows-1252", newline="\r\n") as f:
    f.write(vba_code)
```

## Documentation API

Pour voir l'API complète, utiliser les docstrings Python :

```python
from xlmanage import VBAManager, MacroRunner
import inspect

print(inspect.getdoc(VBAManager))
print(inspect.getdoc(MacroRunner))
```
