# Exceptions xlManage - Hiérarchie et Handling

## Hiérarchie des Exceptions

```
Exception (Python builtin)
└── ExcelManageError (base xlmanage)
    ├── ExcelConnectionError
    ├── ExcelInstanceNotFoundError
    ├── ExcelRPCError
    ├── WorkbookNotFoundError
    ├── WorkbookAlreadyOpenError
    ├── WorkbookSaveError
    ├── WorksheetNotFoundError
    ├── WorksheetAlreadyExistsError
    ├── WorksheetDeleteError
    ├── WorksheetNameError
    ├── TableNotFoundError
    ├── TableAlreadyExistsError
    ├── TableNameError
    ├── TableRangeError
    ├── VBAProjectAccessError
    ├── VBAWorkbookFormatError
    ├── VBAModuleNotFoundError
    ├── VBAModuleAlreadyExistsError
    ├── VBAImportError
    ├── VBAExportError
    └── VBAMacroError
```

## Base Exception

### ExcelManageError

**Exception de base pour toutes les erreurs xlManage.**

```python
try:
    with ExcelManager() as mgr:
        mgr.start()
except ExcelManageError as e:
    print(f"xlManage error: {e.message}")
```

## Exceptions Instance/COM

### ExcelConnectionError

**Connexion COM Excel échouée.** Levée quand Excel n'est pas installé ou COM server indisponible.

**Attributs:**
- `hresult: int` - Code HRESULT COM
- `message: str` - Message d'erreur

```python
from xlmanage import ExcelManager, ExcelConnectionError

try:
    mgr = ExcelManager()
    mgr.start()
except ExcelConnectionError as e:
    print(f"COM unavailable: {e.message}, HRESULT=0x{e.hresult:X}")
    # Vérifier qu'Excel est installé et pywin32 est disponible
```

### ExcelInstanceNotFoundError

**Instance Excel introuvable.** Levée quand une instance demandée n'existe pas.

**Attributs:**
- `instance_id: str` - Identifiant de l'instance
- `message: str` - Message d'erreur

```python
from xlmanage import ExcelManager, ExcelInstanceNotFoundError

mgr = ExcelManager()
try:
    mgr.stop_instance(99999)  # PID inexistant
except ExcelInstanceNotFoundError as e:
    print(f"Instance {e.instance_id} not found")
```

### ExcelRPCError

**Erreur RPC Excel.** Levée quand COM server est déconnecté ou indisponible.

**Attributs:**
- `hresult: int` - Code HRESULT COM
- `message: str` - Message d'erreur

```python
from xlmanage import ExcelManager, ExcelRPCError

mgr = ExcelManager()
try:
    mgr.stop_instance(12345)
except ExcelRPCError as e:
    print(f"RPC error (instance likely dead): 0x{e.hresult:X}")
    # Instance est déconnectée, essayer force_kill()
    mgr.force_kill(12345)
```

## Exceptions Workbook

### WorkbookNotFoundError

**Classeur introuvable sur le disque.** Levée quand le fichier n'existe pas.

**Attributs:**
- `path: Path` - Chemin du fichier introuvable
- `message: str` - Message d'erreur

```python
from xlmanage import WorkbookManager, WorkbookNotFoundError

wb_mgr = WorkbookManager(app)
try:
    wb_mgr.open(Path("nonexistent.xlsx"))
except WorkbookNotFoundError as e:
    print(f"File not found: {e.path}")
```

### WorkbookAlreadyOpenError

**Classeur déjà ouvert dans l'instance.** Levée quand on tente d'ouvrir un fichier déjà ouvert.

**Attributs:**
- `path: Path` - Chemin du fichier
- `name: str` - Nom du classeur
- `message: str` - Message d'erreur

```python
from xlmanage import WorkbookManager, WorkbookAlreadyOpenError

wb_mgr = WorkbookManager(app)
try:
    # Première ouverture réussie
    wb_mgr.open(Path("data.xlsx"))
    # Deuxième ouverture échoue
    wb_mgr.open(Path("data.xlsx"))
except WorkbookAlreadyOpenError as e:
    print(f"'{e.name}' is already open")
    # Utiliser le classeur déjà ouvert ou le fermer d'abord
```

### WorkbookSaveError

**Échec de sauvegarde du classeur.** Levée par permissions, path invalide, ou format incompatible.

**Attributs:**
- `path: Path` - Chemin du fichier
- `hresult: int` - Code HRESULT COM
- `message: str` - Message d'erreur

```python
from xlmanage import WorkbookManager, WorkbookSaveError

wb_mgr = WorkbookManager(app)
try:
    wb_mgr.save(name="Workbook1", path=Path("/readonly/result.xlsx"))
except WorkbookSaveError as e:
    print(f"Save failed: {e.message}, HRESULT=0x{e.hresult:X}")
    # Vérifier permissions et path
```

## Exceptions Worksheet

### WorksheetNotFoundError

**Feuille introuvable dans le classeur.**

**Attributs:**
- `name: str` - Nom de la feuille
- `workbook_name: str` - Nom du classeur

```python
from xlmanage import WorksheetManager, WorksheetNotFoundError

ws_mgr = WorksheetManager(app)
try:
    ws_mgr.delete("NonExistentSheet")
except WorksheetNotFoundError as e:
    print(f"Sheet '{e.name}' not found in '{e.workbook_name}'")
```

### WorksheetAlreadyExistsError

**Nom de feuille déjà utilisé.**

**Attributs:**
- `name: str` - Nom de la feuille
- `workbook_name: str` - Nom du classeur

```python
from xlmanage import WorksheetManager, WorksheetAlreadyExistsError

ws_mgr = WorksheetManager(app)
try:
    ws_mgr.create("Sheet1")  # Existe déjà
except WorksheetAlreadyExistsError as e:
    print(f"Sheet '{e.name}' already exists")
```

### WorksheetDeleteError

**Suppression de feuille impossible.** Levée quand la dernière feuille visible ne peut être supprimée.

**Attributs:**
- `name: str` - Nom de la feuille
- `reason: str` - Raison de l'échec

```python
from xlmanage import WorksheetManager, WorksheetDeleteError

ws_mgr = WorksheetManager(app)
try:
    ws_mgr.delete("LastSheet")  # Dernière feuille visible
except WorksheetDeleteError as e:
    print(f"Cannot delete '{e.name}': {e.reason}")
    # Créer une autre feuille d'abord
```

### WorksheetNameError

**Nom de feuille invalide.** Violation des règles de nommage Excel.

**Attributs:**
- `name: str` - Nom invalide
- `reason: str` - Raison (caractères interdits, longueur, etc.)

```python
from xlmanage import WorksheetManager, WorksheetNameError

ws_mgr = WorksheetManager(app)
try:
    ws_mgr.create("Sheet:1")  # ':' est interdit
except WorksheetNameError as e:
    print(f"Invalid name '{e.name}': {e.reason}")
```

## Exceptions Table (ListObject)

### TableNotFoundError

**Table introuvable dans la feuille.**

**Attributs:**
- `name: str` - Nom de la table
- `worksheet_name: str` - Nom de la feuille

```python
from xlmanage import TableManager, TableNotFoundError

tbl_mgr = TableManager(app)
try:
    tbl_mgr.delete("tblNonExistent")
except TableNotFoundError as e:
    print(f"Table '{e.name}' not found in '{e.worksheet_name}'")
```

### TableAlreadyExistsError

**Nom de table déjà utilisé.**

**Attributs:**
- `name: str` - Nom de la table
- `workbook_name: str` - Nom du classeur

```python
from xlmanage import TableManager, TableAlreadyExistsError

tbl_mgr = TableManager(app)
try:
    tbl_mgr.create("tblSales", "A1:D100")
    tbl_mgr.create("tblSales", "A1:D100")  # Dupliqué
except TableAlreadyExistsError as e:
    print(f"Table '{e.name}' already exists")
```

### TableNameError

**Nom de table invalide.** Violation des règles de nommage Excel.

**Attributs:**
- `name: str` - Nom invalide
- `reason: str` - Raison

```python
from xlmanage import TableManager, TableNameError

tbl_mgr = TableManager(app)
try:
    tbl_mgr.create("1stTable", "A1:D10")  # Commence par chiffre
except TableNameError as e:
    print(f"Invalid name '{e.name}': {e.reason}")
```

### TableRangeError

**Plage de table invalide.** Erreur de syntaxe, plage vide, chevauchement, etc.

**Attributs:**
- `range_ref: str` - Référence de plage invalide
- `reason: str` - Raison

```python
from xlmanage import TableManager, TableRangeError

tbl_mgr = TableManager(app)
try:
    tbl_mgr.create("tblEmpty", "A1:A1")  # Plage 1x1 invalide
except TableRangeError as e:
    print(f"Invalid range '{e.range_ref}': {e.reason}")
```

## Exceptions VBA

### VBAProjectAccessError

**Accès au projet VBA refusé par le Trust Center.**

**Attributs:**
- `workbook_name: str` - Nom du classeur

```python
from xlmanage import VBAManager, VBAProjectAccessError

vba_mgr = VBAManager(app)
try:
    modules = vba_mgr.list_modules()
except VBAProjectAccessError as e:
    print(f"VBA access blocked: Enable 'Trust access to the VBA project object model' "
          f"in Excel Trust Center for '{e.workbook_name}'")
```

### VBAWorkbookFormatError

**Classeur au format .xlsx ne supportant pas les macros.**

**Attributs:**
- `workbook_name: str` - Nom du classeur

```python
from xlmanage import VBAManager, VBAWorkbookFormatError

vba_mgr = VBAManager(app)
try:
    vba_mgr.import_module(Path("module.bas"))
except VBAWorkbookFormatError as e:
    print(f"File '{e.workbook_name}' is .xlsx format. "
          f"Convert to .xlsm to support VBA")
```

### VBAModuleNotFoundError

**Module VBA introuvable dans le projet.** Ou module non-suppressible (document modules).

**Attributs:**
- `module_name: str` - Nom du module
- `workbook_name: str` - Nom du classeur
- `reason: str` - Raison supplémentaire

```python
from xlmanage import VBAManager, VBAModuleNotFoundError

vba_mgr = VBAManager(app)
try:
    vba_mgr.export_module("modNonExistent", Path("output.bas"))
except VBAModuleNotFoundError as e:
    print(f"Module '{e.module_name}' not found")
    if e.reason:
        print(f"Reason: {e.reason}")
```

### VBAModuleAlreadyExistsError

**Module VBA avec ce nom existe déjà.**

**Attributs:**
- `module_name: str` - Nom du module
- `workbook_name: str` - Nom du classeur

```python
from xlmanage import VBAManager, VBAModuleAlreadyExistsError

vba_mgr = VBAManager(app)
try:
    vba_mgr.import_module(Path("modExisting.bas"))
except VBAModuleAlreadyExistsError as e:
    print(f"Module '{e.module_name}' already exists")
    # Renommer ou supprimer d'abord
```

### VBAImportError

**Échec d'import de module VBA.** Fichier invalide, mauvais encodage, etc.

**Attributs:**
- `module_file: Path` - Chemin du fichier
- `reason: str` - Raison

```python
from xlmanage import VBAManager, VBAImportError

vba_mgr = VBAManager(app)
try:
    vba_mgr.import_module(Path("invalid_module.bas"))
except VBAImportError as e:
    print(f"Import failed for '{e.module_file}': {e.reason}")
```

### VBAExportError

**Échec d'export de module VBA.** Permissions, path invalide, etc.

**Attributs:**
- `module_name: str` - Nom du module
- `output_path: Path` - Chemin de sortie
- `reason: str` - Raison

```python
from xlmanage import VBAManager, VBAExportError

vba_mgr = VBAManager(app)
try:
    vba_mgr.export_module("modUtils", Path("/readonly/output.bas"))
except VBAExportError as e:
    print(f"Export failed: {e.reason}")
```

### VBAMacroError

**Échec d'exécution ou de parsing de macro VBA.**

**Attributs:**
- `macro_name: str` - Nom de la macro
- `reason: str` - Explication de l'échec

```python
from xlmanage import MacroRunner, VBAMacroError

runner = MacroRunner(mgr)
try:
    result = runner.run("Module1.NonExistentMacro")
except VBAMacroError as e:
    print(f"Macro error: {e.macro_name} - {e.reason}")
```

## Pattern d'Error Handling Complet

```python
from pathlib import Path
from xlmanage import (
    ExcelManager, WorkbookManager, WorksheetManager,
    TableManager, VBAManager, MacroRunner,
    ExcelManageError,  # Importer toutes les exceptions spécifiques
    # ... autres imports ...
)

def robust_processing(file_path: Path):
    """Traite un fichier Excel avec error handling complet"""
    try:
        with ExcelManager(visible=False) as mgr:
            mgr.start()

            # Ouverture
            wb_mgr = WorkbookManager(mgr.app)
            info = wb_mgr.open(file_path)

            # Création feuille
            ws_mgr = WorksheetManager(mgr.app)
            sheet = ws_mgr.create("Processed")

            # Création table
            tbl_mgr = TableManager(mgr.app)
            tbl = tbl_mgr.create("tblResults", "A1:D100")

            # Exécution macro
            runner = MacroRunner(mgr)
            result = runner.run("modProcess.Run")

            return result

    except WorkbookNotFoundError as e:
        print(f"Input file not found: {e.path}")
        return None

    except WorksheetAlreadyExistsError as e:
        print(f"Sheet '{e.name}' exists, using it")
        # Continuer avec feuille existante

    except TableAlreadyExistsError as e:
        print(f"Table '{e.name}' exists, replacing it")
        # Supprimer et recréer

    except VBAMacroError as e:
        print(f"Macro failed: {e.reason}")
        # Continuer sans macro

    except ExcelManageError as e:
        print(f"xlManage error: {e.message}")
        return None

    except Exception as e:
        print(f"Unexpected error: {type(e).__name__}: {e}")
        return None
```

## Documentation API

Pour voir la liste complète des exceptions et leurs détails, utiliser Python :

```python
from xlmanage import ExcelManageError
import inspect

# Voir la hiérarchie des exceptions
print(inspect.getdoc(ExcelManageError))

# Lister toutes les sous-classes d'exceptions
from xlmanage import *
exceptions = [name for name in dir() if 'Error' in name]
for exc_name in exceptions:
    print(exc_name)
```
