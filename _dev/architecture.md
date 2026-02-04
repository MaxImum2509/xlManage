# Architecture Document

## xlManage

**Version:** 1.0.0
**Last Updated:** 2026-02-01

---

## 1. Overview

xlManage est une application CLI Windows en Python qui permet de piloter Microsoft Excel via COM automation (pywin32). Elle est destinee aux agents LLM et aux developpeurs qui ont besoin d'un controle programmatique complet sur Excel : demarrage/arret d'instances, gestion des classeurs, feuilles, tables (ListObject), modules VBA et execution de macros.

**Portee fonctionnelle :**

- Cycle de vie des instances Excel (start, stop, status)
- Operations CRUD sur les Workbooks, Worksheets, ListObjects
- Import/export/suppression de modules VBA
- Execution de macros VBA (Sub et Function)
- Optimisation des performances Excel (ecran, calcul, evenements)

## 2. System Context

### 2.1 System Boundary

```
+-----------------------------------------------------+
|                     xlManage CLI                     |
|  (Python process - single thread STA COM client)     |
|                                                      |
|  cli.py -> *_manager.py -> pywin32 COM -> Excel.exe  |
+-----------------------------------------------------+
        |                           |
        v                           v
  Terminal (stdin/stdout)     Excel.exe (out-of-process COM server)
  via Typer + Rich            via Running Object Table (ROT)
```

**A l'interieur du systeme :**

- Le process Python avec toutes les couches (CLI, managers, optimizers)
- Les references COM vers les objets Excel (Application, Workbook, Worksheet, etc.)

**A l'exterieur du systeme :**

- Le process `EXCEL.EXE` (serveur COM out-of-process)
- Le systeme de fichiers Windows (classeurs .xlsx/.xlsm)
- Le Running Object Table (ROT) Windows
- Le Trust Center Excel (parametres de securite VBA)

### 2.2 External Dependencies

| Dependance  | Role                                                | Version             |
| ----------- | --------------------------------------------------- | ------------------- |
| `typer`     | Framework CLI avec gestion des commandes et options | >= 0.12.0           |
| `rich`      | Formatage terminal (tables, couleurs, panels)       | >= 13.0.0           |
| `pywin32`   | Bridge Python <-> COM Windows (acces a Excel)       | >= 305              |
| `pythoncom` | API COM bas niveau (ROT, CoInitialize, monikers)    | inclus dans pywin32 |
| `ctypes`    | Appel natif Windows (GetWindowThreadProcessId)      | stdlib              |

### 2.3 External Interfaces

- **Excel COM API** : `win32com.client.Dispatch("Excel.Application")` - interface principale
- **Running Object Table (ROT)** : `pythoncom.GetRunningObjectTable()` - enumeration des instances
- **Windows Process API** : `tasklist` / `taskkill` - fallback pour la gestion des processus
- **Systeme de fichiers** : lecture/ecriture de fichiers .xlsx, .xlsm, .bas, .cls, .frm

## 3. Architecture Overview

### 3.1 High-Level Architecture

L'architecture suit un modele en 3 couches :

```
+---------------------------------------------------------------------+
|                        COUCHE CLI (cli.py)                          |
|  Typer commands -> validation args -> appel managers -> Rich output |
+---------------------------------------------------------------------+
        |               |               |               |
        v               v               v               v
+-------------+ +---------------+ +----------------+ +-------------+
| ExcelManager| |WorkbookManager| |WorksheetManager| | TableManager|
| (Epic 5)    | | (Epic 6)      | | (Epic 7)       | | (Epic 8)    |
+-------------+ +---------------+ +----------------+ +-------------+
        |               |               |               |
        |       +---------------+ +------------+        |
        |       | VBAManager    | |MacroRunner |        |
        |       | (Epic 9)      | |(Epic 12)   |        |
        |       +---------------+ +------------+        |
        |               |               |               |
+-------------------------------------------------------------------+
|                  COUCHE OPTIMISATION (existante)                  |
|  ExcelOptimizer | ScreenOptimizer | CalculationOptimizer          |
+-------------------------------------------------------------------+
        |
        v
+------------------------------------------------------------------+
|                    COUCHE COM (pywin32)                          |
|  win32com.client.Dispatch / DispatchEx / gencache.EnsureDispatch |
|  pywintypes.com_error | pythoncom (ROT, CoInitialize)            |
+------------------------------------------------------------------+
        |
        v
+------------------------------------------------------------------+
|                    EXCEL.EXE (out-of-process)                    |
+------------------------------------------------------------------+
```

### 3.2 Architecture Patterns

- **RAII (Resource Acquisition Is Initialization)** : tous les managers implementent `__enter__` / `__exit__` pour garantir la liberation ordonnee des references COM meme en cas d'exception. C'est le pattern central du projet.
- **Injection de dependances** : chaque manager recoit un `ExcelManager` en parametre au lieu de creer sa propre instance Excel. Cela permet le partage d'une instance unique et facilite le test avec des mocks.
- **Couche CLI mince** : `cli.py` ne contient aucune logique metier. Il parse les arguments, appelle le manager correspondant, et affiche le resultat via Rich.
- **Exceptions typees** : chaque erreur a sa propre classe d'exception avec des attributs metier (path, name, hresult, etc.) pour un diagnostic precis.

### 3.3 Technology Stack

| Couche          | Technologie                                        | Version   |
| --------------- | -------------------------------------------------- | --------- |
| Language        | Python                                             | >= 3.14   |
| CLI Framework   | Typer                                              | >= 0.12.0 |
| Terminal Output | Rich                                               | >= 13.0.0 |
| COM Bridge      | pywin32                                            | >= 305    |
| Testing         | pytest + pytest-cov + pytest-mock + pytest-timeout | derniere  |
| Linting         | ruff                                               | derniere  |
| Type Checking   | mypy                                               | derniere  |
| Documentation   | Sphinx + sphinx-rtd-theme                          | derniere  |
| Package Manager | Poetry                                             | derniere  |

## 4. Component Architecture

### 4.1 `exceptions.py` - Hierarchie d'exceptions

**Fichier :** `src/xlmanage/exceptions.py`

**Responsabilite :** Definir toutes les exceptions specifiques au projet. Chaque exception porte des attributs metier pour faciliter le diagnostic.

**Existant :**

```python
class ExcelManageError(Exception):             # Base de toutes les exceptions
class FileNotFoundError(ExcelManageError):      # filepath: Path
class EpicNotFoundError(ExcelManageError):      # epic_name: str
class StoryNotFoundError(ExcelManageError):     # epic_name: str, story_name: str
class TaskNotFoundError(ExcelManageError):      # task_pattern: str
class DuplicateEpicError(ExcelManageError):     # epic_name: str
class DuplicateStoryError(ExcelManageError):    # epic_name: str, story_name: str
```

**A creer (Epic 5, Story 1) - Exceptions COM :**

```python
class ExcelConnectionError(ExcelManageError):
    """Echec de connexion COM (Excel non installe, serveur COM indisponible)."""
    def __init__(self, hresult: int, message: str = "Excel connection failed"):
        self.hresult = hresult   # ex: 0x80080005
        self.message = message

class ExcelInstanceNotFoundError(ExcelManageError):
    """Instance demandee introuvable (pour stop/status)."""
    def __init__(self, instance_id: str, message: str = "Instance not found"):
        self.instance_id = instance_id
        self.message = message

class ExcelRPCError(ExcelManageError):
    """Erreur RPC - serveur COM deconnecte ou indisponible."""
    def __init__(self, hresult: int, message: str = "RPC error"):
        self.hresult = hresult   # ex: 0x800706BE, 0x80010108
        self.message = message
```

**A creer (Epic 6, Story 1) - Exceptions Workbook :**

```python
class WorkbookNotFoundError(ExcelManageError):
    """Classeur introuvable sur le disque."""
    def __init__(self, path: Path, message: str = "Workbook not found"):
        self.path = path
        self.message = message

class WorkbookAlreadyOpenError(ExcelManageError):
    """Classeur deja ouvert dans l'instance Excel."""
    def __init__(self, path: Path, name: str, message: str = "Workbook already open"):
        self.path = path
        self.name = name
        self.message = message

class WorkbookSaveError(ExcelManageError):
    """Echec de sauvegarde (permissions, chemin invalide, format incompatible)."""
    def __init__(self, path: Path, hresult: int = 0, message: str = "Save failed"):
        self.path = path
        self.hresult = hresult
        self.message = message
```

**A creer (Epic 7, Story 1) - Exceptions Worksheet :**

```python
class WorksheetNotFoundError(ExcelManageError):
    """Feuille introuvable dans le classeur."""
    def __init__(self, name: str, workbook_name: str):
        self.name = name
        self.workbook_name = workbook_name

class WorksheetAlreadyExistsError(ExcelManageError):
    """Nom de feuille deja utilise."""
    def __init__(self, name: str, workbook_name: str):
        self.name = name
        self.workbook_name = workbook_name

class WorksheetDeleteError(ExcelManageError):
    """Suppression impossible (derniere feuille visible, protegee)."""
    def __init__(self, name: str, reason: str):
        self.name = name
        self.reason = reason

class WorksheetNameError(ExcelManageError):
    """Nom de feuille invalide (trop long, caracteres interdits)."""
    def __init__(self, name: str, reason: str):
        self.name = name
        self.reason = reason
```

**A creer (Epic 8, Story 1) - Exceptions Table :**

```python
class TableNotFoundError(ExcelManageError):
    """Table introuvable dans la feuille."""
    def __init__(self, name: str, worksheet_name: str):
        self.name = name
        self.worksheet_name = worksheet_name

class TableAlreadyExistsError(ExcelManageError):
    """Nom de table deja utilise dans le classeur."""
    def __init__(self, name: str, workbook_name: str):
        self.name = name
        self.workbook_name = workbook_name

class TableRangeError(ExcelManageError):
    """Plage invalide (syntaxe, plage vide, chevauchement)."""
    def __init__(self, range_ref: str, reason: str):
        self.range_ref = range_ref
        self.reason = reason

class TableNameError(ExcelManageError):
    """Nom de table invalide."""
    def __init__(self, name: str, reason: str):
        self.name = name
        self.reason = reason
```

**A creer (Epic 9, Story 1) - Exceptions VBA :**

```python
class VBAProjectAccessError(ExcelManageError):
    """Acces au projet VBA refuse (Trust Center)."""
    def __init__(self, workbook_name: str):
        self.workbook_name = workbook_name

class VBAModuleNotFoundError(ExcelManageError):
    """Module introuvable dans le projet VBA."""
    def __init__(self, module_name: str, workbook_name: str):
        self.module_name = module_name
        self.workbook_name = workbook_name

class VBAModuleAlreadyExistsError(ExcelManageError):
    """Module avec ce nom existe deja."""
    def __init__(self, module_name: str, workbook_name: str):
        self.module_name = module_name
        self.workbook_name = workbook_name

class VBAImportError(ExcelManageError):
    """Echec d'import (fichier invalide, encodage incorrect)."""
    def __init__(self, module_file: str, reason: str):
        self.module_file = module_file
        self.reason = reason

class VBAExportError(ExcelManageError):
    """Echec d'export (permissions, chemin invalide)."""
    def __init__(self, module_name: str, output_path: str, reason: str):
        self.module_name = module_name
        self.output_path = output_path
        self.reason = reason

class VBAMacroError(ExcelManageError):
    """Echec d'execution de macro."""
    def __init__(self, macro_name: str, reason: str):
        self.macro_name = macro_name
        self.reason = reason

class VBAWorkbookFormatError(ExcelManageError):
    """Classeur au format .xlsx ne supportant pas les macros."""
    def __init__(self, workbook_name: str):
        self.workbook_name = workbook_name
```

**Toutes les nouvelles exceptions doivent etre ajoutees dans `__init__.py` et `__all__`.**

---

### 4.2 `excel_manager.py` - Gestion du cycle de vie Excel (Epic 5)

**Fichier a creer :** `src/xlmanage/excel_manager.py`

**Responsabilite :** Demarrer, arreter, enumerer les instances Excel. C'est le composant central dont tous les autres managers dependent.

**Dataclass :**

```python
from dataclasses import dataclass

@dataclass
class InstanceInfo:
    """Informations sur une instance Excel en cours d'execution."""
    pid: int                 # Process ID du process EXCEL.EXE
    visible: bool            # True si l'instance est visible a l'ecran
    workbooks_count: int     # Nombre de classeurs ouverts
    hwnd: int                # Handle de fenetre Windows (pour identification unique)
```

**Classe principale :**

```python
class ExcelManager:
    """Gestionnaire du cycle de vie d'une instance Excel.

    Implemente le pattern RAII via context manager.
    JAMAIS d'appel a app.Quit() - voir protocole d'arret dans stop().
    """

    def __init__(self, visible: bool = False):
        """
        Args:
            visible: Si True, l'instance sera visible a l'ecran.
                     Par defaut False (mode automatise).
        """
        self._app: CDispatch | None = None
        self._visible: bool = visible

    def __enter__(self) -> "ExcelManager": ...
    def __exit__(self, exc_type, exc_val, exc_tb) -> None: ...

    @property
    def app(self) -> CDispatch:
        """Retourne l'objet COM Application. Raise si non demarre."""
        ...

    # --- DEMARRAGE (Epic 5, Story 2) ---

    def start(self, new: bool = False) -> InstanceInfo:
        """Demarre ou se connecte a une instance Excel.

        Args:
            new: Si False, win32.Dispatch() reutilise une instance via le ROT.
                 Si True, win32.DispatchEx() cree un process isole.

        Returns:
            InstanceInfo avec les infos de l'instance connectee.

        Raises:
            ExcelConnectionError: Si Excel n'est pas installe ou COM indisponible.
        """
        ...

    def get_running_instance(self) -> InstanceInfo | None:
        """Recupere l'instance Excel active via GetActiveObject.

        Returns:
            InstanceInfo si une instance existe, None sinon.
        """
        ...

    def get_instance_info(self, app: CDispatch) -> InstanceInfo:
        """Lit les infos d'une instance Excel.

        Utilise app.Hwnd + ctypes.windll.user32.GetWindowThreadProcessId
        pour recuperer le PID.

        Args:
            app: Objet COM Excel.Application

        Returns:
            InstanceInfo rempli.
        """
        ...

    # --- ARRET (Epic 11, Stories 1-2) ---

    def stop(self, save: bool = True) -> None:
        """Arret propre de l'instance geree.

        Protocole :
        1. app.DisplayAlerts = False
        2. Pour chaque wb dans app.Workbooks : wb.Close(SaveChanges=save)
        3. del wb (pour chaque reference)
        4. del self._app
        5. gc.collect()
        6. self._app = None

        Args:
            save: Si True, sauvegarde chaque classeur avant fermeture.

        IMPORTANT : JAMAIS d'appel a app.Quit().
        """
        ...

    def stop_instance(self, pid: int, save: bool = True) -> None:
        """Arret d'une instance identifiee par PID.

        Se connecte via ROT ou HWND, puis applique le protocole stop().

        Args:
            pid: Process ID de l'instance Excel cible.
            save: Si True, sauvegarde avant fermeture.

        Raises:
            ExcelInstanceNotFoundError: Si le PID n'existe pas.
            ExcelRPCError: Si l'instance est deconnectee.
        """
        ...

    def stop_all(self, save: bool = True) -> list[int]:
        """Arret de toutes les instances Excel.

        Enumere via ROT, applique stop_instance() pour chacune.

        Args:
            save: Si True, sauvegarde avant fermeture.

        Returns:
            Liste des PIDs arretes.
        """
        ...

    def force_kill(self, pid: int) -> None:
        """Arret force via taskkill /f /pid.

        Utilise subprocess.run(["taskkill", "/f", "/pid", str(pid)]).
        A utiliser UNIQUEMENT quand l'arret propre echoue.

        Args:
            pid: Process ID a tuer.

        Raises:
            ExcelInstanceNotFoundError: Si le PID n'existe pas.
        """
        ...

    def list_running_instances(self) -> list[InstanceInfo]:
        """Enumere toutes les instances Excel actives.

        Parcourt le ROT via pythoncom.GetRunningObjectTable().
        Fallback : tasklist si ROT inaccessible.

        Returns:
            Liste d'InstanceInfo pour chaque instance trouvee.
        """
        ...
```

**Fonctions utilitaires du module (Epic 11, Story 2) :**

```python
def enumerate_excel_instances() -> list[tuple[CDispatch, InstanceInfo]]:
    """Enumeration via le Running Object Table (ROT).

    Parcourt pythoncom.GetRunningObjectTable(), filtre les monikers
    contenant "Excel.Application", retourne les CDispatch + infos.
    """
    ...

def enumerate_excel_pids() -> list[int]:
    """Fallback : enumeration des PIDs via tasklist.

    Utilise : subprocess.run(["tasklist", "/fi",
              "imagename eq EXCEL.EXE", "/fo", "csv", "/nh"])
    Parse le CSV pour extraire les PIDs.
    """
    ...

def connect_by_hwnd(hwnd: int) -> CDispatch | None:
    """Connexion a une instance Excel par son handle de fenetre.

    Utilise ctypes.windll.oleacc.AccessibleObjectFromWindow.
    Fallback si l'instance n'est pas dans le ROT.
    """
    ...
```

**Dependencies :** `pywin32`, `pythoncom`, `ctypes`, `subprocess`, `gc`

---

### 4.3 `workbook_manager.py` - Gestion des classeurs (Epic 6)

**Fichier a creer :** `src/xlmanage/workbook_manager.py`

**Responsabilite :** Operations CRUD sur les classeurs Excel (open, create, close, save, list).

**Dataclass :**

```python
@dataclass
class WorkbookInfo:
    """Informations sur un classeur Excel."""
    name: str              # Nom du fichier (ex: "data.xlsx")
    full_path: Path        # Chemin complet (ex: C:\Users\...\data.xlsx)
    read_only: bool        # True si ouvert en lecture seule
    saved: bool            # True si toutes les modifications sont sauvegardees
    sheets_count: int      # Nombre de feuilles dans le classeur
```

**Constante : mapping extension -> FileFormat :**

```python
FILE_FORMAT_MAP: dict[str, int] = {
    ".xlsx": 51,   # xlOpenXMLWorkbook
    ".xlsm": 52,   # xlOpenXMLWorkbookMacroEnabled
    ".xls":  56,   # xlExcel8 (format 97-2003)
    ".xlsb": 50,   # xlExcel12 (binaire)
}
```

**Classe principale :**

```python
class WorkbookManager:
    """Gestionnaire CRUD des classeurs Excel.

    Depend de ExcelManager pour l'acces a app.
    """

    def __init__(self, excel_manager: ExcelManager):
        """
        Args:
            excel_manager: Instance ExcelManager deja demarree.
        """
        self._mgr = excel_manager

    def open(self, path: Path, read_only: bool = False) -> WorkbookInfo:
        """Ouvre un classeur existant.

        Etapes :
        1. Verifier existence fichier avec Path.exists()
        2. Verifier si deja ouvert (iteration app.Workbooks sur FullName)
        3. app.Workbooks.Open(str(path.resolve()), ReadOnly=read_only)

        Args:
            path: Chemin du fichier Excel.
            read_only: Si True, ouvre en lecture seule.

        Returns:
            WorkbookInfo du classeur ouvert.

        Raises:
            WorkbookNotFoundError: Si le fichier n'existe pas sur le disque.
            WorkbookAlreadyOpenError: Si le classeur est deja ouvert.
            ExcelConnectionError: Si la connexion COM echoue.
        """
        ...

    def create(self, path: Path, template: Path | None = None) -> WorkbookInfo:
        """Cree un nouveau classeur.

        Sans template : app.Workbooks.Add()
        Avec template : app.Workbooks.Add(str(template.resolve()))
        Puis : wb.SaveAs(str(path.resolve()), FileFormat)

        Args:
            path: Chemin de destination pour le nouveau classeur.
            template: Chemin vers un modele (optionnel).

        Returns:
            WorkbookInfo du classeur cree.

        Raises:
            WorkbookNotFoundError: Si le template n'existe pas.
            WorkbookSaveError: Si la sauvegarde initiale echoue.
        """
        ...

    def close(self, path: Path, save: bool = True, force: bool = False) -> None:
        """Ferme un classeur ouvert.

        Etapes :
        1. Trouver le classeur via _find_open_workbook()
        2. Si force : app.DisplayAlerts = False
        3. wb.Close(SaveChanges=save)
        4. del wb

        Args:
            path: Chemin du classeur a fermer.
            save: Si True, sauvegarde avant fermeture.
            force: Si True, supprime les dialogues de confirmation.
        """
        ...

    def save(self, path: Path, output: Path | None = None) -> None:
        """Sauvegarde un classeur.

        Sans output : wb.Save()
        Avec output : wb.SaveAs(str(output.resolve()), FileFormat)

        Args:
            path: Chemin du classeur a sauvegarder.
            output: Chemin de destination pour SaveAs (optionnel).

        Raises:
            WorkbookSaveError: Si la sauvegarde echoue.
        """
        ...

    def list(self) -> list[WorkbookInfo]:
        """Liste tous les classeurs ouverts.

        Itere app.Workbooks, extrait Name, FullName, ReadOnly, Saved,
        Worksheets.Count.

        Returns:
            Liste de WorkbookInfo.
        """
        ...
```

**Fonctions utilitaires du module :**

```python
def _detect_file_format(path: Path) -> int:
    """Retourne le code FileFormat Excel depuis l'extension du fichier.

    Args:
        path: Chemin du fichier (.xlsx, .xlsm, .xls, .xlsb)

    Returns:
        Code FileFormat (51, 52, 56, 50)

    Raises:
        ValueError: Si l'extension n'est pas reconnue.
    """
    ...

def _find_open_workbook(app: CDispatch, path: Path) -> CDispatch | None:
    """Recherche un classeur ouvert par FullName puis par Name.

    Args:
        app: Objet COM Excel.Application
        path: Chemin du classeur recherche.

    Returns:
        Objet COM Workbook si trouve, None sinon.
    """
    ...
```

**Dependencies :** `ExcelManager`, `pathlib.Path`

---

### 4.4 `worksheet_manager.py` - Gestion des feuilles (Epic 7)

**Fichier a creer :** `src/xlmanage/worksheet_manager.py`

**Responsabilite :** Operations CRUD sur les feuilles de calcul (create, delete, list, copy).

**Dataclass :**

```python
@dataclass
class WorksheetInfo:
    """Informations sur une feuille de calcul."""
    name: str             # Nom de la feuille (ex: "Feuille1")
    index: int            # Position dans le classeur (1-based dans Excel)
    visible: bool         # True si la feuille est visible
    rows_used: int        # Nombre de lignes contenant des donnees
    columns_used: int     # Nombre de colonnes contenant des donnees
```

**Constantes de validation :**

```python
SHEET_NAME_MAX_LENGTH: int = 31
SHEET_NAME_FORBIDDEN_CHARS: str = r'\/*?:[]'
```

**Classe principale :**

```python
class WorksheetManager:
    """Gestionnaire CRUD des feuilles de calcul.

    Depend de ExcelManager pour l'acces a app.
    """

    def __init__(self, excel_manager: ExcelManager):
        self._mgr = excel_manager

    def create(self, name: str, workbook: Path | None = None) -> WorksheetInfo:
        """Cree une nouvelle feuille en derniere position.

        Etapes :
        1. _validate_sheet_name(name)
        2. _find_worksheet(wb, name) -> raise si existe deja
        3. wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        4. ws.Name = name

        Args:
            name: Nom de la nouvelle feuille (max 31 chars, pas de \\/*?:[])
            workbook: Classeur cible. Si None, utilise le classeur actif.

        Returns:
            WorksheetInfo de la feuille creee.

        Raises:
            WorksheetNameError: Si le nom est invalide.
            WorksheetAlreadyExistsError: Si le nom est deja utilise.
        """
        ...

    def delete(self, name: str, workbook: Path | None = None,
               force: bool = False) -> None:
        """Supprime une feuille.

        Etapes :
        1. _find_worksheet(wb, name) -> raise si introuvable
        2. Verifier que ce n'est PAS la derniere feuille visible
        3. app.DisplayAlerts = False (OBLIGATOIRE pour eviter le dialogue)
        4. ws.Delete()
        5. Restaurer app.DisplayAlerts
        6. del ws

        Args:
            name: Nom de la feuille a supprimer.
            workbook: Classeur cible.
            force: Non utilise ici (DisplayAlerts est toujours desactive).

        Raises:
            WorksheetNotFoundError: Si la feuille n'existe pas.
            WorksheetDeleteError: Si c'est la derniere feuille visible.
        """
        ...

    def list(self, workbook: Path | None = None) -> list[WorksheetInfo]:
        """Liste toutes les feuilles du classeur.

        Itere wb.Worksheets, extrait Name, index, Visible,
        UsedRange.Rows.Count, UsedRange.Columns.Count.

        Args:
            workbook: Classeur cible. Si None, utilise le classeur actif.

        Returns:
            Liste de WorksheetInfo.
        """
        ...

    def copy(self, source: str, destination: str,
             workbook: Path | None = None) -> WorksheetInfo:
        """Copie une feuille et renomme la copie.

        Etapes :
        1. _find_worksheet(wb, source) -> raise si introuvable
        2. _validate_sheet_name(destination)
        3. Verifier unicite destination
        4. ws_source.Copy(After=ws_source)
        5. wb.ActiveSheet.Name = destination (la copie est auto-activee)

        Args:
            source: Nom de la feuille source.
            destination: Nom de la copie.
            workbook: Classeur cible.

        Returns:
            WorksheetInfo de la copie.

        Raises:
            WorksheetNotFoundError: Si la source n'existe pas.
            WorksheetNameError: Si le nom destination est invalide.
            WorksheetAlreadyExistsError: Si le nom destination existe deja.
        """
        ...
```

**Fonctions utilitaires du module :**

```python
def _resolve_workbook(app: CDispatch, workbook: Path | None) -> CDispatch:
    """Resout le classeur cible.

    Si workbook fourni : cherche via _find_open_workbook() ou ouvre.
    Si None : retourne app.ActiveWorkbook (raise si aucun classeur actif).

    Reutilisee par WorksheetManager, TableManager, VBAManager.
    """
    ...

def _validate_sheet_name(name: str) -> None:
    """Valide un nom de feuille Excel.

    Regles :
    - Longueur <= 31 caracteres
    - Aucun caractere parmi : \\ / * ? : [ ]

    Raises:
        WorksheetNameError: Si le nom est invalide (avec raison detaillee).
    """
    ...

def _find_worksheet(wb: CDispatch, name: str) -> CDispatch | None:
    """Recherche une feuille par nom (case-insensitive).

    Itere wb.Worksheets, compare ws.Name.lower() == name.lower().

    Returns:
        Objet COM Worksheet si trouve, None sinon.
    """
    ...
```

**Dependencies :** `ExcelManager`, `WorkbookManager._find_open_workbook()`

---

### 4.5 `table_manager.py` - Gestion des tables ListObject (Epic 8)

**Fichier a creer :** `src/xlmanage/table_manager.py`

**Responsabilite :** Operations CRUD sur les tables Excel (ListObject).

**Dataclass :**

```python
@dataclass
class TableInfo:
    """Informations sur une table ListObject."""
    name: str                # Nom de la table (ex: "tbVentes")
    worksheet_name: str      # Nom de la feuille parente
    range_address: str       # Adresse de la plage (ex: "$A$1:$D$10")
    columns: list[str]       # Noms des colonnes d'en-tete
    rows_count: int          # Nombre de lignes de donnees (hors en-tetes)
    header_row: str          # Adresse de la ligne d'en-tete (ex: "$A$1:$D$1")
```

**Constantes de validation :**

```python
TABLE_NAME_MAX_LENGTH: int = 255
# Regex : commence par lettre ou _, suivi de lettres/chiffres/_
TABLE_NAME_PATTERN: str = r'^[A-Za-z_][A-Za-z0-9_]*$'
```

**Classe principale :**

```python
class TableManager:
    """Gestionnaire CRUD des tables ListObject.

    Depend de ExcelManager pour l'acces a app.
    """

    def __init__(self, excel_manager: ExcelManager):
        self._mgr = excel_manager

    def create(self, name: str, range_ref: str,
               workbook: Path | None = None,
               worksheet: str | None = None) -> TableInfo:
        """Cree une nouvelle table ListObject.

        Etapes :
        1. _validate_table_name(name)
        2. _find_table(wb, name) -> raise si existe deja (unicite dans le classeur)
        3. Resoudre la feuille (worksheet ou ActiveSheet)
        4. _validate_range(ws, range_ref)
        5. ws.ListObjects.Add(SourceType=1, Source=range_obj,
                              XlListObjectHasHeaders=1)
        6. lo.Name = name

        Args:
            name: Nom de la table (pas d'espaces, pas de chiffre en tete).
            range_ref: Plage de cellules (ex: "A1:D10").
            workbook: Classeur cible (optionnel).
            worksheet: Feuille cible (optionnel).

        Returns:
            TableInfo de la table creee.

        Raises:
            TableNameError: Nom invalide.
            TableAlreadyExistsError: Nom deja utilise dans le classeur.
            TableRangeError: Plage invalide ou chevauchante.
        """
        ...

    def delete(self, name: str, workbook: Path | None = None,
               worksheet: str | None = None, force: bool = False) -> None:
        """Supprime une table.

        force=False : lo.Unlist() -> conserve les donnees, supprime la structure
        force=True  : lo.Delete() -> supprime table ET donnees

        Args:
            name: Nom de la table.
            workbook: Classeur cible (optionnel).
            worksheet: Feuille cible (optionnel).
            force: Si True, supprime aussi les donnees.

        Raises:
            TableNotFoundError: Table introuvable.
        """
        ...

    def list(self, workbook: Path | None = None,
             worksheet: str | None = None) -> list[TableInfo]:
        """Liste les tables.

        Si worksheet specifie : itere ws.ListObjects uniquement.
        Si absent : itere toutes les feuilles du classeur.

        Extrait : Name, Range.Address, ListColumns (noms), ListRows.Count,
                  HeaderRowRange.Address.

        Returns:
            Liste de TableInfo.
        """
        ...
```

**Fonctions utilitaires du module :**

```python
def _validate_table_name(name: str) -> None:
    """Valide un nom de table Excel.

    Regles :
    - Pas d'espaces
    - Pas de caracteres speciaux (sauf _)
    - Ne commence pas par un chiffre
    - Max 255 caracteres
    - Match TABLE_NAME_PATTERN

    Raises:
        TableNameError: Si le nom est invalide.
    """
    ...

def _find_table(wb: CDispatch, name: str) -> tuple[CDispatch, CDispatch] | None:
    """Recherche une table dans tout le classeur (noms uniques au classeur).

    Itere toutes les feuilles, puis tous les ListObjects de chaque feuille.

    Returns:
        (worksheet, listobject) si trouvee, None sinon.
    """
    ...

def _validate_range(ws: CDispatch, range_ref: str) -> CDispatch:
    """Valide et retourne un objet Range COM.

    Verifie :
    1. ws.Range(range_ref) ne raise pas (syntaxe valide)
    2. La plage ne chevauche pas un ListObject existant

    Returns:
        Objet COM Range.

    Raises:
        TableRangeError: Si la plage est invalide ou chevauche une table.
    """
    ...
```

**Dependencies :** `ExcelManager`, `_resolve_workbook()`, `_find_worksheet()`

---

### 4.6 `vba_manager.py` - Gestion des modules VBA (Epic 9)

**Fichier a creer :** `src/xlmanage/vba_manager.py`

**Responsabilite :** Import, export, listage et suppression de modules VBA (standard, classe, UserForm).

**Constantes :**

```python
# Types de composants VBA (constantes Excel)
VBEXT_CT_STD_MODULE: int = 1    # Module standard (.bas)
VBEXT_CT_CLASS_MODULE: int = 2  # Module de classe (.cls)
VBEXT_CT_MS_FORM: int = 3       # UserForm (.frm + .frx)
VBEXT_CT_DOCUMENT: int = 100    # Module de document (ThisWorkbook, Sheet1)

# Mapping type code -> nom lisible
VBA_TYPE_NAMES: dict[int, str] = {
    1: "standard",
    2: "class",
    3: "userform",
    100: "document",
}

# Mapping extension -> type attendu
EXTENSION_TO_TYPE: dict[str, str] = {
    ".bas": "standard",
    ".cls": "class",
    ".frm": "userform",
}

# Encodage obligatoire pour les fichiers VBA
VBA_ENCODING: str = "windows-1252"
```

**Dataclass :**

```python
@dataclass
class VBAModuleInfo:
    """Informations sur un module VBA."""
    name: str                 # Nom du module (ex: "Module1")
    module_type: str          # "standard", "class", "userform", "document"
    lines_count: int          # Nombre de lignes de code
    has_predeclared_id: bool  # True si PredeclaredId (classes uniquement)
```

**Classe principale :**

```python
class VBAManager:
    """Gestionnaire des modules VBA.

    Depend de ExcelManager pour l'acces a app.
    IMPORTANT : le classeur doit etre au format .xlsm pour supporter les macros.
    IMPORTANT : "Trust access to the VBA project object model" doit etre active
    dans Excel (File > Options > Trust Center > Macro Settings).
    """

    def __init__(self, excel_manager: ExcelManager):
        self._mgr = excel_manager

    def import_module(self, module_file: Path,
                      module_type: str | None = None,
                      workbook: Path | None = None,
                      overwrite: bool = False) -> VBAModuleInfo:
        """Importe un module VBA depuis un fichier.

        Module standard (.bas) :
            VBComponents.Import(str(path.resolve()))

        Module de classe (.cls) - process en 5 etapes :
            1. Lire le fichier avec encoding='windows-1252'
            2. Extraire VB_Name depuis 'Attribute VB_Name = "..."'
            3. Extraire VB_PredeclaredId (True/False)
            4. Striper l'en-tete (tout avant Option Explicit)
            5. VBComponents.Add(2) -> nommer -> Properties("PredeclaredId")
               -> CodeModule.AddFromString()

        UserForm (.frm) :
            VBComponents.Import(str(path.resolve()))
            Le fichier .frx doit etre dans le meme repertoire.

        Args:
            module_file: Chemin du fichier .bas, .cls ou .frm.
            module_type: Type force (auto-detection depuis l'extension si None).
            workbook: Classeur cible (optionnel).
            overwrite: Si True, supprime le module existant avant import.

        Returns:
            VBAModuleInfo du module importe.

        Raises:
            VBAImportError: Fichier invalide ou encodage incorrect.
            VBAModuleAlreadyExistsError: Module existe et overwrite=False.
            VBAProjectAccessError: Trust Center refuse l'acces.
            VBAWorkbookFormatError: Classeur .xlsx (pas de support VBA).
        """
        ...

    def export_module(self, module_name: str, output_file: Path,
                      workbook: Path | None = None) -> Path:
        """Exporte un module VBA vers un fichier.

        Modules standard/classe/UserForm : component.Export(str(output_file.resolve()))
        Modules de document (Type=100) : export manuel via
            CodeModule.Lines(1, count) car Export() n'est pas supporte.

        Args:
            module_name: Nom du module dans le projet VBA.
            output_file: Chemin de destination.
            workbook: Classeur source (optionnel).

        Returns:
            Chemin effectif du fichier exporte.

        Raises:
            VBAModuleNotFoundError: Module introuvable.
            VBAExportError: Echec d'ecriture.
        """
        ...

    def list_modules(self, workbook: Path | None = None) -> list[VBAModuleInfo]:
        """Liste tous les modules VBA du classeur.

        Itere VBComponents, extrait Name, Type (mappe vers nom lisible),
        CodeModule.CountOfLines.

        Returns:
            Liste de VBAModuleInfo.

        Raises:
            VBAProjectAccessError: Trust Center refuse l'acces.
        """
        ...

    def delete_module(self, module_name: str,
                      workbook: Path | None = None,
                      force: bool = False) -> None:
        """Supprime un module VBA.

        Les modules de document (Type=100 : ThisWorkbook, Sheet1, etc.)
        ne sont PAS supprimables -> raise VBAModuleNotFoundError avec message.

        Appel : VBComponents.Remove(component) puis del component.

        Args:
            module_name: Nom du module a supprimer.
            workbook: Classeur cible (optionnel).
            force: Pas de dialogue de confirmation.

        Raises:
            VBAModuleNotFoundError: Module introuvable ou non supprimable.
        """
        ...
```

**Fonctions utilitaires du module :**

```python
def _get_vba_project(wb: CDispatch) -> CDispatch:
    """Acces au VBProject avec gestion d'erreur.

    Catch pywintypes.com_error 0x800A03EC -> raise VBAProjectAccessError.
    Verifie que le classeur est au format .xlsm.

    Returns:
        Objet COM VBProject.
    """
    ...

def _find_component(vb_project: CDispatch, name: str) -> CDispatch | None:
    """Recherche un composant VBA par nom."""
    ...

def _detect_module_type(path: Path) -> str:
    """Detecte le type de module depuis l'extension.

    .bas -> "standard", .cls -> "class", .frm -> "userform"

    Raises:
        VBAImportError: Extension non reconnue.
    """
    ...
```

**Dependencies :** `ExcelManager`, `_resolve_workbook()`

---

### 4.7 `macro_runner.py` - Execution de macros (Epic 12)

**Fichier a creer :** `src/xlmanage/macro_runner.py`

**Responsabilite :** Execution de macros VBA (Sub et Function) avec parsing des arguments et gestion des retours.

**Dataclass :**

```python
@dataclass
class MacroResult:
    """Resultat d'execution d'une macro VBA."""
    macro_name: str                # Nom complet de la macro executee
    return_value: object | None    # Valeur brute retournee par app.Run
    return_type: str               # Type Python du retour ("str", "int", "None", etc.)
    success: bool                  # True si execution sans erreur
    error_message: str | None      # Message VBA si erreur (extrait de excepinfo[2])
```

**Classe principale :**

```python
class MacroRunner:
    """Executeur de macros VBA.

    Depend de ExcelManager pour l'acces a app.
    """

    def __init__(self, excel_manager: ExcelManager):
        self._mgr = excel_manager

    def run(self, macro_name: str,
            workbook: Path | None = None,
            args: str | None = None) -> MacroResult:
        """Execute une macro VBA.

        Etapes :
        1. _build_macro_reference() pour construire la reference complete
        2. _parse_macro_args() si args fournis
        3. app.Run(full_ref, *parsed_args)
        4. Encapsuler le resultat dans MacroResult

        Gestion d'erreur :
        - pywintypes.com_error avec HRESULT 0x800A03EC ou 0x80020009
        - Le message d'erreur VBA est dans excepinfo[2]

        Args:
            macro_name: Nom de la macro (ex: "Module1.MySub").
            workbook: Classeur contenant la macro (doit etre ouvert).
            args: Arguments CSV (ex: '"hello",42,3.14,true').

        Returns:
            MacroResult avec la valeur de retour et le statut.

        Raises:
            VBAMacroError: Macro introuvable ou erreur VBA runtime.
            WorkbookNotFoundError: Classeur non ouvert.
        """
        ...
```

**Fonctions utilitaires du module :**

```python
def _build_macro_reference(macro_name: str, workbook: Path | None,
                           app: CDispatch) -> str:
    """Construit la reference de macro complete.

    Sans workbook : retourne macro_name tel quel.
    Avec workbook : "'WorkbookName.xlsm'!macro_name"
    (guillemets simples si espaces ou points dans le nom)

    Raises:
        WorkbookNotFoundError: Si le classeur n'est pas ouvert.
    """
    ...

def _parse_macro_args(args_str: str) -> list[str | int | float | bool]:
    """Parse une chaine CSV en liste d'arguments types.

    Regles de conversion par priorite :
    1. "..." ou '...' -> str (sans les guillemets)
    2. true / false (case-insensitive) -> bool
    3. Contient '.' et parseable -> float
    4. Entierement numerique (avec signe optionnel) -> int
    5. Sinon -> str

    Gere les virgules dans les chaines entre guillemets.

    Raises:
        VBAMacroError: Si > 30 arguments (limite COM).

    Exemples :
        '"hello, world",42,3.14,true'
        -> ["hello, world", 42, 3.14, True]
    """
    ...

def _format_return_value(value: object) -> str:
    """Formate une valeur de retour VBA pour affichage.

    None -> "(aucune valeur de retour)"
    pywintypes.datetime -> format ISO 8601
    tuple de tuple (tableau VBA) -> representation tabulaire
    Autres -> str(value)
    """
    ...
```

**Dependencies :** `ExcelManager`

---

### 4.8 Optimizers existants (Epic 10 - refactorisation)

**Fichiers existants :**

- `src/xlmanage/excel_optimizer.py` - 8 proprietes (complet)
- `src/xlmanage/screen_optimizer.py` - 3 proprietes (ecran)
- `src/xlmanage/calculation_optimizer.py` - 4 proprietes (calcul)

**Etat actuel :** Les 3 classes fonctionnent en mode context manager (`with`). Elles creent chacune leur propre instance Excel si aucune n'est fournie.

**Modifications a apporter (Epic 10) :**

Ajouter a chaque optimizer les methodes `apply()` et `restore()` pour un usage hors context manager (les optimisations persistent apres l'appel CLI) :

```python
# Ajouts a ExcelOptimizer (meme pattern pour ScreenOptimizer et CalculationOptimizer)

def apply(self) -> None:
    """Applique les optimisations SANS context manager.

    Sauvegarde les parametres actuels puis applique les optimisations.
    Les optimisations persistent jusqu'a un appel a restore().
    """
    self._save_current_settings()
    self._apply_optimizations()

def restore(self) -> None:
    """Restaure les parametres sauvegardes par apply()."""
    self._restore_original_settings()

def get_current_settings(self) -> dict[str, object]:
    """Retourne l'etat actuel des proprietes Excel.

    Returns:
        Dictionnaire {nom_propriete: valeur_actuelle}
    """
    ...
```

**Nouvelle dataclass :**

```python
@dataclass
class OptimizationState:
    """Etat des optimisations pour tracking et restauration."""
    screen: dict[str, object]       # {ScreenUpdating, DisplayStatusBar, EnableAnimations}
    calculation: dict[str, object]  # {Calculation, Iteration, MaxIterations, MaxChange}
    full: dict[str, object]         # {EnableEvents, DisplayAlerts, AskToUpdateLinks, ...}
    applied_at: str                 # Timestamp ISO de l'application
    optimizer_type: str             # "screen", "calculation", "all"
```

**Le context manager `__enter__` / `__exit__` existant reste inchange** pour la compatibilite.

**Dependencies :** `ExcelManager` (au lieu de `gencache.EnsureDispatch` autonome)

---

### 4.9 `cli.py` - Interface en ligne de commande

**Fichier existant :** `src/xlmanage/cli.py`

**Responsabilite :** Point d'entree utilisateur. Parse les arguments, appelle les managers, affiche les resultats via Rich. **Aucune logique metier dans ce fichier.**

**Etat actuel :** Toutes les commandes sont des stubs (print uniquement). Le CLI est complet en termes de structure : 5 commandes principales + 4 sous-groupes (workbook, worksheet, vba, table).

**Pattern d'integration pour chaque commande :**

```python
@app.command()
def start(visible: bool, new: bool) -> None:
    """Exemple de pattern d'integration."""
    try:
        mgr = ExcelManager(visible=visible)
        info = mgr.start(new=new)
        # Affichage Rich de InstanceInfo
        print(f"[green]Instance demarree : PID={info.pid}, HWND={info.hwnd}[/green]")
    except ExcelConnectionError as e:
        print(f"[red]Erreur: {e.message} (HRESULT: {e.hresult:#010x})[/red]")
        raise typer.Exit(code=1)
```

**Arbre des commandes (reference) :**

```
xlmanage
  start           --visible/--hidden  --new
  stop            [instance_id]  --all  --force  --no-save
  status
  optimize        --screen  --calculation  --all  --restore  --status  --force-calculate
  run-macro       MACRO_NAME  --workbook  --args  --timeout
  workbook
    open          PATH  --read-only
    create        PATH  --template
    close         PATH  --save/--no-save  --force
    save          PATH  --as
    list
  worksheet
    create        NAME  --workbook
    delete        NAME  --workbook  --force
    list          --workbook
    copy          SOURCE  DESTINATION  --workbook
  vba
    import        MODULE_FILE  --type  --workbook  --overwrite
    export        MODULE_NAME  OUTPUT_FILE  --workbook
    list          --workbook
    delete        MODULE_NAME  --workbook  --force
  table
    create        NAME  RANGE_REF  --workbook  --worksheet
    delete        NAME  --workbook  --worksheet  --force
    list          --workbook  --worksheet
```

## 5. Data Architecture

### 5.1 Data Models (resume)

Toutes les dataclasses du projet :

| Dataclass           | Module                 | Champs                                                                                                               |
| ------------------- | ---------------------- | -------------------------------------------------------------------------------------------------------------------- |
| `InstanceInfo`      | `excel_manager.py`     | `pid: int`, `visible: bool`, `workbooks_count: int`, `hwnd: int`                                                     |
| `WorkbookInfo`      | `workbook_manager.py`  | `name: str`, `full_path: Path`, `read_only: bool`, `saved: bool`, `sheets_count: int`                                |
| `WorksheetInfo`     | `worksheet_manager.py` | `name: str`, `index: int`, `visible: bool`, `rows_used: int`, `columns_used: int`                                    |
| `TableInfo`         | `table_manager.py`     | `name: str`, `worksheet_name: str`, `range_address: str`, `columns: list[str]`, `rows_count: int`, `header_row: str` |
| `VBAModuleInfo`     | `vba_manager.py`       | `name: str`, `module_type: str`, `lines_count: int`, `has_predeclared_id: bool`                                      |
| `MacroResult`       | `macro_runner.py`      | `macro_name: str`, `return_value: object\|None`, `return_type: str`, `success: bool`, `error_message: str\|None`     |
| `OptimizationState` | `excel_optimizer.py`   | `screen: dict`, `calculation: dict`, `full: dict`, `applied_at: str`, `optimizer_type: str`                          |

### 5.2 Data Flow

```
Utilisateur (terminal)
    |
    v
cli.py (parse args, validation)
    |
    v
*_manager.py (logique metier, validation metier)
    |
    v
pywin32 COM (Dispatch/DispatchEx)
    |                ^
    v                |
Excel.exe     Retour COM (objets, valeurs, erreurs)
    |                |
    v                v
*_manager.py (construction Dataclass)
    |
    v
cli.py (affichage Rich)
    |
    v
Utilisateur (terminal)
```

**Flux detaille pour `xlmanage workbook open fichier.xlsx` :**

```
1. cli.py : parse "fichier.xlsx" -> Path
2. cli.py : cree ExcelManager, WorkbookManager
3. WorkbookManager.open(Path("fichier.xlsx"))
   3a. Path.exists() -> True ? (sinon raise WorkbookNotFoundError)
   3b. Itere app.Workbooks pour verifier si deja ouvert
   3c. app.Workbooks.Open(str(path.resolve()))
   3d. Construit WorkbookInfo depuis l'objet COM Workbook
4. cli.py : affiche WorkbookInfo via Rich Table
```

### 5.3 Storage

- **Aucune persistance fichier** : xlManage ne stocke pas d'etat sur le disque.
- **Etat en memoire** : les optimisations sauvegardees (valeurs originales) sont stockees dans l'instance de l'optimizer. Si le process Python meurt, l'etat est perdu (acceptable car Excel conserve ses propres defaults au redemarrage).
- **Fichiers Excel** : toutes les donnees sont dans les classeurs Excel geres par le serveur COM.

## 6. Security Architecture

### 6.1 Authentication

Non applicable. xlManage opere en local sur la machine de l'utilisateur, sans reseau ni authentification.

### 6.2 Authorization

- **Trust Center Excel** : l'option "Trust access to the VBA project object model" doit etre activee manuellement par l'utilisateur pour que les commandes `vba` fonctionnent. Sans cette option, `wb.VBProject` raise `pywintypes.com_error` 0x800A03EC.
- **AutomationSecurity** : la propriete `app.AutomationSecurity` controle l'execution des macros (1=Low, 2=ByUI, 3=ForceDisable). xlManage verifie cette valeur avant `run-macro` et affiche un warning si les macros sont desactivees.

### 6.3 Data Protection

- **Pas de credentials stockees** : aucun fichier .env, aucun secret.
- **Chemins resolus** : tous les chemins sont resolus en absolu via `Path.resolve()` avant passage au COM pour eviter les injections de chemin.
- **Pas d'appel a `eval()`** : le parsing des arguments de macros utilise un parser CSV maison, jamais `eval()`.

## 7. Deployment Architecture

### 7.1 Deployment Model

Installation locale via pip/Poetry :

```
pip install xlmanage
# ou
poetry install
```

Le package s'installe comme un script console (`xlmanage` dans le PATH via pyproject.toml `[tool.poetry.scripts]`).

### 7.2 Infrastructure

```
Machine Windows locale
  +-- Python >= 3.14
  +-- pip / Poetry
  +-- xlmanage (package installe)
  +-- Microsoft Excel (installe, avec licence)
  +-- COM runtime Windows (natif)
```

### 7.3 Environments

- **Development** : `poetry install --with dev` (inclut pytest, ruff, mypy, sphinx)
- **Production** : `pip install xlmanage` (uniquement typer, rich, pywin32)

Pas d'environnement staging (outil local, pas de serveur).

## 8. Error Handling and Resilience

### 8.1 Error Handling Strategy

**Principe : les erreurs COM sont traduites en exceptions metier typees.**

```
pywintypes.com_error  --catch-->  Exception xlManage specifique
                                       |
                                       v
                                  cli.py catch -> Rich message + typer.Exit(code=1)
```

**Couches de gestion d'erreur :**

1. **Couche COM** : `pywintypes.com_error` brute avec HRESULT et excepinfo
2. **Couche Manager** : catch COM error, raise exception xlManage avec contexte metier
3. **Couche CLI** : catch exception xlManage, affiche message Rich, exit code 1

**HRESULT courants et leur traduction :**

| HRESULT    | Signification                        | Exception xlManage                         |
| ---------- | ------------------------------------ | ------------------------------------------ |
| 0x80080005 | Serveur COM non disponible           | `ExcelConnectionError`                     |
| 0x800706BE | Serveur deconnecte (RPC)             | `ExcelRPCError`                            |
| 0x80010108 | Objet COM deconnecte                 | `ExcelRPCError`                            |
| 0x800A03EC | Erreur generique Excel / VBA runtime | `VBAMacroError` ou `VBAProjectAccessError` |
| 0x80020009 | Exception COM avec excepinfo         | Extraire `excepinfo[2]` pour le message    |

### 8.2 Resilience Patterns

- **Retry avec backoff** : pour les erreurs RPC transitoires (0x800706BE, 0x80010108), retenter 3 fois avec delai exponentiel avant de raise.
- **Fallback ROT -> tasklist** : si `pythoncom.GetRunningObjectTable()` echoue, utiliser `tasklist /fi "imagename eq EXCEL.EXE"` pour enumerer les processus.
- **Fallback arret propre -> taskkill** : si l'arret via COM echoue (instance zombie), `taskkill /f /pid` en dernier recours avec le flag `--force`.
- **Liberation ordonnee des refs COM** : toujours `del` dans l'ordre inner -> outer (ws avant wb avant app) puis `gc.collect()`. Cela previent les erreurs RPC lors du garbage collection.

### 8.3 Monitoring and Observability

- **Logging Python** : utiliser `logging` standard pour tracer les operations COM (niveau DEBUG pour le detail, WARNING pour les classeurs non sauvegardes, ERROR pour les erreurs COM).
- **Rich output** : messages colores dans le terminal (vert=succes, jaune=warning, rouge=erreur).
- **Exit codes** : 0=succes, 1=erreur.

## 9. Performance Considerations

### 9.1 Performance Requirements

- Les commandes CLI doivent repondre en moins de 2 secondes (hors temps de demarrage d'Excel).
- Le demarrage d'Excel (`DispatchEx`) peut prendre 3-5 secondes au premier appel.
- `Dispatch()` (reutilisation via ROT) est quasi-instantane.

### 9.2 Scalability

Non applicable. xlManage est un outil CLI mono-utilisateur. Chaque thread accedant au COM doit appeler `pythoncom.CoInitialize()` (modele STA).

### 9.3 Resource Usage

- **Memoire Python** : negligeable (< 50 MB)
- **Process Excel** : chaque `DispatchEx` cree un nouveau process EXCEL.EXE (~100 MB RAM)
- **Refs COM** : liberer les references inutilisees rapidement (`del` + `gc.collect()`) pour eviter les fuites de memoire dans Excel

## 10. Testing Strategy

### 10.1 Testing Levels

- **Tests unitaires** : chaque module `*_manager.py` avec mocks COM (`unittest.mock.Mock`). Pas de COM reel dans les tests unitaires.
- **Tests CLI** : `typer.testing.CliRunner` avec mock des managers. Verifie les arguments, les messages Rich et les exit codes.
- **Tests d'integration** : scenarios end-to-end complets avec mocks COM coordonnes (ex: open -> save -> close).

### 10.2 Testing Tools

| Outil                     | Role                                                 |
| ------------------------- | ---------------------------------------------------- |
| `pytest`                  | Framework de test principal                          |
| `pytest-cov`              | Couverture de code (seuil minimum : 90%)             |
| `pytest-mock`             | Injection de mocks via fixture `mocker`              |
| `pytest-timeout`          | Timeout 60s par test (prevenir les blocages COM)     |
| `unittest.mock.Mock`      | Mock des objets COM (pas de COM reel dans les tests) |
| `typer.testing.CliRunner` | Test des commandes CLI en isolation                  |

**Pattern de mock COM :**

```python
def test_workbook_open(mocker):
    """Exemple de test avec mock COM."""
    # Creer un mock de l'application Excel
    mock_app = mocker.Mock()
    mock_wb = mocker.Mock()
    mock_wb.Name = "test.xlsx"
    mock_wb.FullName = "C:\\test.xlsx"
    mock_wb.ReadOnly = False
    mock_wb.Saved = True
    mock_wb.Worksheets.Count = 3
    mock_app.Workbooks.Open.return_value = mock_wb
    mock_app.Workbooks.__iter__ = mocker.Mock(return_value=iter([]))

    # Creer les managers avec le mock
    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app
    wb_mgr = WorkbookManager(mgr)

    # Executer et verifier
    info = wb_mgr.open(Path("C:\\test.xlsx"))
    assert info.name == "test.xlsx"
    mock_app.Workbooks.Open.assert_called_once()
```

## 11. Constraints and Assumptions

### 11.1 Technical Constraints

- **Windows uniquement** : pywin32 et COM automation ne fonctionnent que sur Windows.
- **Excel installe** : Microsoft Excel doit etre installe avec une licence valide.
- **Python >= 3.14** : utilise les fonctionnalites recentes (union types `X | Y`, etc.).
- **Single thread STA** : Excel est un serveur STA. Chaque thread accedant au COM doit appeler `pythoncom.CoInitialize()`.
- **JAMAIS `app.Quit()`** : provoque des erreurs RPC car Python detient encore des references COM.
- **pytest-xdist incompatible** : le parallelisme de tests est incompatible avec le COM (partage d'etat global via ROT).

### 11.2 Business Constraints

- **Langue du code** : anglais (noms de variables, fonctions, classes, commentaires techniques).
- **Langue de la CLI** : francais (messages utilisateur, aide des commandes).
- **Licence** : GPL-3.0-or-later.
- **Pas d'emojis** dans le code ou les strings.
- **Chemins** : toujours `pathlib.Path`, jamais `os.path`.
- **Encodage** : UTF-8 partout sauf fichiers VBA (windows-1252).

### 11.3 Assumptions

- L'utilisateur a les droits d'administration locaux pour installer Excel et Python.
- Excel est la seule application utilisant le ProgID `"Excel.Application"` dans le ROT.
- Le Trust Center est configurable par l'utilisateur (pas bloque par une GPO).
- Les fichiers Excel ne sont pas proteges par Azure Information Protection ou SharePoint locks.

## 12. Risks and Mitigations

| Risk                                      | Likelihood | Impact | Mitigation                                                                                 |
| ----------------------------------------- | ---------- | ------ | ------------------------------------------------------------------------------------------ |
| Process Excel zombie apres crash Python   | High       | Med    | Liberation ordonnee (`del` + `gc.collect()`), `--force` avec `taskkill` en dernier recours |
| RPC error 0x800706BE pendant arret        | Med        | High   | Retry avec backoff, puis fallback `taskkill /f`                                            |
| Trust Center bloque par GPO               | Low        | High   | Documentation claire, message d'erreur explicite (`VBAProjectAccessError`)                 |
| Perte de donnees en mode `--force`        | Med        | High   | Avertissement Rich en rouge, lister les classeurs non sauvegardes avant arret              |
| Conflit ROT avec plusieurs instances      | Med        | Med    | Enumeration complete via ROT + fallback tasklist                                           |
| Encodage VBA incorrect (pas windows-1252) | Low        | Med    | Forcer l'encodage `windows-1252` a la lecture et a l'ecriture                              |

## 13. References

- [pywin32 documentation](https://github.com/mhammond/pywin32)
- [Excel COM Object Model Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel)
- [Typer documentation](https://typer.tiangolo.com/)
- [Rich documentation](https://rich.readthedocs.io/)

---

## Annexe A - Structure des fichiers (cible)

```
src/xlmanage/
    __init__.py                # Exports publics + __version__
    cli.py                     # Commandes Typer (couche mince)
    exceptions.py              # Toutes les exceptions du projet
    excel_manager.py           # [A CREER] Cycle de vie Excel (Epic 5, 11)
    workbook_manager.py        # [A CREER] CRUD classeurs (Epic 6)
    worksheet_manager.py       # [A CREER] CRUD feuilles (Epic 7)
    table_manager.py           # [A CREER] CRUD tables ListObject (Epic 8)
    vba_manager.py             # [A CREER] Import/export/delete VBA (Epic 9)
    macro_runner.py            # [A CREER] Execution de macros (Epic 12)
    excel_optimizer.py         # [EXISTANT] Optimisation complete (8 props)
    screen_optimizer.py        # [EXISTANT] Optimisation ecran (3 props)
    calculation_optimizer.py   # [EXISTANT] Optimisation calcul (4 props)
```

## Annexe B - Ordre d'implementation recommande

L'ordre suit les dependances entre les modules :

```
1. exceptions.py        (Epic 5 S1, 6 S1, 7 S1, 8 S1, 9 S1)  <- aucune dependance
2. excel_manager.py     (Epic 5 S2)                             <- exceptions
3. workbook_manager.py  (Epic 6 S2)                             <- excel_manager
4. worksheet_manager.py (Epic 7 S2)                             <- excel_manager, workbook_manager
5. table_manager.py     (Epic 8 S2)                             <- excel_manager, worksheet_manager
6. vba_manager.py       (Epic 9 S2)                             <- excel_manager, workbook_manager
7. macro_runner.py      (Epic 12 S1)                            <- excel_manager
8. optimizers refacto   (Epic 10 S1)                            <- excel_manager
9. cli.py integration   (Epics 5-12, Stories 3)                 <- tous les managers
10. stop complet        (Epic 11)                               <- excel_manager, ROT
```

## Annexe C - Regles COM critiques (a connaitre par coeur)

1. **JAMAIS `app.Quit()`** : provoque RPC error 0x800706BE.
2. **Liberation ordonnee** : `del ws` avant `del wb` avant `del app` puis `gc.collect()`.
3. **`Dispatch()` vs `DispatchEx()`** : `Dispatch` reutilise via ROT, `DispatchEx` cree un process isole.
4. **`DisplayAlerts = False`** : obligatoire avant toute operation pouvant declencher un dialogue (Close, Delete sheet, etc.).
5. **Chemins absolus** : toujours `str(path.resolve())` avant passage au COM.
6. **Thread STA** : chaque thread doit appeler `pythoncom.CoInitialize()`.
7. **Encodage VBA** : `windows-1252` avec fin de ligne CRLF (`\r\n`).

---

**Document History**

| Version | Date       | Author | Changes                                                    |
| ------- | ---------- | ------ | ---------------------------------------------------------- |
| 1.0.0   | 2026-02-01 | Claude | Initial version - architecture complete depuis PROGRESS.md |
