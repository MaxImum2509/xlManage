# Epic 11 - Story 1: Implémenter les fonctions d'énumération Excel (ROT + tasklist)

**Statut** : ⏳ À faire

**En tant que** développeur
**Je veux** des fonctions pour énumérer les instances Excel actives
**Afin de** pouvoir les arrêter sélectivement ou toutes à la fois

## Contexte

Pour arrêter des instances Excel, il faut d'abord pouvoir les lister. Excel peut avoir plusieurs instances (processus) actives en même temps. On doit :

1. Lister via le **Running Object Table (ROT)** Windows pour récupérer les objets COM
2. Fallback via **tasklist** si le ROT n'est pas accessible
3. Extraire le PID de chaque instance pour identification

## Critères d'acceptation

1. ⬜ La fonction `enumerate_excel_instances()` énumère via le ROT
2. ⬜ La fonction `enumerate_excel_pids()` énumère via tasklist (fallback)
3. ⬜ La fonction `connect_by_hwnd()` se connecte à une instance par handle
4. ⬜ La méthode `list_running_instances()` dans ExcelManager utilise ces fonctions
5. ⬜ Les tests couvrent l'énumération ROT et le fallback tasklist

## Tâches techniques

### Tâche 1.1 : Implémenter enumerate_excel_instances()

**Fichier** : `src/xlmanage/excel_manager.py`

Ajouter cette fonction au niveau module (pas dans la classe) :

```python
import pythoncom
import pywintypes
from typing import Iterator


def enumerate_excel_instances() -> list[tuple[CDispatch, InstanceInfo]]:
    """Énumère les instances Excel via le Running Object Table (ROT).

    Le ROT Windows contient tous les objets COM actifs. On filtre ceux qui
    correspondent à Excel.Application.

    Returns:
        list[tuple[CDispatch, InstanceInfo]]: Liste de (app, info) pour chaque instance

    Note:
        Cette fonction peut échouer si l'accès au ROT est bloqué. Utiliser
        enumerate_excel_pids() comme fallback.
    """
    instances: list[tuple[CDispatch, InstanceInfo]] = []

    try:
        # Obtenir le ROT
        rot = pythoncom.GetRunningObjectTable()
        # Énumérer les monikers (identificateurs d'objets COM)
        monikers = rot.EnumRunning()

        for moniker in monikers:
            try:
                # Obtenir le nom du moniker
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)

                # Filtrer pour Excel.Application
                # Le nom contient "!Microsoft_Excel_Application"
                if "Excel.Application" not in name:
                    continue

                # Obtenir l'objet COM depuis le ROT
                obj = rot.GetObject(moniker)

                # Cast vers CDispatch
                app = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))

                # Extraire les infos de l'instance
                info = _get_instance_info_from_app(app)

                instances.append((app, info))

            except pywintypes.com_error:
                # Instance inaccessible ou déconnectée, ignorer
                continue

    except pywintypes.com_error:
        # ROT inaccessible, retourner liste vide (fallback nécessaire)
        return []

    return instances


def _get_instance_info_from_app(app: CDispatch) -> InstanceInfo:
    """Extrait InstanceInfo depuis un objet Application.

    Args:
        app: Objet COM Excel.Application

    Returns:
        InstanceInfo: Informations de l'instance
    """
    import ctypes

    # Récupérer le HWND (handle de fenêtre)
    hwnd = app.Hwnd

    # Extraire le PID depuis le HWND via l'API Windows
    process_id = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
    pid = process_id.value

    # Récupérer les autres infos
    visible = app.Visible
    workbooks_count = app.Workbooks.Count

    return InstanceInfo(
        pid=pid,
        visible=visible,
        workbooks_count=workbooks_count,
        hwnd=hwnd
    )
```

**Points d'attention** :
- Le ROT peut être inaccessible (permissions, environnement sandbox)
- Les monikers contiennent "Excel.Application" dans leur DisplayName
- `GetWindowThreadProcessId` est une API Windows (user32.dll)
- Certaines instances peuvent être déconnectées (COM error), on les ignore

### Tâche 1.2 : Implémenter enumerate_excel_pids()

Fallback via tasklist quand le ROT n'est pas accessible :

```python
import subprocess
import re


def enumerate_excel_pids() -> list[int]:
    """Énumère les PIDs Excel via tasklist (fallback).

    Utilisé quand le ROT n'est pas accessible. Retourne seulement les PIDs,
    pas les objets COM.

    Returns:
        list[int]: Liste des PIDs EXCEL.EXE

    Raises:
        RuntimeError: Si tasklist échoue (commande introuvable)
    """
    try:
        # Appeler tasklist avec filtre sur EXCEL.EXE
        result = subprocess.run(
            ["tasklist", "/fi", "imagename eq EXCEL.EXE", "/fo", "csv", "/nh"],
            capture_output=True,
            text=True,
            check=True,
            timeout=10
        )

        pids: list[int] = []

        # Parser le CSV de sortie
        # Format: "EXCEL.EXE","12345","Console","1","123,456 K"
        for line in result.stdout.strip().split("\\n"):
            if not line or "INFO:" in line:
                continue

            # Extraire le PID (2ème colonne)
            match = re.search(r'"EXCEL\\.EXE","(\\d+)"', line)
            if match:
                pid = int(match.group(1))
                pids.append(pid)

        return pids

    except subprocess.TimeoutExpired:
        raise RuntimeError("Timeout lors de l'énumération des processus Excel")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Échec de tasklist: {e}")
    except FileNotFoundError:
        raise RuntimeError("Commande tasklist introuvable (Windows requis)")
```

**Points d'attention** :
- `tasklist` est une commande Windows native
- Le format CSV (`/fo csv`) facilite le parsing
- `/nh` supprime l'en-tête
- Timeout de 10 secondes pour éviter les blocages

### Tâche 1.3 : Implémenter connect_by_hwnd()

Connexion à une instance via son handle de fenêtre (fallback avancé) :

```python
def connect_by_hwnd(hwnd: int) -> CDispatch | None:
    """Se connecte à une instance Excel par son handle de fenêtre.

    Utilisé quand l'instance n'est pas dans le ROT mais est encore active.

    Args:
        hwnd: Handle de fenêtre Windows (HWND)

    Returns:
        CDispatch | None: Objet Application Excel, ou None si échec
    """
    import ctypes
    from ctypes import POINTER, c_void_p, c_long
    from ctypes.wintypes import DWORD

    try:
        # Charger oleacc.dll (Accessibility API)
        oleacc = ctypes.windll.oleacc

        # Constantes
        OBJID_NATIVEOM = -16  # ID pour l'objet natif Office

        # Obtenir IDispatch depuis le HWND
        ptr = c_void_p()
        result = oleacc.AccessibleObjectFromWindow(
            hwnd,
            DWORD(OBJID_NATIVEOM),
            ctypes.byref(pythoncom.IID_IDispatch),
            ctypes.byref(ptr)
        )

        if result != 0 or not ptr:
            return None

        # Convertir IDispatch en CDispatch
        dispatch = pythoncom.ObjectFromLresult(
            ptr.value,
            pythoncom.IID_IDispatch,
            0
        )
        app = win32com.client.Dispatch(dispatch)

        return app

    except Exception:
        # Échec de connexion, retourner None
        return None
```

**Points d'attention** :
- `AccessibleObjectFromWindow` est une API Windows avancée
- Cette méthode fonctionne même si l'instance n'est pas dans le ROT
- Utile pour les instances "zombie" encore actives mais déconnectées

### Tâche 1.4 : Implémenter list_running_instances() dans ExcelManager

**Fichier** : `src/xlmanage/excel_manager.py` (dans la classe)

```python
def list_running_instances(self) -> list[InstanceInfo]:
    """Énumère toutes les instances Excel actives.

    Utilise le ROT en priorité, puis fallback sur tasklist si le ROT échoue.

    Returns:
        list[InstanceInfo]: Liste des instances avec leurs informations

    Example:
        >>> mgr = ExcelManager()
        >>> instances = mgr.list_running_instances()
        >>> for inst in instances:
        ...     print(f"PID {inst.pid}: {inst.workbooks_count} classeurs")
    """
    # Essayer via le ROT
    rot_instances = enumerate_excel_instances()

    if rot_instances:
        # Extraire juste les InstanceInfo
        return [info for app, info in rot_instances]

    # Fallback : tasklist pour obtenir les PIDs
    try:
        pids = enumerate_excel_pids()

        # Convertir les PIDs en InstanceInfo (infos limitées)
        instances = []
        for pid in pids:
            # On ne peut pas obtenir visible/workbooks_count sans COM
            info = InstanceInfo(
                pid=pid,
                visible=False,  # Inconnu
                workbooks_count=0,  # Inconnu
                hwnd=0  # Inconnu
            )
            instances.append(info)

        return instances

    except RuntimeError:
        # Fallback échoué aussi, retourner liste vide
        return []
```

**Points d'attention** :
- Le ROT donne des infos complètes (visible, classeurs, hwnd)
- tasklist donne seulement les PIDs (infos limitées)
- Si les deux échouent, on retourne une liste vide

## Tests à implémenter

Créer `tests/test_excel_enumeration.py` :

```python
import pytest
from unittest.mock import Mock, patch, MagicMock
import pythoncom
import pywintypes

from xlmanage.excel_manager import (
    enumerate_excel_instances,
    enumerate_excel_pids,
    connect_by_hwnd,
    ExcelManager,
    InstanceInfo,
)


def test_enumerate_excel_instances_success():
    """Test successful enumeration via ROT."""
    # Mock du ROT
    mock_moniker = Mock()
    mock_moniker.GetDisplayName.return_value = "!Microsoft_Excel_Application"

    mock_rot = Mock()
    mock_rot.EnumRunning.return_value = [mock_moniker]

    mock_obj = Mock()
    mock_rot.GetObject.return_value = mock_obj

    # Mock de l'app Excel
    mock_app = Mock()
    mock_app.Hwnd = 12345
    mock_app.Visible = True
    mock_app.Workbooks.Count = 2

    with patch("pythoncom.GetRunningObjectTable", return_value=mock_rot), \\
         patch("win32com.client.Dispatch", return_value=mock_app), \\
         patch("ctypes.windll.user32.GetWindowThreadProcessId") as mock_get_pid:

        # Simuler GetWindowThreadProcessId
        def set_pid(hwnd, pid_ref):
            pid_ref.value = 9876

        mock_get_pid.side_effect = set_pid

        instances = enumerate_excel_instances()

        assert len(instances) == 1
        app, info = instances[0]
        assert info.pid == 9876
        assert info.visible is True
        assert info.workbooks_count == 2


def test_enumerate_excel_instances_rot_error():
    """Test fallback when ROT is inaccessible."""
    with patch("pythoncom.GetRunningObjectTable", side_effect=pywintypes.com_error()):
        instances = enumerate_excel_instances()
        assert instances == []


def test_enumerate_excel_pids_success():
    """Test enumeration via tasklist."""
    tasklist_output = '''
"EXCEL.EXE","12345","Console","1","100,000 K"
"EXCEL.EXE","67890","Console","1","120,000 K"
'''

    mock_result = Mock()
    mock_result.stdout = tasklist_output

    with patch("subprocess.run", return_value=mock_result):
        pids = enumerate_excel_pids()

        assert len(pids) == 2
        assert 12345 in pids
        assert 67890 in pids


def test_enumerate_excel_pids_no_instances():
    """Test tasklist with no Excel instances."""
    mock_result = Mock()
    mock_result.stdout = "INFO: No tasks found.\\n"

    with patch("subprocess.run", return_value=mock_result):
        pids = enumerate_excel_pids()
        assert pids == []


def test_enumerate_excel_pids_timeout():
    """Test timeout handling."""
    with patch("subprocess.run", side_effect=subprocess.TimeoutExpired("tasklist", 10)):
        with pytest.raises(RuntimeError, match="Timeout"):
            enumerate_excel_pids()


def test_list_running_instances_via_rot():
    """Test list_running_instances using ROT."""
    mock_info = InstanceInfo(pid=12345, visible=True, workbooks_count=2, hwnd=9999)

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(Mock(), mock_info)]

        mgr = ExcelManager()
        instances = mgr.list_running_instances()

        assert len(instances) == 1
        assert instances[0].pid == 12345


def test_list_running_instances_fallback_tasklist():
    """Test list_running_instances fallback to tasklist."""
    with patch("xlmanage.excel_manager.enumerate_excel_instances", return_value=[]), \\
         patch("xlmanage.excel_manager.enumerate_excel_pids", return_value=[12345, 67890]):

        mgr = ExcelManager()
        instances = mgr.list_running_instances()

        assert len(instances) == 2
        assert instances[0].pid == 12345
        assert instances[1].pid == 67890
```

## Dépendances

- Epic 5, Story 2 (ExcelManager.start et dataclass InstanceInfo)

## Définition of Done

- [ ] `enumerate_excel_instances()` fonctionne avec le ROT
- [ ] `enumerate_excel_pids()` fonctionne avec tasklist
- [ ] `connect_by_hwnd()` est implémentée
- [ ] `list_running_instances()` utilise ROT puis fallback tasklist
- [ ] Tous les tests passent (8+ tests)
- [ ] Couverture > 90% pour les fonctions d'énumération
- [ ] Les docstrings sont complètes
