# Epic 11 - Story 2: Implémenter stop(), stop_instance() et stop_all()

**Statut** : ⏳ À faire

**En tant que** utilisateur
**Je veux** arrêter proprement les instances Excel
**Afin de** libérer les ressources et éviter les processus zombie

## Contexte

L'arrêt propre d'Excel est **critique**. La règle absolue : **JAMAIS `app.Quit()`**.

Le protocole d'arrêt propre :
1. `app.DisplayAlerts = False`
2. Fermer tous les classeurs : `wb.Close(SaveChanges=save)`
3. Libérer les références : `del wb`, `del app`
4. Garbage collection : `gc.collect()`
5. `self._app = None`

Excel se ferme automatiquement quand toutes les références COM sont libérées.

## Critères d'acceptation

1. ⬜ La méthode `stop()` arrête l'instance gérée par ExcelManager
2. ⬜ La méthode `stop_instance(pid)` arrête une instance spécifique
3. ⬜ La méthode `stop_all()` arrête toutes les instances Excel
4. ⬜ Le paramètre `save` fonctionne correctement
5. ⬜ **AUCUN** appel à `app.Quit()` dans le code
6. ⬜ Les tests vérifient la libération ordonnée des références

## Tâches techniques

### Tâche 2.1 : Implémenter stop()

**Fichier** : `src/xlmanage/excel_manager.py` (dans la classe)

```python
import gc


def stop(self, save: bool = True) -> None:
    """Arrête proprement l'instance Excel gérée.

    Protocole d'arrêt :
    1. Désactiver les alertes
    2. Fermer tous les classeurs
    3. Libérer les références COM (del)
    4. Garbage collection
    5. Mettre _app à None

    IMPORTANT : JAMAIS d'appel à app.Quit() - provoque des erreurs RPC.

    Args:
        save: Si True, sauvegarde chaque classeur avant fermeture

    Example:
        >>> mgr = ExcelManager()
        >>> mgr.start()
        >>> # ... travail ...
        >>> mgr.stop(save=True)
    """
    if self._app is None:
        # Déjà arrêté, rien à faire
        return

    try:
        # 1. Désactiver les alertes (évite les dialogues de confirmation)
        self._app.DisplayAlerts = False

        # 2. Fermer tous les classeurs
        workbooks = []
        try:
            # Copier la liste pour éviter les problèmes d'itération
            for wb in self._app.Workbooks:
                workbooks.append(wb)
        except pywintypes.com_error:
            # Erreur lors de l'énumération, ignorer
            pass

        for wb in workbooks:
            try:
                wb.Close(SaveChanges=save)
                del wb
            except pywintypes.com_error:
                # Classeur déjà fermé ou inaccessible
                continue

        # 3. Libérer la référence principale
        del self._app

    except pywintypes.com_error as e:
        # Erreur RPC (serveur déconnecté), ignorer
        # L'instance est probablement déjà morte
        pass

    finally:
        # 4. Garbage collection pour libérer toutes les références COM
        gc.collect()

        # 5. Marquer comme arrêté
        self._app = None
```

**Points d'attention** :
- L'ordre de libération est critique : `wb` avant `app`
- `DisplayAlerts = False` évite les dialogues "Voulez-vous sauvegarder ?"
- `gc.collect()` force Python à libérer immédiatement les références COM
- Les erreurs RPC sont normales si l'instance est déconnectée

### Tâche 2.2 : Implémenter stop_instance()

```python
def stop_instance(self, pid: int, save: bool = True) -> None:
    """Arrête une instance Excel identifiée par son PID.

    Se connecte à l'instance via le ROT ou HWND, puis applique
    le protocole stop().

    Args:
        pid: Process ID de l'instance Excel cible
        save: Si True, sauvegarde avant fermeture

    Raises:
        ExcelInstanceNotFoundError: Si le PID n'existe pas ou n'est pas Excel
        ExcelRPCError: Si l'instance est déconnectée

    Example:
        >>> mgr = ExcelManager()
        >>> mgr.stop_instance(12345, save=False)
    """
    # Énumérer toutes les instances
    instances = enumerate_excel_instances()

    # Chercher l'instance avec le bon PID
    target_app = None
    for app, info in instances:
        if info.pid == pid:
            target_app = app
            break

    if target_app is None:
        # Fallback : vérifier via tasklist si le PID existe
        all_pids = enumerate_excel_pids()
        if pid not in all_pids:
            raise ExcelInstanceNotFoundError(
                str(pid),
                "Process ID not found or not an Excel instance"
            )

        # PID existe mais inaccessible via COM
        raise ExcelRPCError(
            0x800706BE,
            f"Excel instance PID {pid} is disconnected or inaccessible"
        )

    # Appliquer le protocole d'arrêt
    try:
        target_app.DisplayAlerts = False

        # Fermer tous les classeurs
        workbooks = []
        for wb in target_app.Workbooks:
            workbooks.append(wb)

        for wb in workbooks:
            try:
                wb.Close(SaveChanges=save)
                del wb
            except pywintypes.com_error:
                continue

        # Libérer la référence
        del target_app

    except pywintypes.com_error as e:
        # Erreur RPC
        raise ExcelRPCError(e.hresult, f"RPC error during shutdown: {e}") from e

    finally:
        gc.collect()
```

**Points d'attention** :
- On cherche d'abord dans le ROT pour avoir l'objet COM
- Si le PID existe (tasklist) mais pas dans le ROT, c'est une instance déconnectée
- `ExcelRPCError` est levée si l'instance est zombie

### Tâche 2.3 : Implémenter stop_all()

```python
def stop_all(self, save: bool = True) -> list[int]:
    """Arrête toutes les instances Excel actives.

    Énumère via le ROT et applique stop_instance() pour chacune.

    Args:
        save: Si True, sauvegarde avant fermeture

    Returns:
        list[int]: Liste des PIDs arrêtés avec succès

    Example:
        >>> mgr = ExcelManager()
        >>> stopped = mgr.stop_all(save=True)
        >>> print(f"{len(stopped)} instances arrêtées")
    """
    # Énumérer toutes les instances
    instances = enumerate_excel_instances()

    stopped_pids: list[int] = []

    for app, info in instances:
        try:
            # Appliquer le protocole d'arrêt
            app.DisplayAlerts = False

            workbooks = []
            for wb in app.Workbooks:
                workbooks.append(wb)

            for wb in workbooks:
                try:
                    wb.Close(SaveChanges=save)
                    del wb
                except pywintypes.com_error:
                    continue

            del app

            stopped_pids.append(info.pid)

        except pywintypes.com_error:
            # Instance déconnectée, ignorer
            continue

    # Garbage collection final
    gc.collect()

    return stopped_pids
```

**Points d'attention** :
- On continue même si une instance échoue (try/except par instance)
- Retourne seulement les PIDs arrêtés avec succès
- `gc.collect()` final pour tout nettoyer

## Tests à implémenter

Créer `tests/test_excel_stop.py` :

```python
import pytest
from unittest.mock import Mock, MagicMock, patch
import pywintypes
import gc

from xlmanage.excel_manager import ExcelManager
from xlmanage.exceptions import ExcelInstanceNotFoundError, ExcelRPCError


def test_stop_success():
    """Test successful stop of managed instance."""
    mock_wb = Mock()
    mock_app = Mock()
    mock_app.Workbooks = [mock_wb]
    mock_app.DisplayAlerts = True

    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app

    mgr.stop(save=True)

    # Vérifier le protocole
    assert mock_app.DisplayAlerts is False
    mock_wb.Close.assert_called_once_with(SaveChanges=True)
    assert mgr._app is None


def test_stop_no_save():
    """Test stop without saving."""
    mock_wb = Mock()
    mock_app = Mock()
    mock_app.Workbooks = [mock_wb]

    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app

    mgr.stop(save=False)

    mock_wb.Close.assert_called_once_with(SaveChanges=False)


def test_stop_already_stopped():
    """Test stop when already stopped."""
    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = None

    # Ne doit pas lever d'erreur
    mgr.stop()

    assert mgr._app is None


def test_stop_with_rpc_error():
    """Test stop handles RPC errors gracefully."""
    mock_app = Mock()
    mock_app.DisplayAlerts = True
    mock_app.Workbooks = Mock(side_effect=pywintypes.com_error(-2147352567, "RPC", None, None))

    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app

    # Ne doit pas lever d'erreur
    mgr.stop()

    assert mgr._app is None


def test_stop_instance_success():
    """Test stopping a specific instance by PID."""
    mock_wb = Mock()
    mock_app = Mock()
    mock_app.Workbooks = [mock_wb]

    mock_info = Mock()
    mock_info.pid = 12345

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(mock_app, mock_info)]

        mgr = ExcelManager()
        mgr.stop_instance(12345, save=True)

        mock_wb.Close.assert_called_once()


def test_stop_instance_not_found():
    """Test error when PID doesn't exist."""
    with patch("xlmanage.excel_manager.enumerate_excel_instances", return_value=[]), \\
         patch("xlmanage.excel_manager.enumerate_excel_pids", return_value=[]):

        mgr = ExcelManager()

        with pytest.raises(ExcelInstanceNotFoundError):
            mgr.stop_instance(99999)


def test_stop_all_success():
    """Test stopping all Excel instances."""
    mock_wb1 = Mock()
    mock_app1 = Mock()
    mock_app1.Workbooks = [mock_wb1]

    mock_wb2 = Mock()
    mock_app2 = Mock()
    mock_app2.Workbooks = [mock_wb2]

    mock_info1 = Mock()
    mock_info1.pid = 12345

    mock_info2 = Mock()
    mock_info2.pid = 67890

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(mock_app1, mock_info1), (mock_app2, mock_info2)]

        mgr = ExcelManager()
        stopped = mgr.stop_all(save=False)

        assert len(stopped) == 2
        assert 12345 in stopped
        assert 67890 in stopped


def test_stop_all_with_errors():
    """Test stop_all continues when one instance fails."""
    mock_app1 = Mock()
    mock_app1.DisplayAlerts = Mock(side_effect=pywintypes.com_error())

    mock_wb2 = Mock()
    mock_app2 = Mock()
    mock_app2.Workbooks = [mock_wb2]

    mock_info1 = Mock()
    mock_info1.pid = 12345

    mock_info2 = Mock()
    mock_info2.pid = 67890

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(mock_app1, mock_info1), (mock_app2, mock_info2)]

        mgr = ExcelManager()
        stopped = mgr.stop_all()

        # Seule la 2ème instance a été arrêtée
        assert stopped == [67890]
```

## Dépendances

- Epic 11, Story 1 (fonctions d'énumération)
- Epic 5, Story 1 (exceptions ExcelRPCError, ExcelInstanceNotFoundError)

## Définition of Done

- [ ] `stop()` est implémentée avec le protocole complet
- [ ] `stop_instance()` fonctionne avec énumération ROT
- [ ] `stop_all()` arrête toutes les instances
- [ ] **Aucun** `app.Quit()` dans le code
- [ ] Les références COM sont libérées dans le bon ordre
- [ ] Tous les tests passent (8+ tests)
- [ ] Couverture > 95%
