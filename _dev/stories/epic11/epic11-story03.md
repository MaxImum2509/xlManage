# Epic 11 - Story 3: Implémenter force_kill()

**Statut** : ⏳ À faire

**En tant que** utilisateur
**Je veux** pouvoir forcer l'arrêt d'une instance Excel zombie
**Afin de** terminer un processus qui ne répond plus au protocole COM

## Contexte

Parfois, une instance Excel devient "zombie" : le processus existe encore mais ne répond plus aux commandes COM. Dans ce cas, le protocole d'arrêt propre (`stop()`) échoue avec une erreur RPC.

La méthode `force_kill()` est le **dernier recours** : elle utilise `taskkill /f /pid` pour terminer brutalement le processus Windows.

**Important** : Cette méthode doit être utilisée **uniquement** quand l'arrêt propre a échoué.

## Critères d'acceptation

1. ⬜ La méthode `force_kill()` est implémentée
2. ⬜ Elle utilise `taskkill /f /pid` (Windows)
3. ⬜ Une exception est levée si le PID n'existe pas
4. ⬜ Un warning est logué car c'est une opération dangereuse
5. ⬜ Les tests vérifient l'appel à subprocess

## Tâches techniques

### Tâche 3.1 : Implémenter force_kill()

**Fichier** : `src/xlmanage/excel_manager.py` (dans la classe)

```python
import subprocess
import logging

logger = logging.getLogger(__name__)


def force_kill(self, pid: int) -> None:
    """Arrêt forcé d'une instance Excel via taskkill.

    **Attention** : Cette méthode termine brutalement le processus sans
    sauvegarder les classeurs. À utiliser UNIQUEMENT quand l'arrêt propre
    a échoué et que l'instance est zombie.

    Utilise : taskkill /f /pid <pid>

    Args:
        pid: Process ID de l'instance à terminer

    Raises:
        ExcelInstanceNotFoundError: Si le PID n'existe pas
        RuntimeError: Si la commande taskkill échoue

    Example:
        >>> mgr = ExcelManager()
        >>> try:
        ...     mgr.stop_instance(12345)
        ... except ExcelRPCError:
        ...     # Arrêt propre échoué, forcer
        ...     mgr.force_kill(12345)
    """
    # Logger un warning (opération dangereuse)
    logger.warning(
        f"Force killing Excel instance PID {pid}. "
        "This will terminate the process without saving workbooks."
    )

    try:
        # Exécuter taskkill /f /pid
        result = subprocess.run(
            ["taskkill", "/f", "/pid", str(pid)],
            capture_output=True,
            text=True,
            check=True,
            timeout=10
        )

        # Vérifier le succès dans stdout
        if "SUCCESS" not in result.stdout:
            raise RuntimeError(f"taskkill failed: {result.stdout}")

        logger.info(f"Successfully force-killed Excel instance PID {pid}")

    except subprocess.CalledProcessError as e:
        # taskkill a échoué (PID inexistant, permissions, etc.)
        if "not found" in e.stdout or "not found" in e.stderr:
            raise ExcelInstanceNotFoundError(
                str(pid),
                "Process not found or not running"
            ) from e
        else:
            raise RuntimeError(
                f"Failed to kill process {pid}: {e.stderr or e.stdout}"
            ) from e

    except subprocess.TimeoutExpired:
        raise RuntimeError(f"Timeout while trying to kill process {pid}")

    except FileNotFoundError:
        raise RuntimeError(
            "taskkill command not found. This feature requires Windows."
        )
```

**Points d'attention** :
- `/f` = force (terminaison immédiate)
- `/pid` = spécifie le process ID
- `check=True` lève une exception si taskkill échoue
- Le message "SUCCESS" dans stdout confirme le succès
- Un timeout de 10 secondes pour éviter les blocages

### Tâche 3.2 : Configurer le logging pour ExcelManager

Ajouter en haut du fichier `excel_manager.py` :

```python
import logging

# Configurer le logger du module
logger = logging.getLogger(__name__)
```

Et optionnellement, dans `__init__.py` du package :

```python
import logging

# Configurer le logging de base pour xlmanage
logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
```

**Points d'attention** :
- Le logger est au niveau module, pas instance
- Niveau par défaut : WARNING (pour ne pas polluer la console)
- Les warnings de `force_kill()` seront visibles par défaut

## Tests à implémenter

Créer `tests/test_excel_force_kill.py` :

```python
import pytest
from unittest.mock import Mock, patch
import subprocess

from xlmanage.excel_manager import ExcelManager
from xlmanage.exceptions import ExcelInstanceNotFoundError


def test_force_kill_success():
    """Test successful force kill."""
    mock_result = Mock()
    mock_result.stdout = "SUCCESS: The process with PID 12345 has been terminated."
    mock_result.stderr = ""

    with patch("subprocess.run", return_value=mock_result) as mock_run:
        mgr = ExcelManager()
        mgr.force_kill(12345)

        # Vérifier l'appel à taskkill
        mock_run.assert_called_once_with(
            ["taskkill", "/f", "/pid", "12345"],
            capture_output=True,
            text=True,
            check=True,
            timeout=10
        )


def test_force_kill_process_not_found():
    """Test error when process doesn't exist."""
    error = subprocess.CalledProcessError(
        returncode=128,
        cmd=["taskkill"],
        output="",
        stderr="ERROR: The process '99999' not found."
    )
    error.stdout = ""
    error.stderr = "ERROR: The process '99999' not found."

    with patch("subprocess.run", side_effect=error):
        mgr = ExcelManager()

        with pytest.raises(ExcelInstanceNotFoundError) as exc_info:
            mgr.force_kill(99999)

        assert "99999" in str(exc_info.value)


def test_force_kill_access_denied():
    """Test error when access is denied."""
    error = subprocess.CalledProcessError(
        returncode=1,
        cmd=["taskkill"],
        output="ERROR: Access denied",
        stderr="ERROR: Access denied"
    )
    error.stdout = "ERROR: Access denied"
    error.stderr = ""

    with patch("subprocess.run", side_effect=error):
        mgr = ExcelManager()

        with pytest.raises(RuntimeError, match="Access denied"):
            mgr.force_kill(12345)


def test_force_kill_timeout():
    """Test timeout handling."""
    with patch("subprocess.run", side_effect=subprocess.TimeoutExpired("taskkill", 10)):
        mgr = ExcelManager()

        with pytest.raises(RuntimeError, match="Timeout"):
            mgr.force_kill(12345)


def test_force_kill_logs_warning(caplog):
    """Test that force_kill logs a warning."""
    import logging

    mock_result = Mock()
    mock_result.stdout = "SUCCESS"

    with patch("subprocess.run", return_value=mock_result):
        with caplog.at_level(logging.WARNING):
            mgr = ExcelManager()
            mgr.force_kill(12345)

            # Vérifier qu'un warning a été logué
            assert "Force killing" in caplog.text
            assert "12345" in caplog.text
```

## Considérations de sécurité

`force_kill()` est une opération **dangereuse** :

1. **Perte de données** : Les classeurs non sauvegardés sont perdus
2. **Corruption** : Si un classeur est en cours d'écriture, il peut être corrompu
3. **Pas de cleanup** : Les fichiers temporaires Excel ne sont pas nettoyés

**Recommandations** :
- Toujours essayer `stop()` ou `stop_instance()` AVANT `force_kill()`
- Logger un warning clair pour tracer l'utilisation
- Documenter les risques dans la docstring

## Dépendances

- Epic 11, Story 1 (fonctions d'énumération)
- Epic 5, Story 1 (exceptions)

## Définition of Done

- [ ] `force_kill()` est implémentée avec taskkill
- [ ] Un warning est logué à chaque utilisation
- [ ] Les erreurs taskkill sont correctement gérées
- [ ] Tous les tests passent (5+ tests)
- [ ] La docstring avertit clairement des risques
- [ ] Le logging est configuré pour le module
