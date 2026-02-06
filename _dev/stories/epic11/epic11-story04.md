# Epic 11 - Story 4: Intégrer les commandes stop dans le CLI

**Statut** : ✅ Terminé

**En tant que** utilisateur
**Je veux** arrêter les instances Excel depuis le CLI
**Afin de** gérer facilement mes processus Excel en ligne de commande

## Critères d'acceptation

1. ✅ `xlmanage stop` arrête l'instance gérée ou active
2. ✅ `xlmanage stop <pid>` arrête une instance spécifique
3. ✅ `xlmanage stop --all` arrête toutes les instances
4. ✅ `xlmanage stop --force` utilise force_kill
5. ✅ `xlmanage stop --no-save` ne sauvegarde pas les classeurs
6. ✅ Un résumé Rich affiche les instances arrêtées
7. ✅ Les tests CLI passent pour toutes les variantes

## Tâches techniques

### Tâche 4.1 : Implémenter `xlmanage stop`

**Fichier** : `src/xlmanage/cli.py`

La commande existe en stub. Il faut l'implémenter :

```python
@app.command()
def stop(
    instance_id: Annotated[
        Optional[str],
        typer.Argument(help="PID de l'instance à arrêter (optionnel)")
    ] = None,
    all_instances: Annotated[
        bool,
        typer.Option("--all", help="Arrêter toutes les instances Excel")
    ] = False,
    force: Annotated[
        bool,
        typer.Option("--force", help="Forcer l'arrêt avec taskkill (sans sauvegarde)")
    ] = False,
    no_save: Annotated[
        bool,
        typer.Option("--no-save", help="Ne pas sauvegarder les classeurs")
    ] = False,
) -> None:
    """Arrête une ou plusieurs instances Excel.

    Sans argument : arrête l'instance active (ou celle gérée par xlManage).
    Avec PID : arrête l'instance spécifique.
    Avec --all : arrête toutes les instances Excel.
    Avec --force : utilise taskkill (perte de données !).

    Exemples:

        xlmanage stop

        xlmanage stop 12345

        xlmanage stop --all --no-save

        xlmanage stop 12345 --force
    """
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table

    from .excel_manager import ExcelManager
    from .exceptions import (
        ExcelInstanceNotFoundError,
        ExcelRPCError,
    )

    console = Console()

    # Validation : --all incompatible avec instance_id
    if all_instances and instance_id:
        console.print(
            "[red]Erreur :[/red] Impossible de spécifier --all ET un PID",
            style="bold"
        )
        raise typer.Exit(code=1)

    # Déterminer save (inverse de no_save)
    save = not no_save

    try:
        mgr = ExcelManager()

        # --force : utiliser force_kill
        if force:
            _force_kill_instances(mgr, instance_id, all_instances, console)
            return

        # --all : arrêter toutes les instances
        if all_instances:
            _stop_all_instances(mgr, save, console)
            return

        # PID spécifique
        if instance_id:
            pid = int(instance_id)
            _stop_single_instance(mgr, pid, save, console)
            return

        # Aucun argument : arrêter l'instance active
        _stop_active_instance(mgr, save, console)

    except ValueError:
        console.print(
            f"[red]Erreur :[/red] PID invalide '{instance_id}'. "
            "Le PID doit être un nombre entier.",
            style="bold"
        )
        raise typer.Exit(code=1)

    except ExcelInstanceNotFoundError as e:
        console.print(f"[red]Instance introuvable :[/red] {e}", style="bold")
        raise typer.Exit(code=1)

    except ExcelRPCError as e:
        console.print(
            f"[red]Erreur RPC :[/red] L'instance est déconnectée ou zombie\\n"
            f"[yellow]Utilisez --force pour terminer le processus[/yellow]",
            style="bold"
        )
        raise typer.Exit(code=1)

    except Exception as e:
        console.print(f"[red]Erreur :[/red] {e}", style="bold")
        raise typer.Exit(code=1)
```

### Tâche 4.2 : Implémenter _stop_active_instance()

Fonctions helper pour gérer les différents modes :

```python
def _stop_active_instance(mgr: ExcelManager, save: bool, console: Console) -> None:
    """Arrête l'instance Excel active."""
    # Chercher l'instance active
    info = mgr.get_running_instance()

    if info is None:
        console.print("[yellow]Aucune instance Excel active[/yellow]")
        return

    console.print(f"[dim]Arrêt de l'instance PID {info.pid}...[/dim]")

    mgr.stop_instance(info.pid, save=save)

    console.print(Panel(
        f"[green]Instance arrêtée avec succès[/green]\\n\\n"
        f"PID : {info.pid}\\n"
        f"Classeurs : {info.workbooks_count}\\n"
        f"Sauvegarde : {'Oui' if save else 'Non'}",
        title="Arrêt Excel",
        border_style="green"
    ))
```

### Tâche 4.3 : Implémenter _stop_single_instance()

```python
def _stop_single_instance(mgr: ExcelManager, pid: int, save: bool, console: Console) -> None:
    """Arrête une instance spécifique par PID."""
    console.print(f"[dim]Arrêt de l'instance PID {pid}...[/dim]")

    mgr.stop_instance(pid, save=save)

    console.print(Panel(
        f"[green]Instance arrêtée avec succès[/green]\\n\\n"
        f"PID : {pid}\\n"
        f"Sauvegarde : {'Oui' if save else 'Non'}",
        title="Arrêt Excel",
        border_style="green"
    ))
```

### Tâche 4.4 : Implémenter _stop_all_instances()

```python
def _stop_all_instances(mgr: ExcelManager, save: bool, console: Console) -> None:
    """Arrête toutes les instances Excel."""
    # Lister d'abord les instances
    instances = mgr.list_running_instances()

    if not instances:
        console.print("[yellow]Aucune instance Excel active[/yellow]")
        return

    console.print(f"[dim]Arrêt de {len(instances)} instance(s)...[/dim]")

    stopped_pids = mgr.stop_all(save=save)

    # Afficher un tableau des instances arrêtées
    table = Table(title="Instances arrêtées", show_header=True)
    table.add_column("PID", justify="right", style="cyan")
    table.add_column("Statut", style="green")

    for pid in stopped_pids:
        table.add_row(str(pid), "Arrêtée")

    # Instances qui ont échoué
    failed_pids = [info.pid for info in instances if info.pid not in stopped_pids]
    for pid in failed_pids:
        table.add_row(str(pid), "[red]Échec[/red]")

    console.print(table)

    console.print(
        f"\\n[green]{len(stopped_pids)} instance(s) arrêtée(s) avec succès[/green]"
    )

    if failed_pids:
        console.print(
            f"[yellow]{len(failed_pids)} instance(s) en échec - "
            f"utilisez --force si nécessaire[/yellow]"
        )
```

### Tâche 4.5 : Implémenter _force_kill_instances()

```python
def _force_kill_instances(
    mgr: ExcelManager,
    instance_id: str | None,
    all_instances: bool,
    console: Console
) -> None:
    """Force l'arrêt brutal avec taskkill."""
    # Avertissement
    console.print(
        "[red bold]ATTENTION : Force kill terminera brutalement Excel "
        "sans sauvegarder les classeurs ![/red bold]\\n"
    )

    if all_instances:
        # Lister toutes les instances
        instances = mgr.list_running_instances()

        if not instances:
            console.print("[yellow]Aucune instance Excel active[/yellow]")
            return

        console.print(f"[dim]Force kill de {len(instances)} instance(s)...[/dim]")

        for info in instances:
            try:
                mgr.force_kill(info.pid)
                console.print(f"[green]PID {info.pid} : terminé[/green]")
            except Exception as e:
                console.print(f"[red]PID {info.pid} : échec - {e}[/red]")

    elif instance_id:
        pid = int(instance_id)
        console.print(f"[dim]Force kill de PID {pid}...[/dim]")

        mgr.force_kill(pid)

        console.print(Panel(
            f"[green]Processus terminé avec force[/green]\\n\\n"
            f"PID : {pid}\\n"
            f"[red]Classeurs perdus (non sauvegardés)[/red]",
            title="Force Kill",
            border_style="red"
        ))

    else:
        # Force kill de l'instance active
        info = mgr.get_running_instance()
        if info is None:
            console.print("[yellow]Aucune instance Excel active[/yellow]")
            return

        mgr.force_kill(info.pid)

        console.print(Panel(
            f"[green]Processus terminé avec force[/green]\\n\\n"
            f"PID : {info.pid}\\n"
            f"[red]Classeurs perdus (non sauvegardés)[/red]",
            title="Force Kill",
            border_style="red"
        ))
```

**Points d'attention** :
- L'avertissement en rouge BOLD pour --force
- Le border_style est "red" pour force_kill (danger)
- On affiche explicitement "Classeurs perdus"

## Tests à implémenter

Créer `tests/test_cli_stop.py` :

```python
import pytest
from typer.testing import CliRunner
from unittest.mock import Mock, patch

from xlmanage.cli import app
from xlmanage.excel_manager import InstanceInfo

runner = CliRunner()


def test_stop_active_instance():
    """Test stop without arguments (active instance)."""
    mock_info = InstanceInfo(pid=12345, visible=True, workbooks_count=2, hwnd=9999)

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:
        mock_mgr = Mock()
        mock_mgr.get_running_instance.return_value = mock_info
        mock_mgr.stop_instance = Mock()
        mock_mgr_class.return_value = mock_mgr

        result = runner.invoke(app, ["stop"])

        assert result.exit_code == 0
        assert "arrêtée avec succès" in result.stdout
        mock_mgr.stop_instance.assert_called_once_with(12345, save=True)


def test_stop_specific_pid():
    """Test stop with specific PID."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:
        mock_mgr = Mock()
        mock_mgr.stop_instance = Mock()
        mock_mgr_class.return_value = mock_mgr

        result = runner.invoke(app, ["stop", "12345"])

        assert result.exit_code == 0
        mock_mgr.stop_instance.assert_called_once_with(12345, save=True)


def test_stop_no_save():
    """Test stop with --no-save."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:
        mock_mgr = Mock()
        mock_mgr.stop_instance = Mock()
        mock_mgr_class.return_value = mock_mgr

        result = runner.invoke(app, ["stop", "12345", "--no-save"])

        assert result.exit_code == 0
        mock_mgr.stop_instance.assert_called_once_with(12345, save=False)


def test_stop_all():
    """Test stop --all."""
    instances = [
        InstanceInfo(12345, True, 1, 1111),
        InstanceInfo(67890, False, 2, 2222),
    ]

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:
        mock_mgr = Mock()
        mock_mgr.list_running_instances.return_value = instances
        mock_mgr.stop_all.return_value = [12345, 67890]
        mock_mgr_class.return_value = mock_mgr

        result = runner.invoke(app, ["stop", "--all"])

        assert result.exit_code == 0
        assert "2 instance(s) arrêtée(s)" in result.stdout
        mock_mgr.stop_all.assert_called_once()


def test_stop_force():
    """Test stop --force."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:
        mock_mgr = Mock()
        mock_mgr.force_kill = Mock()
        mock_mgr_class.return_value = mock_mgr

        result = runner.invoke(app, ["stop", "12345", "--force"])

        assert result.exit_code == 0
        assert "ATTENTION" in result.stdout
        assert "Force kill" in result.stdout
        mock_mgr.force_kill.assert_called_once_with(12345)


def test_stop_all_and_pid_error():
    """Test error when --all and PID are both specified."""
    result = runner.invoke(app, ["stop", "12345", "--all"])

    assert result.exit_code == 1
    assert "Impossible de spécifier --all ET un PID" in result.stdout


def test_stop_invalid_pid():
    """Test error with invalid PID format."""
    result = runner.invoke(app, ["stop", "not-a-number"])

    assert result.exit_code == 1
    assert "PID invalide" in result.stdout


def test_stop_no_instances():
    """Test stop when no instances are running."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:
        mock_mgr = Mock()
        mock_mgr.get_running_instance.return_value = None
        mock_mgr_class.return_value = mock_mgr

        result = runner.invoke(app, ["stop"])

        assert result.exit_code == 0
        assert "Aucune instance Excel active" in result.stdout
```

## Dépendances

- Epic 11, Stories 1-3 (toutes les méthodes d'arrêt)

## Définition of Done

- [x] `xlmanage stop` arrête l'instance active
- [x] `xlmanage stop <pid>` arrête une instance spécifique
- [x] `xlmanage stop --all` arrête toutes les instances
- [x] `xlmanage stop --force` utilise force_kill avec avertissement
- [x] `--no-save` fonctionne correctement
- [x] Les erreurs sont gérées avec messages Rich appropriés
- [x] Tous les tests CLI passent (15 tests)
- [x] L'aide CLI est complète avec exemples
