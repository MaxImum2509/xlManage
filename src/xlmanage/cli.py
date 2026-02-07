"""
CLI module for xlmanage.

This file is part of xlManage.

xlManage is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

xlManage is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with xlManage.  If not, see <https://www.gnu.org/licenses/>.
"""

from pathlib import Path
from typing import cast

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

try:
    from .excel_manager import ExcelManager, InstanceInfo
    from .exceptions import (
        ExcelConnectionError,
        ExcelInstanceNotFoundError,
        ExcelManageError,
        ExcelRPCError,
        TableAlreadyExistsError,
        TableNameError,
        TableNotFoundError,
        TableRangeError,
        VBAExportError,
        VBAImportError,
        VBAMacroError,
        VBAModuleAlreadyExistsError,
        VBAModuleNotFoundError,
        VBAProjectAccessError,
        VBAWorkbookFormatError,
        WorkbookAlreadyOpenError,
        WorkbookNotFoundError,
        WorkbookSaveError,
        WorksheetAlreadyExistsError,
        WorksheetDeleteError,
        WorksheetNameError,
        WorksheetNotFoundError,
    )
    from .macro_runner import MacroResult, MacroRunner, _format_return_value
    from .table_manager import TableManager
    from .vba_manager import VBAManager
    from .workbook_manager import WorkbookManager
    from .worksheet_manager import WorksheetManager
except ImportError:
    from xlmanage.excel_manager import ExcelManager
    from xlmanage.exceptions import (
        ExcelConnectionError,
        ExcelInstanceNotFoundError,
        ExcelManageError,
        ExcelRPCError,
        TableAlreadyExistsError,
        TableNameError,
        TableNotFoundError,
        TableRangeError,
        VBAExportError,
        VBAImportError,
        VBAMacroError,
        VBAModuleAlreadyExistsError,
        VBAModuleNotFoundError,
        VBAProjectAccessError,
        VBAWorkbookFormatError,
        WorkbookAlreadyOpenError,
        WorkbookNotFoundError,
        WorkbookSaveError,
        WorksheetAlreadyExistsError,
        WorksheetDeleteError,
        WorksheetNameError,
        WorksheetNotFoundError,
    )
    from xlmanage.macro_runner import MacroResult, MacroRunner, _format_return_value
    from xlmanage.table_manager import TableManager
    from xlmanage.vba_manager import VBAManager
    from xlmanage.workbook_manager import WorkbookManager
    from xlmanage.worksheet_manager import WorksheetManager

app = typer.Typer(
    name="xlmanage",
    help="Excel automation CLI tool",
    no_args_is_help=True,
)
console = Console()


@app.command()
def version():
    """Show version information."""
    console.print("[bold green]xlmanage[/bold green] version [cyan]0.1.0[/cyan]")


@app.command()
def start(
    visible: bool = typer.Option(
        False,
        "--visible",
        "-v",
        help="Make the Excel instance visible on screen",
    ),
    new: bool = typer.Option(
        False,
        "--new",
        "-n",
        help="Force creation of a new Excel instance",
    ),
):
    """Start a new Excel instance or connect to an existing one.

    By default, connects to an existing instance if available (via ROT).
    Use --new to force creation of a new isolated instance.
    Use --visible to make the Excel window visible on screen.
    """
    try:
        manager = ExcelManager(visible=visible)
        info = manager.start(new=new)

        # Create an empty workbook to keep Excel alive
        # Excel will close automatically if no workbooks are open
        manager.app.Workbooks.Add()

        # Display success message
        mode = "new" if new else "existing"
        visibility = "visible" if visible else "hidden"

        console.print(
            Panel.fit(
                f"[green]OK[/green] Excel instance started successfully\n\n"
                f"[bold]Mode:[/bold] {mode}\n"
                f"[bold]Visibility:[/bold] {visibility}\n"
                f"[bold]Process ID:[/bold] {info.pid}\n"
                f"[bold]Window Handle:[/bold] {info.hwnd}\n"
                f"[bold]Workbooks:[/bold] {info.workbooks_count + 1}",
                title="Excel Instance Started",
                border_style="green",
            )
        )

    except ExcelConnectionError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Failed to start Excel instance\n\n"
                f"[bold]Error:[/bold] {e}",
                title="Connection Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Excel management error\n\n[bold]Error:[/bold] {e}",
                title="Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except Exception as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Unexpected error\n\n[bold]Error:[/bold] {e}",
                title="Unexpected Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


# Helper functions for stop command


def _stop_active_instance(mgr: ExcelManager, save: bool, console: Console) -> None:
    """Stop the active Excel instance."""
    # Find active instance
    info = mgr.get_running_instance()

    if info is None:
        console.print("[yellow]Aucune instance Excel active[/yellow]")
        return

    console.print(f"[dim]Arrêt de l'instance PID {info.pid}...[/dim]")

    mgr.stop_instance(info.pid, save=save)

    console.print(
        Panel(
            f"[green]Instance arrêtée avec succès[/green]\n\n"
            f"PID : {info.pid}\n"
            f"Classeurs : {info.workbooks_count}\n"
            f"Sauvegarde : {'Oui' if save else 'Non'}",
            title="Arrêt Excel",
            border_style="green",
        )
    )


def _stop_single_instance(
    mgr: ExcelManager, pid: int, save: bool, console: Console
) -> None:
    """Stop a specific instance by PID."""
    console.print(f"[dim]Arrêt de l'instance PID {pid}...[/dim]")

    mgr.stop_instance(pid, save=save)

    console.print(
        Panel(
            f"[green]Instance arrêtée avec succès[/green]\n\n"
            f"PID : {pid}\n"
            f"Sauvegarde : {'Oui' if save else 'Non'}",
            title="Arrêt Excel",
            border_style="green",
        )
    )


def _stop_all_instances(mgr: ExcelManager, save: bool, console: Console) -> None:
    """Stop all Excel instances."""
    # List instances first
    instances = mgr.list_running_instances()

    if not instances:
        console.print("[yellow]Aucune instance Excel active[/yellow]")
        return

    console.print(f"[dim]Arrêt de {len(instances)} instance(s)...[/dim]")

    stopped_pids = mgr.stop_all(save=save)

    # Display table of stopped instances
    table = Table(title="Instances arrêtées", show_header=True)
    table.add_column("PID", justify="right", style="cyan")
    table.add_column("Statut", style="green")

    for pid in stopped_pids:
        table.add_row(str(pid), "Arrêtée")

    # Failed instances
    failed_pids = [info.pid for info in instances if info.pid not in stopped_pids]
    for pid in failed_pids:
        table.add_row(str(pid), "[red]Échec[/red]")

    console.print(table)

    console.print(
        f"\n[green]{len(stopped_pids)} instance(s) arrêtée(s) avec succès[/green]"
    )

    if failed_pids:
        console.print(
            f"[yellow]{len(failed_pids)} instance(s) en échec - "
            f"utilisez --force si nécessaire[/yellow]"
        )


def _force_kill_instances(
    mgr: ExcelManager, instance_id: str | None, all_instances: bool, console: Console
) -> None:
    """Force kill with taskkill."""
    # Warning
    console.print(
        "[red bold]ATTENTION : Force kill terminera brutalement Excel "
        "sans sauvegarder les classeurs ![/red bold]\n"
    )

    if all_instances:
        # List all instances
        instances = mgr.list_running_instances()

        if not instances:
            console.print("[yellow]Aucune instance Excel active[/yellow]")
            return

        console.print(f"[dim]Force kill de {len(instances)} instance(s)...[/dim]")

        for instance in instances:
            try:
                mgr.force_kill(instance.pid)
                console.print(f"[green]PID {instance.pid} : terminé[/green]")
            except Exception as e:
                console.print(f"[red]PID {instance.pid} : échec - {e}[/red]")

    elif instance_id:
        pid = int(instance_id)
        console.print(f"[dim]Force kill de PID {pid}...[/dim]")

        mgr.force_kill(pid)

        console.print(
            Panel(
                f"[green]Processus terminé avec force[/green]\n\n"
                f"PID : {pid}\n"
                f"[red]Classeurs perdus (non sauvegardés)[/red]",
                title="Force Kill",
                border_style="red",
            )
        )

    else:
        # Force kill active instance
        info: InstanceInfo | None = mgr.get_running_instance()
        if info is None:
            console.print("[yellow]Aucune instance Excel active[/yellow]")
            return

        mgr.force_kill(info.pid)

        console.print(
            Panel(
                f"[green]Processus terminé avec force[/green]\n\n"
                f"PID : {info.pid}\n"
                f"[red]Classeurs perdus (non sauvegardés)[/red]",
                title="Force Kill",
                border_style="red",
            )
        )


@app.command()
def stop(
    instance_id: str | None = typer.Argument(
        None, help="PID de l'instance à arrêter (optionnel)"
    ),
    all_instances: bool = typer.Option(
        False, "--all", help="Arrêter toutes les instances Excel"
    ),
    force: bool = typer.Option(
        False, "--force", help="Forcer l'arrêt avec taskkill (sans sauvegarde)"
    ),
    no_save: bool = typer.Option(
        False, "--no-save", help="Ne pas sauvegarder les classeurs"
    ),
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
    # Validation: --all incompatible with instance_id
    if all_instances and instance_id:
        console.print(
            "[red]Erreur :[/red] Impossible de spécifier --all ET un PID", style="bold"
        )
        raise typer.Exit(code=1)

    # Determine save (inverse of no_save)
    save = not no_save

    try:
        mgr = ExcelManager()

        # --force: use force_kill
        if force:
            _force_kill_instances(mgr, instance_id, all_instances, console)
            return

        # --all: stop all instances
        if all_instances:
            _stop_all_instances(mgr, save, console)
            return

        # Specific PID
        if instance_id:
            pid = int(instance_id)
            _stop_single_instance(mgr, pid, save, console)
            return

        # No argument: stop active instance
        _stop_active_instance(mgr, save, console)

    except ValueError:
        console.print(
            f"[red]Erreur :[/red] PID invalide '{instance_id}'. "
            "Le PID doit être un nombre entier.",
            style="bold",
        )
        raise typer.Exit(code=1)

    except ExcelInstanceNotFoundError as e:
        console.print(f"[red]Instance introuvable :[/red] {e}", style="bold")
        raise typer.Exit(code=1)

    except ExcelRPCError:
        console.print(
            "[red]Erreur RPC :[/red] L'instance est déconnectée ou zombie\n"
            "[yellow]Utilisez --force pour terminer le processus[/yellow]",
            style="bold",
        )
        raise typer.Exit(code=1)

    except Exception as e:
        console.print(f"[red]Erreur :[/red] {e}", style="bold")
        raise typer.Exit(code=1)


@app.command()
def status():
    """Display status of running Excel instances.

    Shows information about all currently running Excel instances including
    process ID, visibility, number of open workbooks, and window handle.
    """
    try:
        manager = ExcelManager()
        instances = manager.list_running_instances()

        if not instances:
            console.print(
                Panel.fit(
                    "[yellow]i[/yellow] No running Excel instances found",
                    title="Excel Status",
                    border_style="yellow",
                )
            )
            return

        # Create a table for instances
        table = Table(title=f"Running Excel Instances ({len(instances)} found)")
        table.add_column("PID", style="cyan", justify="right")
        table.add_column("HWND", style="magenta", justify="right")
        table.add_column("Visible", style="green")
        table.add_column("Workbooks", style="yellow", justify="right")

        for info in instances:
            visible_text = "Yes" if info.visible else "No"
            visible_color = "green" if info.visible else "red"

            table.add_row(
                str(info.pid),
                str(info.hwnd),
                f"[{visible_color}]{visible_text}[/{visible_color}]",
                str(info.workbooks_count),
            )

        console.print(table)

    except ExcelConnectionError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Failed to get Excel status\n\n[bold]Error:[/bold] {e}",
                title="Connection Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Excel management error\n\n[bold]Error:[/bold] {e}",
                title="Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except Exception as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Unexpected error\n\n[bold]Error:[/bold] {e}",
                title="Unexpected Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@app.command()
def optimize(
    screen: bool = typer.Option(
        False,
        "--screen",
        help="Optimiser uniquement l'affichage écran",
    ),
    calculation: bool = typer.Option(
        False,
        "--calculation",
        help="Optimiser uniquement le calcul",
    ),
    all_opt: bool = typer.Option(
        False,
        "--all",
        help="Appliquer toutes les optimisations",
    ),
    restore: bool = typer.Option(
        False,
        "--restore",
        help="Restaurer les paramètres originaux",
    ),
    status_opt: bool = typer.Option(
        False,
        "--status",
        help="Afficher l'état actuel des paramètres",
    ),
    force_calculate: bool = typer.Option(
        False,
        "--force-calculate",
        help="Forcer le recalcul complet du classeur actif",
    ),
    visible: bool = typer.Option(
        False,
        "--visible",
        help="Rendre Excel visible",
    ),
) -> None:
    """Optimise les performances Excel ou affiche l'état actuel.

    Par défaut (sans option), applique toutes les optimisations.

    Exemples:

        xlmanage optimize --screen

        xlmanage optimize --all

        xlmanage optimize --status

        xlmanage optimize --restore
    """
    try:
        from .calculation_optimizer import CalculationOptimizer
        from .excel_optimizer import ExcelOptimizer
        from .screen_optimizer import ScreenOptimizer
    except ImportError:
        from xlmanage.calculation_optimizer import CalculationOptimizer
        from xlmanage.excel_optimizer import ExcelOptimizer
        from xlmanage.screen_optimizer import ScreenOptimizer

    # Validation : une seule option principale à la fois
    options_count = sum(
        [screen, calculation, all_opt, restore, status_opt, force_calculate]
    )
    if options_count == 0:
        # Par défaut : --all
        all_opt = True
    elif options_count > 1:
        console.print(
            "[red]Erreur :[/red] Spécifiez une seule option parmi "
            "--screen, --calculation, --all, --restore, --status, --force-calculate",
            style="bold",
        )
        raise typer.Exit(code=1)

    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            app_com = excel_mgr.app

            # --status : afficher l'état actuel
            if status_opt:
                _display_optimization_status(excel_mgr, console)
                return

            # --restore : restaurer les paramètres
            if restore:
                _restore_optimizations(excel_mgr, screen, calculation, all_opt, console)
                return

            # --force-calculate : forcer le recalcul
            if force_calculate:
                _force_calculate(app_com, console)
                return

            # --screen : optimiser l'écran
            if screen:
                screen_opt = ScreenOptimizer(excel_mgr)
                state = screen_opt.apply()
                _display_applied_optimizations(state, console)
                return

            # --calculation : optimiser le calcul
            if calculation:
                calc_opt = CalculationOptimizer(excel_mgr)
                state = calc_opt.apply()
                _display_applied_optimizations(state, console)
                return

            # --all : tout optimiser
            if all_opt:
                excel_opt = ExcelOptimizer(excel_mgr)
                state = excel_opt.apply()
                _display_applied_optimizations(state, console)
                return

    except ExcelConnectionError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Failed to connect to Excel\n\n[bold]Error:[/bold] {e}",
                title="Connection Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except Exception as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Unexpected error\n\n[bold]Error:[/bold] {e}",
                title="Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


def _display_optimization_status(excel_mgr, console_obj: Console) -> None:
    """Affiche l'état actuel des paramètres Excel."""
    try:
        from .excel_optimizer import ExcelOptimizer
    except ImportError:
        from xlmanage.excel_optimizer import ExcelOptimizer

    optimizer = ExcelOptimizer(excel_mgr)
    settings = optimizer.get_current_settings()

    if not settings:
        console_obj.print(
            Panel.fit(
                "[yellow]i[/yellow] Impossible de lire l'état des paramètres Excel",
                title="État des optimisations",
                border_style="yellow",
            )
        )
        return

    table = Table(title="État actuel des optimisations Excel", show_header=True)
    table.add_column("Propriété", style="cyan", width=25)
    table.add_column("Valeur actuelle", width=20)
    table.add_column("Optimisé", justify="center", width=15)

    # Valeurs optimisées attendues
    optimized_values = {
        "ScreenUpdating": False,
        "DisplayStatusBar": False,
        "EnableAnimations": False,
        "Calculation": -4135,  # xlCalculationManual
        "EnableEvents": False,
        "DisplayAlerts": False,
        "AskToUpdateLinks": False,
        "Iteration": False,
    }

    for prop, value in settings.items():
        # Formater la valeur
        if isinstance(value, bool):
            value_str = "Oui" if value else "Non"
        elif prop == "Calculation":
            calc_names = {-4135: "Manuel", -4105: "Automatique"}
            value_str = calc_names.get(cast(int, value), str(value))
        else:
            value_str = str(value)

        # Vérifier si optimisé
        is_optimized = value == optimized_values.get(prop)
        optimized_str = "[green]Oui[/green]" if is_optimized else "[yellow]Non[/yellow]"

        table.add_row(prop, value_str, optimized_str)

    console_obj.print(table)


def _restore_optimizations(
    excel_mgr, screen: bool, calculation: bool, all_opt: bool, console_obj: Console
) -> None:
    """Restaure les paramètres optimisés."""
    try:
        from .calculation_optimizer import CalculationOptimizer
        from .excel_optimizer import ExcelOptimizer
        from .screen_optimizer import ScreenOptimizer
    except ImportError:
        from xlmanage.calculation_optimizer import CalculationOptimizer
        from xlmanage.excel_optimizer import ExcelOptimizer
        from xlmanage.screen_optimizer import ScreenOptimizer

    try:
        optimizer: ScreenOptimizer | CalculationOptimizer | ExcelOptimizer
        if screen:
            optimizer = ScreenOptimizer(excel_mgr)
        elif calculation:
            optimizer = CalculationOptimizer(excel_mgr)
        else:  # all_opt
            optimizer = ExcelOptimizer(excel_mgr)

        optimizer.restore()

        console_obj.print(
            Panel.fit(
                "[green]OK[/green] Paramètres restaurés avec succès",
                title="Restauration",
                border_style="green",
            )
        )

    except RuntimeError as e:
        console_obj.print(
            Panel.fit(
                f"[yellow]i[/yellow] {e}\n\n"
                "Les optimisations doivent d'abord être appliquées avec "
                "--screen, --calculation ou --all",
                title="Attention",
                border_style="yellow",
            )
        )


def _display_applied_optimizations(state, console_obj: Console) -> None:
    """Affiche un résumé des optimisations appliquées."""
    optimizer_names = {
        "screen": "Écran",
        "calculation": "Calcul",
        "all": "Toutes les optimisations",
    }

    optimizer_name = optimizer_names.get(state.optimizer_type, state.optimizer_type)

    # Compter les propriétés optimisées
    if state.optimizer_type == "screen":
        count = len(state.screen)
    elif state.optimizer_type == "calculation":
        count = len(state.calculation)
    else:
        count = len(state.full)

    message = (
        f"[green]OK[/green] Optimisations appliquées avec succès\n\n"
        f"Type : [bold]{optimizer_name}[/bold]\n"
        f"Propriétés modifiées : {count}\n"
        f"Appliqué à : {state.applied_at}\n\n"
        "[dim]Les optimisations resteront actives "
        "jusqu'à l'appel de --restore[/dim]"
    )

    console_obj.print(
        Panel.fit(
            message,
            title="Optimisation Excel",
            border_style="green",
        )
    )


def _force_calculate(app_com, console_obj: Console) -> None:
    """Force le recalcul complet du classeur actif."""
    try:
        wb = app_com.ActiveWorkbook
        if wb is None:
            console_obj.print(
                Panel.fit(
                    "[yellow]i[/yellow] Aucun classeur actif",
                    title="Recalcul",
                    border_style="yellow",
                )
            )
            return

        console_obj.print("[dim]Recalcul en cours...[/dim]")

        # Forcer le recalcul complet
        app_com.CalculateFullRebuild()

        console_obj.print(
            Panel.fit(
                f"[green]OK[/green] Recalcul complet terminé\n\n"
                f"Classeur : [bold]{wb.Name}[/bold]",
                title="Recalcul forcé",
                border_style="green",
            )
        )

    except Exception as e:
        console_obj.print(
            Panel.fit(
                f"[red]X[/red] Erreur lors du recalcul\n\n[bold]Error:[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


workbook_app = typer.Typer(help="Manage Excel workbooks")
app.add_typer(workbook_app, name="workbook")


@workbook_app.command("open")
def workbook_open(
    path: Path = typer.Argument(..., help="Path to the workbook file"),
    read_only: bool = typer.Option(
        False,
        "--read-only",
        "-r",
        help="Open in read-only mode",
    ),
):
    """Open an existing workbook.

    Opens a workbook file in the active Excel instance.
    The file must exist on disk.
    """
    try:
        with ExcelManager() as excel_mgr:
            wb_mgr = WorkbookManager(excel_mgr)
            info = wb_mgr.open(path, read_only=read_only)

            mode = "lecture seule" if info.read_only else "lecture/écriture"
            saved_status = "sauvegardé" if info.saved else "non sauvegardé"

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Classeur ouvert avec succès\n\n"
                    f"[bold]Nom :[/bold] {info.name}\n"
                    f"[bold]Chemin :[/bold] {info.full_path}\n"
                    f"[bold]Mode :[/bold] {mode}\n"
                    f"[bold]État :[/bold] {saved_status}\n"
                    f"[bold]Feuilles :[/bold] {info.sheets_count}",
                    title="Classeur ouvert",
                    border_style="green",
                )
            )

    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Fichier introuvable\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorkbookAlreadyOpenError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur déjà ouvert\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@workbook_app.command("create")
def workbook_create(
    path: Path = typer.Argument(..., help="Path for the new workbook"),
    template: Path = typer.Option(
        None,
        "--template",
        "-t",
        help="Template file to use",
    ),
):
    """Create a new workbook.

    Creates a new Excel workbook and saves it to the specified path.
    Optionally uses a template file as starting point.
    """
    try:
        with ExcelManager() as excel_mgr:
            wb_mgr = WorkbookManager(excel_mgr)
            info = wb_mgr.create(path, template=template)

            template_info = f"Basé sur : {template.name}" if template else "Vierge"

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Classeur créé avec succès\n\n"
                    f"[bold]Nom :[/bold] {info.name}\n"
                    f"[bold]Chemin :[/bold] {info.full_path}\n"
                    f"[bold]Type :[/bold] {template_info}\n"
                    f"[bold]Feuilles :[/bold] {info.sheets_count}",
                    title="Classeur créé",
                    border_style="green",
                )
            )

    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Template introuvable\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorkbookSaveError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Échec de sauvegarde\n\n"
                f"[bold]Chemin :[/bold] {e.path}\n"
                f"[bold]Raison :[/bold] {e.message}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@workbook_app.command("close")
def workbook_close(
    path: Path = typer.Argument(..., help="Path to the workbook to close"),
    save: bool = typer.Option(
        True,
        "--save/--no-save",
        help="Save changes before closing",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Force close without confirmation dialogs",
    ),
):
    """Close an open workbook.

    Closes a workbook that is currently open.
    By default, saves changes before closing.
    """
    try:
        with ExcelManager() as excel_mgr:
            wb_mgr = WorkbookManager(excel_mgr)
            wb_mgr.close(path, save=save, force=force)

            save_info = "avec sauvegarde" if save else "sans sauvegarde"

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Classeur fermé {save_info}\n\n"
                    f"[bold]Fichier :[/bold] {path.name}",
                    title="Succès",
                    border_style="green",
                )
            )

    except WorkbookNotFoundError:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non ouvert\n\n"
                f"[bold]Fichier :[/bold] {path.name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@workbook_app.command("save")
def workbook_save(
    path: Path = typer.Argument(..., help="Path to the open workbook"),
    output: Path = typer.Option(
        None,
        "--as",
        "-o",
        help="Save to a different file (SaveAs)",
    ),
):
    """Save a workbook.

    Saves an open workbook to disk.
    Use --as to save to a different file (SaveAs).
    """
    try:
        with ExcelManager() as excel_mgr:
            wb_mgr = WorkbookManager(excel_mgr)
            wb_mgr.save(path, output=output)

            if output:
                target = f"{path.name} → {output.name}"
                operation = "SaveAs"
            else:
                target = path.name
                operation = "Save"

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Classeur sauvegardé\n\n"
                    f"[bold]Opération :[/bold] {operation}\n"
                    f"[bold]Fichier :[/bold] {target}",
                    title="Succès",
                    border_style="green",
                )
            )

    except WorkbookNotFoundError:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non ouvert\n\n"
                f"[bold]Fichier :[/bold] {path.name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorkbookSaveError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Échec de sauvegarde\n\n"
                f"[bold]Chemin :[/bold] {e.path}\n"
                f"[bold]Raison :[/bold] {e.message}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@workbook_app.command("list")
def workbook_list():
    """List all open workbooks.

    Displays information about all workbooks currently open
    in the Excel instance.
    """
    try:
        with ExcelManager() as excel_mgr:
            wb_mgr = WorkbookManager(excel_mgr)
            workbooks = wb_mgr.list()

            if not workbooks:
                console.print(
                    Panel.fit(
                        "[yellow]i[/yellow] Aucun classeur ouvert",
                        title="Classeurs",
                        border_style="yellow",
                    )
                )
                return

            table = Table(title=f"Classeurs ouverts ({len(workbooks)} trouvé(s))")
            table.add_column("Nom", style="cyan")
            table.add_column("Feuilles", justify="right", style="yellow")
            table.add_column("Mode", style="magenta")
            table.add_column("État", style="green")

            for info in workbooks:
                mode = "R/O" if info.read_only else "R/W"
                mode_color = "red" if info.read_only else "green"

                saved_text = "Oui" if info.saved else "X"
                saved_color = "green" if info.saved else "yellow"

                table.add_row(
                    info.name,
                    str(info.sheets_count),
                    f"[{mode_color}]{mode}[/{mode_color}]",
                    f"[{saved_color}]{saved_text}[/{saved_color}]",
                )

            console.print(table)

    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


worksheet_app = typer.Typer(help="Manage Excel worksheets")
app.add_typer(worksheet_app, name="worksheet")


@worksheet_app.command("create")
def worksheet_create(
    name: str = typer.Argument(..., help="Name of the new worksheet"),
    workbook: Path = typer.Option(
        None,
        "--workbook",
        "-w",
        help="Path to the target workbook (defaults to active workbook)",
    ),
):
    """Create a new worksheet.

    Creates a new worksheet in the specified workbook.
    If no workbook is specified, creates it in the active workbook.
    """
    try:
        with ExcelManager() as excel_mgr:
            ws_mgr = WorksheetManager(excel_mgr)
            info = ws_mgr.create(name, workbook=workbook)

            workbook_info = (
                f"Classeur : {workbook.name}" if workbook else "Classeur actif"
            )

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Feuille créée avec succès\n\n"
                    f"[bold]Nom :[/bold] {info.name}\n"
                    f"[bold]Position :[/bold] {info.index}\n"
                    f"[bold]{workbook_info}[/bold]\n"
                    f"[bold]Visible :[/bold] {'Oui' if info.visible else 'Non'}",
                    title="Feuille créée",
                    border_style="green",
                )
            )

    except WorksheetNameError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Nom de feuille invalide\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Raison :[/bold] {e.reason}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorksheetAlreadyExistsError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Feuille déjà existante\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Classeur :[/bold] {e.workbook_name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non trouvé\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@worksheet_app.command("delete")
def worksheet_delete(
    name: str = typer.Argument(..., help="Name of the worksheet to delete"),
    workbook: Path = typer.Option(
        None,
        "--workbook",
        "-w",
        help="Path to the target workbook (defaults to active workbook)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Force deletion without confirmation",
    ),
):
    """Delete a worksheet.

    Deletes the specified worksheet from the workbook.
    By default, asks for confirmation before deleting.
    """
    try:
        # Confirmation (sauf si --force)
        if not force:
            workbook_info = (
                f"dans le classeur '{workbook.name}'"
                if workbook
                else "dans le classeur actif"
            )
            console.print(
                f"\n[yellow]Attention :[/yellow] Vous allez supprimer la feuille "
                f"'{name}' {workbook_info}"
            )
            confirm = typer.confirm("Êtes-vous sûr de vouloir continuer ?")
            if not confirm:
                console.print("[yellow]Opération annulée[/yellow]")
                return

        with ExcelManager() as excel_mgr:
            ws_mgr = WorksheetManager(excel_mgr)
            ws_mgr.delete(name, workbook=workbook)

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Feuille supprimée avec succès\n\n"
                    f"[bold]Nom :[/bold] {name}",
                    title="Succès",
                    border_style="green",
                )
            )

    except WorksheetNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Feuille introuvable\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Classeur :[/bold] {e.workbook_name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorksheetDeleteError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Suppression impossible\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Raison :[/bold] {e.reason}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non trouvé\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@worksheet_app.command("list")
def worksheet_list(
    workbook: Path = typer.Option(
        None,
        "--workbook",
        "-w",
        help="Path to the target workbook (defaults to active workbook)",
    ),
):
    """List all worksheets in a workbook.

    Displays information about all worksheets including position,
    visibility, and data dimensions.
    """
    try:
        with ExcelManager() as excel_mgr:
            ws_mgr = WorksheetManager(excel_mgr)
            worksheets = ws_mgr.list(workbook=workbook)

            if not worksheets:
                console.print(
                    Panel.fit(
                        "[yellow]i[/yellow] Aucune feuille trouvée",
                        title="Feuilles",
                        border_style="yellow",
                    )
                )
                return

            workbook_info = f" - {workbook.name}" if workbook else " - Classeur actif"
            title = f"Feuilles du classeur ({len(worksheets)} trouvée(s))"
            table = Table(title=f"{title}{workbook_info}")
            table.add_column("Position", justify="right", style="cyan")
            table.add_column("Nom", style="yellow")
            table.add_column("Visible", style="green")
            table.add_column("Lignes", justify="right", style="magenta")
            table.add_column("Colonnes", justify="right", style="magenta")

            for info in worksheets:
                visible_text = "Oui" if info.visible else "X"
                visible_color = "green" if info.visible else "red"

                table.add_row(
                    str(info.index),
                    info.name,
                    f"[{visible_color}]{visible_text}[/{visible_color}]",
                    str(info.rows_used),
                    str(info.columns_used),
                )

            console.print(table)

    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non trouvé\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@worksheet_app.command("copy")
def worksheet_copy(
    source: str = typer.Argument(..., help="Name of the source worksheet"),
    destination: str = typer.Argument(..., help="Name for the copy"),
    workbook: Path = typer.Option(
        None,
        "--workbook",
        "-w",
        help="Path to the target workbook (defaults to active workbook)",
    ),
):
    """Copy a worksheet.

    Creates a copy of the source worksheet with a new name.
    The copy is placed immediately after the source worksheet.
    """
    try:
        with ExcelManager() as excel_mgr:
            ws_mgr = WorksheetManager(excel_mgr)
            info = ws_mgr.copy(source, destination, workbook=workbook)

            workbook_info = (
                f"Classeur : {workbook.name}" if workbook else "Classeur actif"
            )

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Feuille copiée avec succès\n\n"
                    f"[bold]Source :[/bold] {source}\n"
                    f"[bold]Destination :[/bold] {info.name}\n"
                    f"[bold]Position :[/bold] {info.index}\n"
                    f"[bold]{workbook_info}[/bold]",
                    title="Feuille copiée",
                    border_style="green",
                )
            )

    except WorksheetNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Feuille source introuvable\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Classeur :[/bold] {e.workbook_name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorksheetNameError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Nom de destination invalide\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Raison :[/bold] {e.reason}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorksheetAlreadyExistsError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Feuille de destination déjà existante\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Classeur :[/bold] {e.workbook_name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non trouvé\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


table_app = typer.Typer(help="Manage Excel tables")
app.add_typer(table_app, name="table")


@table_app.command("create")
def table_create(
    name: str = typer.Argument(..., help="Name of the new table"),
    range_ref: str = typer.Argument(..., help="Range reference (e.g., 'A1:D100')"),
    worksheet: str = typer.Option(
        None,
        "--worksheet",
        "-ws",
        help="Target worksheet name (defaults to active worksheet)",
    ),
    workbook: Path = typer.Option(
        None,
        "--workbook",
        "-w",
        help="Path to the target workbook (defaults to active workbook)",
    ),
):
    """Create a new table.

    Creates a new Excel table (ListObject) in the specified worksheet.
    The table must have a valid name and range reference.
    """
    try:
        with ExcelManager() as excel_mgr:
            table_mgr = TableManager(excel_mgr)
            info = table_mgr.create(
                name, range_ref, worksheet=worksheet, workbook=workbook
            )

            workbook_info = (
                f"Classeur : {workbook.name}" if workbook else "Classeur actif"
            )

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Table créée avec succès\n\n"
                    f"[bold]Nom :[/bold] {info.name}\n"
                    f"[bold]Feuille :[/bold] {info.worksheet_name}\n"
                    f"[bold]Plage :[/bold] {info.range_address}\n"
                    f"[bold]Lignes :[/bold] {info.rows_count}\n"
                    f"[bold]{workbook_info}[/bold]",
                    title="Table créée",
                    border_style="green",
                )
            )

    except TableNameError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Nom de table invalide\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Raison :[/bold] {e.reason}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except TableRangeError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Plage invalide\n\n"
                f"[bold]Plage :[/bold] {e.range_ref}\n"
                f"[bold]Raison :[/bold] {e.reason}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except TableAlreadyExistsError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Table déjà existante\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Classeur :[/bold] {e.workbook_name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorksheetNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Feuille introuvable\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Classeur :[/bold] {e.workbook_name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non trouvé\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@table_app.command("delete")
def table_delete(
    name: str = typer.Argument(..., help="Name of the table to delete"),
    worksheet: str = typer.Option(
        None,
        "--worksheet",
        "-ws",
        help="Worksheet containing the table (searches all if not specified)",
    ),
    workbook: Path = typer.Option(
        None,
        "--workbook",
        "-w",
        help="Path to the target workbook (defaults to active workbook)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Force deletion without confirmation",
    ),
):
    """Delete a table.

    Deletes the specified table from the workbook.
    By default, asks for confirmation before deleting.
    """
    try:
        # Confirmation (sauf si --force)
        if not force:
            workbook_info = (
                f"dans le classeur '{workbook.name}'"
                if workbook
                else "dans le classeur actif"
            )
            worksheet_info = (
                f"de la feuille '{worksheet}'"
                if worksheet
                else "de n'importe quelle feuille"
            )
            console.print(
                f"\n[yellow]Attention :[/yellow] Vous allez supprimer la table "
                f"'{name}' {worksheet_info} {workbook_info}"
            )
            confirm = typer.confirm("Êtes-vous sûr de vouloir continuer ?")
            if not confirm:
                console.print("[yellow]Opération annulée[/yellow]")
                return

        with ExcelManager() as excel_mgr:
            table_mgr = TableManager(excel_mgr)
            table_mgr.delete(name, worksheet=worksheet, workbook=workbook)

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Table supprimée avec succès\n\n"
                    f"[bold]Nom :[/bold] {name}",
                    title="Succès",
                    border_style="green",
                )
            )

    except TableNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Table introuvable\n\n"
                f"[bold]Nom :[/bold] {e.name}\n"
                f"[bold]Feuille :[/bold] {e.worksheet_name}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non trouvé\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@table_app.command("list")
def table_list(
    worksheet: str = typer.Option(
        None,
        "--worksheet",
        "-ws",
        help="Worksheet to list tables from (lists all if not specified)",
    ),
    workbook: Path = typer.Option(
        None,
        "--workbook",
        "-w",
        help="Path to the target workbook (defaults to active workbook)",
    ),
):
    """List all tables.

    Displays information about all tables in the workbook or worksheet.
    If no worksheet is specified, lists tables from all worksheets.
    """
    try:
        with ExcelManager() as excel_mgr:
            table_mgr = TableManager(excel_mgr)
            tables = table_mgr.list(worksheet=worksheet, workbook=workbook)

            if not tables:
                console.print(
                    Panel.fit(
                        "[yellow]i[/yellow] Aucune table trouvée",
                        title="Tables",
                        border_style="yellow",
                    )
                )
                return

            workbook_info = f" - {workbook.name}" if workbook else " - Classeur actif"
            worksheet_info = f" - Feuille '{worksheet}'" if worksheet else ""
            title = f"Tables ({len(tables)} trouvée(s))"
            table = Table(title=f"{title}{workbook_info}{worksheet_info}")
            table.add_column("Nom", style="cyan")
            table.add_column("Feuille", style="yellow")
            table.add_column("Plage", style="magenta")
            table.add_column("Lignes", justify="right", style="green")

            for info in tables:
                table.add_row(
                    info.name,
                    info.worksheet_name,
                    info.range_address,
                    str(info.rows_count),
                )

            console.print(table)

    except WorkbookNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Classeur non trouvé\n\n[bold]Chemin :[/bold] {e.path}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


# ============================================================================
# VBA Commands
# ============================================================================

vba_app = typer.Typer(help="Manage VBA modules")
app.add_typer(vba_app, name="vba")


@vba_app.command("import")
def vba_import(
    module_file: Path = typer.Argument(
        ..., help="Chemin du fichier module (.bas, .cls, .frm)"
    ),
    module_type: str = typer.Option(
        None,
        "--type",
        "-t",
        help="Type de module (standard|class|userform). Auto-détecté si omis",
    ),
    workbook: Path = typer.Option(
        None, "--workbook", "-w", help="Classeur cible (actif si omis)"
    ),
    overwrite: bool = typer.Option(
        False, "--overwrite", help="Remplacer le module s'il existe déjà"
    ),
    visible: bool = typer.Option(False, "--visible", help="Rendre Excel visible"),
):
    """Importe un module VBA depuis un fichier.

    Exemples:

        xlmanage vba import Module1.bas

        xlmanage vba import MyClass.cls --workbook data.xlsm --overwrite

        xlmanage vba import UserForm1.frm --type userform
    """
    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            vba_mgr = VBAManager(excel_mgr)

            # Importer le module
            info = vba_mgr.import_module(
                module_file=module_file,
                module_type=module_type,
                workbook=workbook,
                overwrite=overwrite,
            )

            # Affichage du succès
            console.print(
                Panel(
                    f"[green]OK[/green] Module VBA importé avec succès\n\n"
                    f"[bold]Nom :[/bold] {info.name}\n"
                    f"[bold]Type :[/bold] {info.module_type}\n"
                    f"[bold]Lignes :[/bold] {info.lines_count}\n"
                    f"[bold]PredeclaredId :[/bold] "
                    f"{'Oui' if info.has_predeclared_id else 'Non'}",
                    title="Import VBA",
                    border_style="green",
                )
            )

    except VBAProjectAccessError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur d'accès VBA\n\n"
                f"[bold]Détails :[/bold] {e}\n\n"
                f"[yellow]Solution :[/yellow] Activez l'option "
                "'Trust access to the VBA project object model' dans Excel :\n"
                "File > Options > Trust Center > Trust Center Settings > "
                "Macro Settings",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)

    except VBAWorkbookFormatError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Format de classeur invalide\n\n"
                f"[bold]Détails :[/bold] {e}\n\n"
                f"[yellow]Solution :[/yellow] Convertissez le classeur au format .xlsm "
                "pour supporter les macros.",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)

    except VBAModuleAlreadyExistsError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Module existant\n\n"
                f"[bold]Module :[/bold] {e.module_name}\n"
                f"[bold]Classeur :[/bold] {e.workbook_name}\n\n"
                f"[yellow]Utilisez --overwrite pour le remplacer[/yellow]",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)

    except VBAImportError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur d'import\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)

    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@vba_app.command("export")
def vba_export(
    module_name: str = typer.Argument(..., help="Nom du module à exporter"),
    output_file: Path = typer.Argument(..., help="Fichier de destination"),
    workbook: Path = typer.Option(
        None, "--workbook", "-w", help="Classeur source (actif si omis)"
    ),
    visible: bool = typer.Option(False, "--visible", help="Rendre Excel visible"),
):
    """Exporte un module VBA vers un fichier.

    Exemples:

        xlmanage vba export Module1 backup/Module1.bas

        xlmanage vba export ThisWorkbook ThisWorkbook.cls --workbook data.xlsm
    """
    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            vba_mgr = VBAManager(excel_mgr)

            # Exporter le module
            exported_path = vba_mgr.export_module(
                module_name=module_name, output_file=output_file, workbook=workbook
            )

            # Affichage du succès
            console.print(
                Panel(
                    f"[green]OK[/green] Module VBA exporté avec succès\n\n"
                    f"[bold]Module :[/bold] {module_name}\n"
                    f"[bold]Fichier :[/bold] {exported_path}",
                    title="Export VBA",
                    border_style="green",
                )
            )

    except VBAModuleNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Module introuvable\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)

    except VBAExportError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur d'export\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)

    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@vba_app.command("list")
def vba_list(
    workbook: Path = typer.Option(
        None, "--workbook", "-w", help="Classeur à analyser (actif si omis)"
    ),
    visible: bool = typer.Option(False, "--visible", help="Rendre Excel visible"),
):
    """Liste tous les modules VBA d'un classeur.

    Exemples:

        xlmanage vba list

        xlmanage vba list --workbook data.xlsm
    """
    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            vba_mgr = VBAManager(excel_mgr)

            # Lister les modules
            modules = vba_mgr.list_modules(workbook=workbook)

            if not modules:
                console.print(
                    Panel.fit(
                        "[yellow]i[/yellow] Aucun module VBA trouvé",
                        title="Modules VBA",
                        border_style="yellow",
                    )
                )
                return

            # Créer un tableau Rich
            workbook_info = f" - {workbook.name}" if workbook else " - Classeur actif"
            table = Table(title=f"Modules VBA{workbook_info}")
            table.add_column("Nom", style="cyan", width=30)
            table.add_column("Type", style="yellow", width=15)
            table.add_column("Lignes", justify="right", style="green", width=10)
            table.add_column(
                "PredeclaredId", justify="center", style="magenta", width=15
            )

            for module in modules:
                predeclared = "Oui" if module.has_predeclared_id else "-"
                table.add_row(
                    module.name,
                    module.module_type,
                    str(module.lines_count),
                    predeclared,
                )

            console.print(table)
            console.print(f"\n[dim]Total : {len(modules)} module(s)[/dim]")

    except VBAProjectAccessError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur d'accès VBA\n\n"
                f"[bold]Détails :[/bold] {e}\n\n"
                f"[yellow]Solution :[/yellow] Activez l'option "
                "'Trust access to the VBA project object model' dans Excel :\n"
                "File > Options > Trust Center > Trust Center Settings > "
                "Macro Settings",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)

    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


@vba_app.command("delete")
def vba_delete(
    module_name: str = typer.Argument(..., help="Nom du module à supprimer"),
    workbook: Path = typer.Option(
        None, "--workbook", "-w", help="Classeur cible (actif si omis)"
    ),
    force: bool = typer.Option(False, "--force", help="Pas de confirmation (réservé)"),
    visible: bool = typer.Option(False, "--visible", help="Rendre Excel visible"),
):
    """Supprime un module VBA.

    Attention: Les modules de document (ThisWorkbook, Sheet1, etc.) ne peuvent
    pas être supprimés.

    Exemples:

        xlmanage vba delete Module1

        xlmanage vba delete MyClass --workbook data.xlsm
    """
    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            vba_mgr = VBAManager(excel_mgr)

            # Supprimer le module
            vba_mgr.delete_module(
                module_name=module_name, workbook=workbook, force=force
            )

            # Affichage du succès
            console.print(
                Panel(
                    f"[green]OK[/green] Module VBA supprimé avec succès\n\n"
                    f"[bold]Module :[/bold] {module_name}",
                    title="Suppression VBA",
                    border_style="green",
                )
            )

    except VBAModuleNotFoundError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        if "Cannot delete document module" in str(e):
            console.print(
                "\n[yellow]Les modules de document (ThisWorkbook, Sheet1, etc.) "
                "font partie du classeur et ne peuvent pas être supprimés.[/yellow]"
            )
        raise typer.Exit(code=1)

    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Erreur\n\n[bold]Détails :[/bold] {e}",
                title="Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


def _display_macro_result(result: MacroResult, console_obj: Console) -> None:
    """Affiche le résultat d'exécution d'une macro avec Rich.

    Args:
        result: Résultat de MacroRunner.run()
        console_obj: Console Rich pour l'affichage
    """
    if not result.success:
        # Affichage erreur
        console_obj.print(
            Panel(
                f"[red]{result.error_message}[/red]",
                title=f"❌ Erreur lors de l'exécution de {result.macro_name}",
                border_style="red",
            )
        )
        return

    # Affichage succès
    if result.return_value is None:
        # Sub VBA (pas de retour)
        console_obj.print(
            Panel(
                "[green]La macro a été exécutée avec succès.[/green]\n"
                "[dim]Aucune valeur de retour (probablement un Sub VBA)[/dim]",
                title=f"✅ {result.macro_name}",
                border_style="green",
            )
        )
    else:
        # Function VBA avec retour
        # Formater la valeur
        formatted_value = _format_return_value(result.return_value)

        # Créer une table pour affichage structuré
        table = Table(show_header=False, box=None, padding=(0, 2))
        table.add_row("[bold]Type:[/bold]", f"[cyan]{result.return_type}[/cyan]")
        table.add_row("[bold]Valeur:[/bold]", f"[green]{formatted_value}[/green]")

        console_obj.print(
            Panel(
                table,
                title=f"✅ {result.macro_name}",
                border_style="green",
            )
        )


@app.command()
def run_macro(
    macro_name: str = typer.Argument(
        ..., help="Nom de la macro VBA à exécuter (ex: 'Module1.MySub' ou 'MySub')"
    ),
    workbook: str | None = typer.Option(
        None,
        "--workbook",
        "-w",
        help=(
            "Chemin du classeur contenant la macro "
            "(optionnel, cherche dans actif + PERSONAL.XLSB sinon)"
        ),
    ),
    args: str | None = typer.Option(
        None,
        "--args",
        "-a",
        help="Arguments CSV pour la macro (ex: '\"hello\",42,3.14,true')",
    ),
    timeout: int = typer.Option(
        60,
        "--timeout",
        "-t",
        help="Timeout d'exécution en secondes (défaut: 60s)",
    ),
) -> None:
    """Exécute une macro VBA (Sub ou Function) avec arguments optionnels.

    Cette commande permet de lancer des macros VBA depuis la ligne de commande
    et d'afficher le résultat (valeur de retour pour les Function).

    \b
    Exemples:
      xlmanage run-macro "Module1.SayHello"
      xlmanage run-macro "Module1.GetSum" --args "10,20"
      xlmanage run-macro "Module1.Process" -w "data.xlsm" -a '"Report",true'
      xlmanage run-macro "Module1.LongTask" --timeout 120

    \b
    Format des arguments (--args):
      Les arguments sont en format CSV avec conversion automatique de types:
      - Chaînes: "hello" ou 'world'
      - Nombres entiers: 42, -10
      - Nombres décimaux: 3.14, -0.5
      - Booléens: true, false (case-insensitive)
      - Exemple: '"Report_2024",100,true,3.5'
    """
    try:
        # Convertir workbook en Path si fourni
        workbook_path: Path | None = None
        if workbook:
            workbook_path = Path(workbook)
            if not workbook_path.exists():
                console.print(
                    f"[red]✗[/red] Fichier introuvable: {workbook}", style="red"
                )
                raise typer.Exit(code=1)

        # Se connecter à Excel (réutiliser instance active ou créer)
        with ExcelManager() as mgr:
            try:
                # Essayer de se connecter à une instance active
                existing = mgr.get_running_instance()
                if existing:
                    console.print(
                        f"[blue]→[/blue] Connexion à l'instance Excel existante "
                        f"(PID {existing.pid})"
                    )
                else:
                    # Démarrer une nouvelle instance
                    console.print(
                        "[blue]→[/blue] Démarrage d'une nouvelle instance Excel..."
                    )
                    mgr.start(new=False)

            except ExcelConnectionError as e:
                console.print(
                    f"[red]✗[/red] Impossible de se connecter à Excel: {e.message}",
                    style="red",
                )
                raise typer.Exit(code=1)

            # Créer le runner et exécuter la macro
            runner = MacroRunner(mgr)

            console.print(f"[blue]→[/blue] Exécution de [bold]{macro_name}[/bold]...")

            # Exécuter avec timeout (via signal ou threading selon OS)
            # Pour simplifier, on exécute directement ici
            # TODO: Implémenter timeout réel dans une version ultérieure
            result = runner.run(
                macro_name=macro_name, workbook=workbook_path, args=args
            )

            # Afficher le résultat
            _display_macro_result(result, console)

            # Exit code selon succès
            if not result.success:
                raise typer.Exit(code=1)

    except VBAMacroError as e:
        console.print(
            Panel(
                f"[red]Erreur VBA:[/red] {e.reason}",
                title="❌ Échec d'exécution",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)

    except WorkbookNotFoundError as e:
        console.print(f"[red]✗[/red] Classeur non ouvert: {e.path.name}", style="red")
        raise typer.Exit(code=1)

    except Exception as e:
        console.print(
            Panel(
                f"[red]Erreur inattendue:[/red] {str(e)}",
                title="❌ Erreur",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


def main_entry():
    """Main entry point for xlmanage CLI."""
    app()


if __name__ == "__main__":
    main_entry()
