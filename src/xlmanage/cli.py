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

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

try:
    from .excel_manager import ExcelManager
    from .exceptions import (
        ExcelConnectionError,
        ExcelManageError,
        WorkbookAlreadyOpenError,
        WorkbookNotFoundError,
        WorkbookSaveError,
        WorksheetAlreadyExistsError,
        WorksheetDeleteError,
        WorksheetNameError,
        WorksheetNotFoundError,
    )
    from .workbook_manager import WorkbookManager
    from .worksheet_manager import WorksheetManager
except ImportError:
    from xlmanage.excel_manager import ExcelManager
    from xlmanage.exceptions import (
        ExcelConnectionError,
        ExcelManageError,
        WorkbookAlreadyOpenError,
        WorkbookNotFoundError,
        WorkbookSaveError,
        WorksheetAlreadyExistsError,
        WorksheetDeleteError,
        WorksheetNameError,
        WorksheetNotFoundError,
    )
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


@app.command()
def stop(
    all_instances: bool = typer.Option(
        False,
        "--all",
        "-a",
        help="Stop all running Excel instances",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Force stop without confirmation",
    ),
    no_save: bool = typer.Option(
        False,
        "--no-save",
        help="Do not save workbooks before closing",
    ),
):
    """Stop Excel instance(s).

    By default, stops the active Excel instance with save prompt.
    Use --all to stop all running Excel instances.
    Use --force to skip confirmation prompts.
    Use --no-save to close without saving workbooks.
    """
    try:
        manager = ExcelManager()

        if all_instances:
            # Get all running instances
            instances = manager.list_running_instances()

            if not instances:
                console.print(
                    Panel.fit(
                        "[yellow]i[/yellow] No running Excel instances found",
                        title="Information",
                        border_style="yellow",
                    )
                )
                return

            # Confirm if not forced
            if not force:
                msg = f"\n[yellow]Warning:[/yellow] About to stop {len(instances)}"
                msg += " Excel instance(s)"
                console.print(msg)
                confirm = typer.confirm("Are you sure you want to continue?")
                if not confirm:
                    console.print("[yellow]Operation cancelled[/yellow]")
                    return

            # Stop all instances
            stopped_count = 0
            for info in instances:
                try:
                    # Connect to each instance and stop it
                    # Note: This is a simplified approach
                    # In production, we'd need to connect by PID or HWND
                    manager.stop(save=not no_save)
                    stopped_count += 1
                except Exception:
                    # Continue with next instance even if one fails
                    continue

            console.print(
                Panel.fit(
                    f"[green]OK[/green] Stopped {stopped_count} Excel instance(s)",
                    title="Success",
                    border_style="green",
                )
            )

        else:
            # Stop current instance
            if not force and not no_save:
                confirm = typer.confirm(
                    "Stop the current Excel instance? Workbooks will be saved."
                )
                if not confirm:
                    console.print("[yellow]Operation cancelled[/yellow]")
                    return

            manager.stop(save=not no_save)
            console.print(
                Panel.fit(
                    "[green]OK[/green] Excel instance stopped successfully",
                    title="Success",
                    border_style="green",
                )
            )

    except ExcelConnectionError as e:
        console.print(
            Panel.fit(
                f"[red]X[/red] Failed to stop Excel instance\n\n"
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


def main_entry():
    """Main entry point for xlmanage CLI."""
    app()


if __name__ == "__main__":
    main_entry()
