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

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from .excel_manager import ExcelManager
from .exceptions import ExcelConnectionError, ExcelManageError

app = typer.Typer(name="xlmanage", help="Excel automation CLI tool")
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

        # Display success message
        mode = "new" if new else "existing"
        visibility = "visible" if visible else "hidden"

        console.print(
            Panel.fit(
                f"[green]✓[/green] Excel instance started successfully\n\n"
                f"[bold]Mode:[/bold] {mode}\n"
                f"[bold]Visibility:[/bold] {visibility}\n"
                f"[bold]Process ID:[/bold] {info.pid}\n"
                f"[bold]Window Handle:[/bold] {info.hwnd}\n"
                f"[bold]Workbooks:[/bold] {info.workbooks_count}",
                title="Excel Instance Started",
                border_style="green",
            )
        )

    except ExcelConnectionError as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Failed to start Excel instance\n\n"
                f"[bold]Error:[/bold] {e}",
                title="Connection Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Excel management error\n\n[bold]Error:[/bold] {e}",
                title="Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except Exception as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Unexpected error\n\n[bold]Error:[/bold] {e}",
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
                        "[yellow]ℹ[/yellow] No running Excel instances found",
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
                    f"[green]✓[/green] Stopped {stopped_count} Excel instance(s)",
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
                    "[green]✓[/green] Excel instance stopped successfully",
                    title="Success",
                    border_style="green",
                )
            )

    except ExcelConnectionError as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Failed to stop Excel instance\n\n"
                f"[bold]Error:[/bold] {e}",
                title="Connection Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Excel management error\n\n[bold]Error:[/bold] {e}",
                title="Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except Exception as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Unexpected error\n\n[bold]Error:[/bold] {e}",
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
                    "[yellow]ℹ[/yellow] No running Excel instances found",
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
            visible_icon = "✓" if info.visible else "✗"
            visible_color = "green" if info.visible else "red"

            table.add_row(
                str(info.pid),
                str(info.hwnd),
                f"[{visible_color}]{visible_icon}[/{visible_color}]",
                str(info.workbooks_count),
            )

        console.print(table)

    except ExcelConnectionError as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Failed to get Excel status\n\n[bold]Error:[/bold] {e}",
                title="Connection Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except ExcelManageError as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Excel management error\n\n[bold]Error:[/bold] {e}",
                title="Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)
    except Exception as e:
        console.print(
            Panel.fit(
                f"[red]✗[/red] Unexpected error\n\n[bold]Error:[/bold] {e}",
                title="Unexpected Error",
                border_style="red",
            )
        )
        raise typer.Exit(code=1)


def main_entry():
    """Main entry point for xlmanage CLI."""
    app()


if __name__ == "__main__":
    main_entry()
