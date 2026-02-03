"""
CLI module for xlmanage.
"""

import typer


app = typer.Typer(name="xlmanage", help="Excel automation CLI tool")


@app.command()
def version():
    """Show version information."""
    typer.echo("xlmanage version 0.1.0")


def main_entry():
    """Main entry point for xlmanage CLI."""
    app()


if __name__ == "__main__":
    main_entry()