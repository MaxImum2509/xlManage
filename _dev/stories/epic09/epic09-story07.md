# Epic 9 - Story 7: Intégrer les commandes VBA dans le CLI

**Statut** : ✅ Terminé

**En tant que** utilisateur
**Je veux** utiliser les commandes VBA depuis le terminal
**Afin de** gérer mes modules VBA via l'interface en ligne de commande

## Critères d'acceptation

1. ✅ Les 4 commandes `vba` sont implémentées dans `cli.py`
2. ✅ `xlmanage vba import` importe un fichier de module
3. ✅ `xlmanage vba export` exporte un module vers un fichier
4. ✅ `xlmanage vba list` affiche la liste des modules
5. ✅ `xlmanage vba delete` supprime un module
6. ✅ Les erreurs sont affichées proprement avec Rich
7. ✅ Les tests CLI passent pour toutes les commandes

## Tâches techniques

### Tâche 7.1 : Implémenter `xlmanage vba import`

**Fichier** : `src/xlmanage/cli.py`

La commande existe déjà en stub. Il faut l'implémenter :

```python
@vba_app.command("import")
def vba_import(
    module_file: Annotated[Path, typer.Argument(help="Chemin du fichier module (.bas, .cls, .frm)")],
    module_type: Annotated[
        Optional[str],
        typer.Option("--type", "-t", help="Type de module (standard|class|userform). Auto-détecté si omis")
    ] = None,
    workbook: Annotated[
        Optional[Path],
        typer.Option("--workbook", "-w", help="Classeur cible (actif si omis)")
    ] = None,
    overwrite: Annotated[
        bool,
        typer.Option("--overwrite", help="Remplacer le module s'il existe déjà")
    ] = False,
    visible: Annotated[bool, typer.Option("--visible", help="Rendre Excel visible")] = False,
) -> None:
    """Importe un module VBA depuis un fichier.

    Exemples:

        xlmanage vba import Module1.bas

        xlmanage vba import MyClass.cls --workbook data.xlsm --overwrite

        xlmanage vba import UserForm1.frm --type userform
    """
    from rich.console import Console
    from rich.panel import Panel

    from .excel_manager import ExcelManager
    from .vba_manager import VBAManager
    from .exceptions import (
        VBAImportError,
        VBAModuleAlreadyExistsError,
        VBAProjectAccessError,
        VBAWorkbookFormatError,
    )

    console = Console()

    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            vba_mgr = VBAManager(excel_mgr)

            # Importer le module
            info = vba_mgr.import_module(
                module_file=module_file,
                module_type=module_type,
                workbook=workbook,
                overwrite=overwrite
            )

            # Affichage du succès
            console.print(Panel(
                f"[green]Module VBA importé avec succès[/green]\\n\\n"
                f"Nom : [bold]{info.name}[/bold]\\n"
                f"Type : {info.module_type}\\n"
                f"Lignes : {info.lines_count}\\n"
                f"PredeclaredId : {'Oui' if info.has_predeclared_id else 'Non'}",
                title="Import VBA",
                border_style="green"
            ))

    except VBAProjectAccessError as e:
        console.print(f"[red]Erreur d'accès VBA :[/red] {e}", style="bold")
        console.print(
            "\\n[yellow]Solution :[/yellow] Activez l'option "
            "'Trust access to the VBA project object model' dans Excel :\\n"
            "File > Options > Trust Center > Trust Center Settings > Macro Settings"
        )
        raise typer.Exit(code=1)

    except VBAWorkbookFormatError as e:
        console.print(f"[red]Format de classeur invalide :[/red] {e}", style="bold")
        console.print(
            "\\n[yellow]Solution :[/yellow] Convertissez le classeur au format .xlsm "
            "pour supporter les macros."
        )
        raise typer.Exit(code=1)

    except VBAModuleAlreadyExistsError as e:
        console.print(
            f"[red]Module existant :[/red] Le module '{e.module_name}' existe déjà "
            f"dans '{e.workbook_name}'",
            style="bold"
        )
        console.print("\\n[yellow]Utilisez --overwrite pour le remplacer[/yellow]")
        raise typer.Exit(code=1)

    except VBAImportError as e:
        console.print(f"[red]Erreur d'import :[/red] {e}", style="bold")
        raise typer.Exit(code=1)
```

**Points d'attention** :
- Le context manager `with ExcelManager()` garantit l'arrêt propre
- Les erreurs VBA ont des messages d'aide spécifiques (Trust Center, format, etc.)
- L'affichage utilise Rich Panel pour un rendu structuré

### Tâche 7.2 : Implémenter `xlmanage vba export`

```python
@vba_app.command("export")
def vba_export(
    module_name: Annotated[str, typer.Argument(help="Nom du module à exporter")],
    output_file: Annotated[Path, typer.Argument(help="Fichier de destination")],
    workbook: Annotated[
        Optional[Path],
        typer.Option("--workbook", "-w", help="Classeur source (actif si omis)")
    ] = None,
    visible: Annotated[bool, typer.Option("--visible", help="Rendre Excel visible")] = False,
) -> None:
    """Exporte un module VBA vers un fichier.

    Exemples:

        xlmanage vba export Module1 backup/Module1.bas

        xlmanage vba export ThisWorkbook ThisWorkbook.cls --workbook data.xlsm
    """
    from rich.console import Console
    from rich.panel import Panel

    from .excel_manager import ExcelManager
    from .vba_manager import VBAManager
    from .exceptions import VBAModuleNotFoundError, VBAExportError

    console = Console()

    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            vba_mgr = VBAManager(excel_mgr)

            # Exporter le module
            exported_path = vba_mgr.export_module(
                module_name=module_name,
                output_file=output_file,
                workbook=workbook
            )

            # Affichage du succès
            console.print(Panel(
                f"[green]Module VBA exporté avec succès[/green]\\n\\n"
                f"Module : [bold]{module_name}[/bold]\\n"
                f"Fichier : {exported_path}",
                title="Export VBA",
                border_style="green"
            ))

    except VBAModuleNotFoundError as e:
        console.print(f"[red]Module introuvable :[/red] {e}", style="bold")
        raise typer.Exit(code=1)

    except VBAExportError as e:
        console.print(f"[red]Erreur d'export :[/red] {e}", style="bold")
        raise typer.Exit(code=1)
```

### Tâche 7.3 : Implémenter `xlmanage vba list`

```python
@vba_app.command("list")
def vba_list(
    workbook: Annotated[
        Optional[Path],
        typer.Option("--workbook", "-w", help="Classeur à analyser (actif si omis)")
    ] = None,
    visible: Annotated[bool, typer.Option("--visible", help="Rendre Excel visible")] = False,
) -> None:
    """Liste tous les modules VBA d'un classeur.

    Exemples:

        xlmanage vba list

        xlmanage vba list --workbook data.xlsm
    """
    from rich.console import Console
    from rich.table import Table

    from .excel_manager import ExcelManager
    from .vba_manager import VBAManager

    console = Console()

    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            vba_mgr = VBAManager(excel_mgr)

            # Lister les modules
            modules = vba_mgr.list_modules(workbook=workbook)

            if not modules:
                console.print("[yellow]Aucun module VBA trouvé[/yellow]")
                return

            # Créer un tableau Rich
            table = Table(title="Modules VBA", show_header=True, header_style="bold cyan")
            table.add_column("Nom", style="bold", width=30)
            table.add_column("Type", width=15)
            table.add_column("Lignes", justify="right", width=10)
            table.add_column("PredeclaredId", justify="center", width=15)

            for module in modules:
                predeclared = "Oui" if module.has_predeclared_id else "-"
                table.add_row(
                    module.name,
                    module.module_type,
                    str(module.lines_count),
                    predeclared
                )

            console.print(table)
            console.print(f"\\n[dim]Total : {len(modules)} module(s)[/dim]")

    except Exception as e:
        console.print(f"[red]Erreur :[/red] {e}", style="bold")
        raise typer.Exit(code=1)
```

**Points d'attention** :
- Utilisation de Rich Table pour un affichage structuré
- La colonne "PredeclaredId" affiche "-" pour les modules non-classe
- Le total est affiché en bas du tableau

### Tâche 7.4 : Implémenter `xlmanage vba delete`

```python
@vba_app.command("delete")
def vba_delete(
    module_name: Annotated[str, typer.Argument(help="Nom du module à supprimer")],
    workbook: Annotated[
        Optional[Path],
        typer.Option("--workbook", "-w", help="Classeur cible (actif si omis)")
    ] = None,
    force: Annotated[
        bool,
        typer.Option("--force", help="Pas de confirmation (réservé)")
    ] = False,
    visible: Annotated[bool, typer.Option("--visible", help="Rendre Excel visible")] = False,
) -> None:
    """Supprime un module VBA.

    Attention: Les modules de document (ThisWorkbook, Sheet1, etc.) ne peuvent
    pas être supprimés.

    Exemples:

        xlmanage vba delete Module1

        xlmanage vba delete MyClass --workbook data.xlsm
    """
    from rich.console import Console
    from rich.panel import Panel

    from .excel_manager import ExcelManager
    from .vba_manager import VBAManager
    from .exceptions import VBAModuleNotFoundError

    console = Console()

    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            vba_mgr = VBAManager(excel_mgr)

            # Supprimer le module
            vba_mgr.delete_module(
                module_name=module_name,
                workbook=workbook,
                force=force
            )

            # Affichage du succès
            console.print(Panel(
                f"[green]Module VBA supprimé avec succès[/green]\\n\\n"
                f"Module : [bold]{module_name}[/bold]",
                title="Suppression VBA",
                border_style="green"
            ))

    except VBAModuleNotFoundError as e:
        console.print(f"[red]Erreur :[/red] {e}", style="bold")
        if "Cannot delete document module" in str(e):
            console.print(
                "\\n[yellow]Les modules de document (ThisWorkbook, Sheet1, etc.) "
                "font partie du classeur et ne peuvent pas être supprimés.[/yellow]"
            )
        raise typer.Exit(code=1)
```

## Tests à implémenter

Créer `tests/test_cli_vba.py` :

```python
import pytest
from pathlib import Path
from typer.testing import CliRunner
from unittest.mock import Mock, patch

from xlmanage.cli import app
from xlmanage.vba_manager import VBAModuleInfo

runner = CliRunner()


def test_vba_import_success(tmp_path):
    """Test vba import command success."""
    bas_file = tmp_path / "Module1.bas"
    bas_file.write_text("Sub Test()\\nEnd Sub", encoding='windows-1252')

    mock_info = VBAModuleInfo(
        name="Module1",
        module_type="standard",
        lines_count=2,
        has_predeclared_id=False
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.VBAManager") as mock_vba_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_vba = Mock()
        mock_vba.import_module.return_value = mock_info
        mock_vba_class.return_value = mock_vba

        result = runner.invoke(app, ["vba", "import", str(bas_file)])

        assert result.exit_code == 0
        assert "importé avec succès" in result.stdout
        assert "Module1" in result.stdout


def test_vba_export_success(tmp_path):
    """Test vba export command success."""
    output_file = tmp_path / "Module1.bas"

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.VBAManager") as mock_vba_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_vba = Mock()
        mock_vba.export_module.return_value = output_file
        mock_vba_class.return_value = mock_vba

        result = runner.invoke(app, ["vba", "export", "Module1", str(output_file)])

        assert result.exit_code == 0
        assert "exporté avec succès" in result.stdout


def test_vba_list_success():
    """Test vba list command success."""
    mock_modules = [
        VBAModuleInfo("Module1", "standard", 42, False),
        VBAModuleInfo("MyClass", "class", 15, True),
        VBAModuleInfo("ThisWorkbook", "document", 8, False),
    ]

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.VBAManager") as mock_vba_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_vba = Mock()
        mock_vba.list_modules.return_value = mock_modules
        mock_vba_class.return_value = mock_vba

        result = runner.invoke(app, ["vba", "list"])

        assert result.exit_code == 0
        assert "Module1" in result.stdout
        assert "MyClass" in result.stdout
        assert "Total : 3 module(s)" in result.stdout


def test_vba_delete_success():
    """Test vba delete command success."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.VBAManager") as mock_vba_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_vba = Mock()
        mock_vba.delete_module.return_value = None
        mock_vba_class.return_value = mock_vba

        result = runner.invoke(app, ["vba", "delete", "Module1"])

        assert result.exit_code == 0
        assert "supprimé avec succès" in result.stdout
```

## Dépendances

- Epic 9, Stories 1-6 (toutes les méthodes VBAManager)

## Définition of Done

- [x] Les 4 commandes VBA sont implémentées dans `cli.py`
- [x] Chaque commande gère les erreurs avec messages Rich appropriés
- [x] `vba list` affiche un tableau Rich structuré
- [x] Les messages d'aide sont clairs pour les erreurs Trust Center et format
- [x] Tous les tests CLI passent (18 tests)
- [x] L'aide CLI (`xlmanage vba --help`) est complète
- [x] Les exemples dans les docstrings sont fonctionnels
