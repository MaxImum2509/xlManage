# Epic 10 - Story 2: Intégrer les commandes optimize dans le CLI

**Statut** : ⏳ À faire

**En tant que** utilisateur
**Je veux** gérer les optimisations Excel depuis le CLI
**Afin de** accélérer mes traitements Excel ou voir l'état actuel des optimisations

## Critères d'acceptation

1. ⬜ La commande `xlmanage optimize --screen` applique les optimisations d'écran
2. ⬜ La commande `xlmanage optimize --calculation` applique les optimisations de calcul
3. ⬜ La commande `xlmanage optimize --all` applique toutes les optimisations
4. ⬜ La commande `xlmanage optimize --restore` restaure les paramètres originaux
5. ⬜ La commande `xlmanage optimize --status` affiche l'état actuel
6. ⬜ La commande `xlmanage optimize --force-calculate` force un recalcul complet
7. ⬜ Les tests CLI passent pour toutes les options

## Tâches techniques

### Tâche 2.1 : Implémenter `xlmanage optimize` avec les options

**Fichier** : `src/xlmanage/cli.py`

La commande existe déjà en stub. Il faut l'implémenter :

```python
@app.command()
def optimize(
    screen: Annotated[
        bool,
        typer.Option("--screen", help="Optimiser uniquement l'affichage écran")
    ] = False,
    calculation: Annotated[
        bool,
        typer.Option("--calculation", help="Optimiser uniquement le calcul")
    ] = False,
    all_opt: Annotated[
        bool,
        typer.Option("--all", help="Appliquer toutes les optimisations")
    ] = False,
    restore: Annotated[
        bool,
        typer.Option("--restore", help="Restaurer les paramètres originaux")
    ] = False,
    status: Annotated[
        bool,
        typer.Option("--status", help="Afficher l'état actuel des paramètres")
    ] = False,
    force_calculate: Annotated[
        bool,
        typer.Option("--force-calculate", help="Forcer le recalcul complet du classeur actif")
    ] = False,
    visible: Annotated[bool, typer.Option("--visible", help="Rendre Excel visible")] = False,
) -> None:
    """Optimise les performances Excel ou affiche l'état actuel.

    Par défaut (sans option), applique toutes les optimisations.

    Exemples:

        xlmanage optimize --screen

        xlmanage optimize --all

        xlmanage optimize --status

        xlmanage optimize --restore
    """
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel

    from .excel_manager import ExcelManager
    from .excel_optimizer import ExcelOptimizer, OptimizationState
    from .screen_optimizer import ScreenOptimizer
    from .calculation_optimizer import CalculationOptimizer

    console = Console()

    # Validation : une seule option principale à la fois
    options_count = sum([screen, calculation, all_opt, restore, status, force_calculate])
    if options_count == 0:
        # Par défaut : --all
        all_opt = True
    elif options_count > 1:
        console.print(
            "[red]Erreur :[/red] Spécifiez une seule option parmi "
            "--screen, --calculation, --all, --restore, --status, --force-calculate",
            style="bold"
        )
        raise typer.Exit(code=1)

    try:
        with ExcelManager(visible=visible) as excel_mgr:
            excel_mgr.start()
            app = excel_mgr.app

            # --status : afficher l'état actuel
            if status:
                _display_optimization_status(app, console)
                return

            # --restore : restaurer les paramètres
            if restore:
                _restore_optimizations(app, screen, calculation, all_opt, console)
                return

            # --force-calculate : forcer le recalcul
            if force_calculate:
                _force_calculate(app, console)
                return

            # --screen : optimiser l'écran
            if screen:
                optimizer = ScreenOptimizer(app)
                state = optimizer.apply()
                _display_applied_optimizations(state, console)
                return

            # --calculation : optimiser le calcul
            if calculation:
                optimizer = CalculationOptimizer(app)
                state = optimizer.apply()
                _display_applied_optimizations(state, console)
                return

            # --all : tout optimiser
            if all_opt:
                optimizer = ExcelOptimizer(app)
                state = optimizer.apply()
                _display_applied_optimizations(state, console)
                return

    except Exception as e:
        console.print(f"[red]Erreur :[/red] {e}", style="bold")
        raise typer.Exit(code=1)
```

### Tâche 2.2 : Implémenter _display_optimization_status()

Fonction helper pour afficher l'état actuel :

```python
def _display_optimization_status(app, console: Console) -> None:
    """Affiche l'état actuel des paramètres Excel."""
    optimizer = ExcelOptimizer(app)
    settings = optimizer.get_current_settings()

    table = Table(title="État actuel des optimisations Excel", show_header=True)
    table.add_column("Propriété", style="cyan", width=25)
    table.add_column("Valeur actuelle", width=20)
    table.add_column("Optimisé", justify="center", width=15)

    # Valeurs optimisées attendues
    optimized_values = {
        'ScreenUpdating': False,
        'DisplayStatusBar': False,
        'EnableAnimations': False,
        'Calculation': -4135,  # xlCalculationManual
        'EnableEvents': False,
        'DisplayAlerts': False,
        'AskToUpdateLinks': False,
        'Iteration': False,
    }

    for prop, value in settings.items():
        # Formater la valeur
        if isinstance(value, bool):
            value_str = "Oui" if value else "Non"
        elif prop == 'Calculation':
            calc_names = {-4135: "Manuel", -4105: "Automatique"}
            value_str = calc_names.get(value, str(value))
        else:
            value_str = str(value)

        # Vérifier si optimisé
        is_optimized = (value == optimized_values.get(prop))
        optimized_str = "[green]Oui[/green]" if is_optimized else "[yellow]Non[/yellow]"

        table.add_row(prop, value_str, optimized_str)

    console.print(table)
```

### Tâche 2.3 : Implémenter _restore_optimizations()

```python
def _restore_optimizations(
    app,
    screen: bool,
    calculation: bool,
    all_opt: bool,
    console: Console
) -> None:
    """Restaure les paramètres optimisés."""
    try:
        if screen:
            optimizer = ScreenOptimizer(app)
        elif calculation:
            optimizer = CalculationOptimizer(app)
        else:  # all_opt
            optimizer = ExcelOptimizer(app)

        optimizer.restore()

        console.print(Panel(
            "[green]Paramètres restaurés avec succès[/green]",
            title="Restauration",
            border_style="green"
        ))

    except RuntimeError as e:
        console.print(
            f"[yellow]Attention :[/yellow] {e}\\n"
            "Les optimisations doivent d'abord être appliquées avec --screen, --calculation ou --all",
            style="dim"
        )
```

### Tâche 2.4 : Implémenter _display_applied_optimizations()

```python
def _display_applied_optimizations(state: OptimizationState, console: Console) -> None:
    """Affiche un résumé des optimisations appliquées."""
    optimizer_names = {
        "screen": "Écran",
        "calculation": "Calcul",
        "all": "Toutes les optimisations"
    }

    optimizer_name = optimizer_names.get(state.optimizer_type, state.optimizer_type)

    # Compter les propriétés optimisées
    if state.optimizer_type == "screen":
        count = len(state.screen)
    elif state.optimizer_type == "calculation":
        count = len(state.calculation)
    else:
        count = len(state.full)

    console.print(Panel(
        f"[green]Optimisations appliquées avec succès[/green]\\n\\n"
        f"Type : [bold]{optimizer_name}[/bold]\\n"
        f"Propriétés modifiées : {count}\\n"
        f"Appliqué à : {state.applied_at}\\n\\n"
        f"[dim]Les optimisations resteront actives jusqu'à l'appel de --restore[/dim]",
        title="Optimisation Excel",
        border_style="green"
    ))
```

### Tâche 2.5 : Implémenter _force_calculate()

```python
def _force_calculate(app, console: Console) -> None:
    """Force le recalcul complet du classeur actif."""
    try:
        wb = app.ActiveWorkbook
        if wb is None:
            console.print("[yellow]Aucun classeur actif[/yellow]")
            return

        console.print("[dim]Recalcul en cours...[/dim]")

        # Forcer le recalcul complet
        app.CalculateFullRebuild()

        console.print(Panel(
            f"[green]Recalcul complet terminé[/green]\\n\\n"
            f"Classeur : [bold]{wb.Name}[/bold]",
            title="Recalcul forcé",
            border_style="green"
        ))

    except Exception as e:
        console.print(f"[red]Erreur lors du recalcul :[/red] {e}", style="bold")
```

**Points d'attention** :
- `CalculateFullRebuild()` force un recalcul complet (plus lent que `Calculate()`)
- Il faut vérifier qu'un classeur est actif avant d'appeler cette méthode

## Tests à implémenter

Créer `tests/test_cli_optimize.py` :

```python
import pytest
from typer.testing import CliRunner
from unittest.mock import Mock, patch

from xlmanage.cli import app
from xlmanage.excel_optimizer import OptimizationState

runner = CliRunner()


def test_optimize_screen():
    """Test optimize --screen command."""
    mock_state = OptimizationState(
        screen={'ScreenUpdating': True},
        calculation={},
        full={},
        applied_at="2026-02-05T10:00:00",
        optimizer_type="screen"
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.ScreenOptimizer") as mock_opt_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.apply.return_value = mock_state
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--screen"])

        assert result.exit_code == 0
        assert "appliquées avec succès" in result.stdout
        mock_opt.apply.assert_called_once()


def test_optimize_calculation():
    """Test optimize --calculation command."""
    mock_state = OptimizationState(
        screen={},
        calculation={'Calculation': -4135},
        full={},
        applied_at="2026-02-05T10:00:00",
        optimizer_type="calculation"
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.CalculationOptimizer") as mock_opt_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.apply.return_value = mock_state
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--calculation"])

        assert result.exit_code == 0
        mock_opt.apply.assert_called_once()


def test_optimize_all():
    """Test optimize --all command."""
    mock_state = OptimizationState(
        screen={},
        calculation={},
        full={'ScreenUpdating': True, 'DisplayAlerts': True},
        applied_at="2026-02-05T10:00:00",
        optimizer_type="all"
    )

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.ExcelOptimizer") as mock_opt_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.apply.return_value = mock_state
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--all"])

        assert result.exit_code == 0


def test_optimize_status():
    """Test optimize --status command."""
    mock_settings = {
        'ScreenUpdating': True,
        'DisplayAlerts': False,
        'Calculation': -4135,
    }

    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.ExcelOptimizer") as mock_opt_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.get_current_settings.return_value = mock_settings
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--status"])

        assert result.exit_code == 0
        assert "État actuel" in result.stdout


def test_optimize_restore():
    """Test optimize --restore command."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, \\
         patch("xlmanage.cli.ExcelOptimizer") as mock_opt_class:

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_opt = Mock()
        mock_opt.restore.return_value = None
        mock_opt_class.return_value = mock_opt

        result = runner.invoke(app, ["optimize", "--restore"])

        assert result.exit_code == 0
        assert "restaurés" in result.stdout


def test_optimize_force_calculate():
    """Test optimize --force-calculate command."""
    with patch("xlmanage.cli.ExcelManager") as mock_mgr_class:

        mock_wb = Mock()
        mock_wb.Name = "test.xlsx"

        mock_app = Mock()
        mock_app.ActiveWorkbook = mock_wb
        mock_app.CalculateFullRebuild = Mock()

        mock_mgr = Mock()
        mock_mgr.app = mock_app
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        result = runner.invoke(app, ["optimize", "--force-calculate"])

        assert result.exit_code == 0
        assert "Recalcul complet terminé" in result.stdout
        mock_app.CalculateFullRebuild.assert_called_once()


def test_optimize_multiple_options_error():
    """Test error when multiple options are specified."""
    result = runner.invoke(app, ["optimize", "--screen", "--calculation"])

    assert result.exit_code == 1
    assert "une seule option" in result.stdout
```

## Dépendances

- Epic 10, Story 1 (apply/restore dans les optimizers)

## Définition of Done

- [ ] La commande `optimize` gère toutes les options
- [ ] `--status` affiche un tableau Rich avec l'état actuel
- [ ] `--restore` restaure les paramètres avec gestion d'erreur
- [ ] `--force-calculate` force le recalcul complet
- [ ] Les options mutuellement exclusives sont validées
- [ ] Tous les tests CLI passent (7+ tests)
- [ ] Les messages Rich sont clairs et informatifs
- [ ] L'aide CLI est complète avec exemples
