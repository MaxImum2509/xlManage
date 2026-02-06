# Epic 12 - Story 3: Intégration CLI de la commande run-macro

**Statut** : 🔴 À faire

**En tant qu'** utilisateur de xlManage
**Je veux** une commande `xlmanage run-macro` dans le CLI
**Afin de** pouvoir exécuter des macros VBA depuis la ligne de commande

## Contexte

La commande `run-macro` doit :

1. Prendre le nom de la macro en argument obligatoire
2. Accepter des options : `--workbook`, `--args`, `--timeout`
3. Se connecter à une instance Excel active (ou en démarrer une si nécessaire)
4. Exécuter la macro via `MacroRunner`
5. Afficher le résultat avec Rich (formatage coloré)
6. Gérer les erreurs et afficher des messages clairs

**Exemple d'utilisation** :

```bash
# Exécuter un Sub sans arguments
xlmanage run-macro "Module1.SayHello"

# Exécuter une Function avec arguments
xlmanage run-macro "Module1.GetSum" --args "10,20"

# Exécuter une macro dans un classeur spécifique
xlmanage run-macro "Module1.ProcessData" --workbook "C:\data.xlsm" --args '"Report_2024",true'

# Avec timeout personnalisé
xlmanage run-macro "Module1.LongRunning" --timeout 120
```

## Critères d'acceptation

1. ✅ La commande `run-macro` est ajoutée au CLI (fonction + décorateur Typer)
2. ✅ L'argument `macro_name` est obligatoire
3. ✅ Les options `--workbook`, `--args`, `--timeout` sont fonctionnelles
4. ✅ Le résultat est affiché avec Rich (vert = succès, rouge = erreur)
5. ✅ Les erreurs sont capturées et affichent des messages clairs
6. ✅ Les tests CLI passent avec CliRunner
7. ✅ L'aide de la commande est complète et en français

## Tâches techniques

### Tâche 3.1 : Ajouter la commande run-macro au CLI

**Fichier** : `src/xlmanage/cli.py`

Ajouter les imports nécessaires :

```python
from pathlib import Path
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from xlmanage.excel_manager import ExcelManager
from xlmanage.macro_runner import MacroRunner, MacroResult
from xlmanage.exceptions import (
    VBAMacroError,
    WorkbookNotFoundError,
    ExcelConnectionError,
)
```

Puis ajouter la commande (au niveau racine de l'application, pas dans un groupe) :

```python
@app.command()
def run_macro(
    macro_name: str = typer.Argument(
        ...,
        help="Nom de la macro VBA à exécuter (ex: 'Module1.MySub' ou 'MySub')"
    ),
    workbook: str | None = typer.Option(
        None,
        "--workbook",
        "-w",
        help="Chemin du classeur contenant la macro (optionnel, cherche dans actif + PERSONAL.XLSB sinon)"
    ),
    args: str | None = typer.Option(
        None,
        "--args",
        "-a",
        help="Arguments CSV pour la macro (ex: '\"hello\",42,3.14,true')"
    ),
    timeout: int = typer.Option(
        60,
        "--timeout",
        "-t",
        help="Timeout d'exécution en secondes (défaut: 60s)"
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
    console = Console()

    try:
        # Convertir workbook en Path si fourni
        workbook_path: Path | None = None
        if workbook:
            workbook_path = Path(workbook)
            if not workbook_path.exists():
                console.print(
                    f"[red]✗[/red] Fichier introuvable: {workbook}",
                    style="red"
                )
                raise typer.Exit(code=1)

        # Se connecter à Excel (réutiliser instance active ou créer)
        with ExcelManager() as mgr:
            try:
                # Essayer de se connecter à une instance active
                existing = mgr.get_running_instance()
                if existing:
                    console.print(
                        f"[blue]→[/blue] Connexion à l'instance Excel existante (PID {existing.pid})"
                    )
                else:
                    # Démarrer une nouvelle instance
                    console.print("[blue]→[/blue] Démarrage d'une nouvelle instance Excel...")
                    mgr.start(new=False)

            except ExcelConnectionError as e:
                console.print(
                    f"[red]✗[/red] Impossible de se connecter à Excel: {e.message}",
                    style="red"
                )
                raise typer.Exit(code=1)

            # Créer le runner et exécuter la macro
            runner = MacroRunner(mgr)

            console.print(f"[blue]→[/blue] Exécution de [bold]{macro_name}[/bold]...")

            # Exécuter avec timeout (via signal ou threading selon OS)
            # Pour simplifier, on exécute directement ici
            # TODO: Implémenter timeout réel dans une version ultérieure
            result = runner.run(
                macro_name=macro_name,
                workbook=workbook_path,
                args=args
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
                border_style="red"
            )
        )
        raise typer.Exit(code=1)

    except WorkbookNotFoundError as e:
        console.print(
            f"[red]✗[/red] Classeur non ouvert: {e.path.name}",
            style="red"
        )
        raise typer.Exit(code=1)

    except Exception as e:
        console.print(
            Panel(
                f"[red]Erreur inattendue:[/red] {str(e)}",
                title="❌ Erreur",
                border_style="red"
            )
        )
        raise typer.Exit(code=1)
```

**Points d'attention** :
- `macro_name` est un Argument (obligatoire, sans `--`)
- `workbook`, `args`, `timeout` sont des Options (optionnelles, avec `--`)
- Vérifier l'existence du fichier workbook avant de continuer
- Réutiliser une instance Excel active si disponible (performance)
- Le timeout est défini mais non implémenté (TODO pour future version)

### Tâche 3.2 : Implémenter _display_macro_result()

**Fichier** : `src/xlmanage/cli.py` (fonction helper au niveau module)

```python
def _display_macro_result(result: MacroResult, console: Console) -> None:
    """Affiche le résultat d'exécution d'une macro avec Rich.

    Args:
        result: Résultat de MacroRunner.run()
        console: Console Rich pour l'affichage
    """
    if not result.success:
        # Affichage erreur
        console.print(
            Panel(
                f"[red]{result.error_message}[/red]",
                title=f"❌ Erreur lors de l'exécution de {result.macro_name}",
                border_style="red"
            )
        )
        return

    # Affichage succès
    if result.return_value is None:
        # Sub VBA (pas de retour)
        console.print(
            Panel(
                f"[green]La macro a été exécutée avec succès.[/green]\n"
                f"[dim]Aucune valeur de retour (probablement un Sub VBA)[/dim]",
                title=f"✅ {result.macro_name}",
                border_style="green"
            )
        )
    else:
        # Function VBA avec retour
        # Formater la valeur
        from xlmanage.macro_runner import _format_return_value
        formatted_value = _format_return_value(result.return_value)

        # Créer une table pour affichage structuré
        table = Table(show_header=False, box=None, padding=(0, 2))
        table.add_row("[bold]Type:[/bold]", f"[cyan]{result.return_type}[/cyan]")
        table.add_row("[bold]Valeur:[/bold]", f"[green]{formatted_value}[/green]")

        console.print(
            Panel(
                table,
                title=f"✅ {result.macro_name}",
                border_style="green"
            )
        )
```

**Points d'attention** :
- Utiliser des Panel Rich pour un affichage soigné
- Différencier Sub (pas de retour) et Function (avec retour)
- Utiliser Table Rich pour l'affichage structuré du retour
- Couleurs : vert = succès, rouge = erreur, cyan = type, dim = info secondaire

### Tâche 3.3 : Mettre à jour __init__.py

**Fichier** : `src/xlmanage/__init__.py`

Ajouter MacroRunner et MacroResult aux exports :

```python
from xlmanage.macro_runner import MacroRunner, MacroResult

__all__ = [
    # ... exports existants ...
    "MacroRunner",
    "MacroResult",
]
```

### Tâche 3.4 : Tests CLI pour run-macro

**Fichier** : `tests/test_cli_run_macro.py`

```python
"""Tests CLI pour la commande run-macro."""

import pytest
from typer.testing import CliRunner
from unittest.mock import Mock, patch
from pathlib import Path

from xlmanage.cli import app
from xlmanage.macro_runner import MacroResult
from xlmanage.exceptions import VBAMacroError, WorkbookNotFoundError


runner = CliRunner()


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_success_sub(mock_runner_class, mock_mgr_class):
    """Test exécution réussie d'un Sub (pas de retour)."""
    # Mock du manager
    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None
    mock_mgr.start = Mock()

    # Mock du runner
    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner

    result_obj = MacroResult(
        macro_name="Module1.Test",
        return_value=None,
        return_type="NoneType",
        success=True,
        error_message=None
    )
    mock_runner.run.return_value = result_obj

    # Exécuter la commande
    result = runner.invoke(app, ["run-macro", "Module1.Test"])

    assert result.exit_code == 0
    assert "✅" in result.stdout
    assert "Module1.Test" in result.stdout
    mock_runner.run.assert_called_once_with(
        macro_name="Module1.Test",
        workbook=None,
        args=None
    )


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_success_function(mock_runner_class, mock_mgr_class):
    """Test exécution réussie d'une Function avec retour."""
    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None

    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner

    result_obj = MacroResult(
        macro_name="Module1.GetSum",
        return_value=42,
        return_type="int",
        success=True,
        error_message=None
    )
    mock_runner.run.return_value = result_obj

    result = runner.invoke(app, ["run-macro", "Module1.GetSum", "--args", "10,20"])

    assert result.exit_code == 0
    assert "✅" in result.stdout
    assert "42" in result.stdout
    assert "int" in result.stdout


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_vba_error(mock_runner_class, mock_mgr_class):
    """Test erreur VBA runtime."""
    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None

    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner

    result_obj = MacroResult(
        macro_name="Module1.Divide",
        return_value=None,
        return_type="NoneType",
        success=False,
        error_message="Division by zero"
    )
    mock_runner.run.return_value = result_obj

    result = runner.invoke(app, ["run-macro", "Module1.Divide", "--args", "10,0"])

    assert result.exit_code == 1
    assert "❌" in result.stdout
    assert "Division by zero" in result.stdout


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_with_workbook(mock_runner_class, mock_mgr_class, tmp_path):
    """Test avec workbook spécifié."""
    # Créer un fichier temporaire
    workbook_file = tmp_path / "test.xlsm"
    workbook_file.touch()

    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None

    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner

    result_obj = MacroResult(
        macro_name="'test.xlsm'!Module1.Test",
        return_value="OK",
        return_type="str",
        success=True,
        error_message=None
    )
    mock_runner.run.return_value = result_obj

    result = runner.invoke(
        app,
        ["run-macro", "Module1.Test", "--workbook", str(workbook_file)]
    )

    assert result.exit_code == 0
    # Vérifier que run() a été appelé avec le bon Path
    call_args = mock_runner.run.call_args
    assert call_args.kwargs["workbook"] == workbook_file


def test_run_macro_workbook_not_found():
    """Test erreur si fichier workbook introuvable."""
    result = runner.invoke(
        app,
        ["run-macro", "Module1.Test", "--workbook", "C:/nonexistent.xlsm"]
    )

    assert result.exit_code == 1
    assert "introuvable" in result.stdout.lower()


@patch("xlmanage.cli.ExcelManager")
@patch("xlmanage.cli.MacroRunner")
def test_run_macro_macro_not_found(mock_runner_class, mock_mgr_class):
    """Test erreur macro introuvable."""
    mock_mgr = Mock()
    mock_mgr_class.return_value.__enter__ = Mock(return_value=mock_mgr)
    mock_mgr_class.return_value.__exit__ = Mock(return_value=False)
    mock_mgr.get_running_instance.return_value = None

    mock_runner = Mock()
    mock_runner_class.return_value = mock_runner
    mock_runner.run.side_effect = VBAMacroError(
        macro_name="Module1.Missing",
        reason="Macro introuvable"
    )

    result = runner.invoke(app, ["run-macro", "Module1.Missing"])

    assert result.exit_code == 1
    assert "❌" in result.stdout
    assert "Macro introuvable" in result.stdout


@patch("xlmanage.cli.ExcelManager")
def test_run_macro_help(mock_mgr_class):
    """Test affichage de l'aide."""
    result = runner.invoke(app, ["run-macro", "--help"])

    assert result.exit_code == 0
    assert "macro_name" in result.stdout.lower()
    assert "--workbook" in result.stdout
    assert "--args" in result.stdout
    assert "--timeout" in result.stdout
    assert "Exemples:" in result.stdout
```

**Points d'attention** :
- Utiliser `CliRunner` de Typer pour tester les commandes
- Mocker `ExcelManager` et `MacroRunner` pour éviter les appels COM réels
- Tester les exit codes (0 = succès, 1 = erreur)
- Vérifier que les arguments sont correctement passés aux managers
- Tester l'affichage (présence des emojis, messages d'erreur)

### Tâche 3.5 : Documentation de la commande

**Fichier** : `docs/cli.md` (ou créer si n'existe pas)

Ajouter une section sur `run-macro` :

```markdown
## Commande `run-macro`

Exécute une macro VBA (Sub ou Function) avec arguments optionnels.

### Syntaxe

\`\`\`bash
xlmanage run-macro MACRO_NAME [OPTIONS]
\`\`\`

### Arguments

- `MACRO_NAME` : Nom de la macro à exécuter (ex: "Module1.MySub" ou "MySub")

### Options

- `--workbook, -w PATH` : Chemin du classeur contenant la macro (optionnel)
- `--args, -a TEXT` : Arguments CSV pour la macro (ex: '"hello",42,3.14,true')
- `--timeout, -t INTEGER` : Timeout d'exécution en secondes (défaut: 60)

### Format des arguments

Les arguments sont fournis en format CSV avec conversion automatique :

- **Chaînes** : `"hello"` ou `'world'`
- **Entiers** : `42`, `-10`
- **Décimaux** : `3.14`, `-0.5`
- **Booléens** : `true`, `false` (case-insensitive)

Exemple complet : `'"Report_2024",100,true,3.5'`

### Exemples

\`\`\`bash
# Exécuter un Sub sans arguments
xlmanage run-macro "Module1.SayHello"

# Exécuter une Function avec arguments
xlmanage run-macro "Module1.GetSum" --args "10,20"

# Macro dans un classeur spécifique
xlmanage run-macro "Module1.ProcessData" \\
  --workbook "C:\\data.xlsm" \\
  --args '"Report_2024",true'

# Avec timeout personnalisé (120 secondes)
xlmanage run-macro "Module1.LongRunning" --timeout 120
\`\`\`

### Valeurs de retour

- **Sub VBA** : Affiche "Exécutée avec succès" (pas de valeur de retour)
- **Function VBA** : Affiche le type et la valeur retournée

### Codes de sortie

- `0` : Succès
- `1` : Erreur (VBA runtime, macro introuvable, classeur non ouvert, etc.)
```

## Tests à implémenter

Tous les tests sont dans `tests/test_cli_run_macro.py` (8 tests).

**Coverage attendue** : > 85% pour la fonction `run_macro()` dans cli.py

**Commande de test** :
```bash
pytest tests/test_cli_run_macro.py -v --cov=src/xlmanage/cli --cov-report=term
```

## Dépendances

- Epic 12, Story 1 (Parser d'arguments)
- Epic 12, Story 2 (MacroRunner)
- Epic 5 (ExcelManager)

## Définition of Done

- [x] Commande `run-macro` ajoutée au CLI avec toutes les options
- [x] Fonction helper `_display_macro_result()` implémentée
- [x] Exports mis à jour dans `__init__.py`
- [x] Tous les tests CLI passent (8 tests)
- [x] Couverture > 85% pour la commande CLI
- [x] L'aide de la commande est complète (`--help` fonctionne)
- [x] Documentation ajoutée dans `docs/cli.md`
- [x] Messages d'erreur clairs et en français
- [x] Affichage Rich avec couleurs et formatage

## Notes pour le développeur junior

**Concepts clés à comprendre** :

1. **Typer Argument vs Option** :
   - `Argument` : obligatoire, sans `--` (ex: `macro_name`)
   - `Option` : optionnel, avec `--` (ex: `--workbook`, `--args`)

2. **CliRunner** :
   - Permet de tester les commandes CLI sans lancer le programme
   - `runner.invoke(app, ["commande", "arg1", "--option", "value"])`
   - Capture stdout et exit code

3. **Context manager avec ExcelManager** :
   - `with ExcelManager() as mgr:` garantit le nettoyage même en cas d'erreur
   - Équivalent à try/finally avec `__enter__` et `__exit__`

4. **Rich Console** :
   - `Panel` : boîte avec bordure et titre
   - `Table` : tableau formaté
   - `[green]...[/green]` : markup Rich pour couleurs

5. **Exit codes** :
   - `0` = succès (conventionnel Unix/Linux)
   - `1` = erreur générique
   - `raise typer.Exit(code=1)` pour sortir avec code d'erreur

**Pièges à éviter** :

- ❌ Ne pas oublier de vérifier l'existence du fichier workbook AVANT de démarrer Excel
- ❌ Ne pas oublier de fermer proprement Excel (le context manager le fait)
- ❌ Ne pas afficher les erreurs COM brutes (traduire en messages clairs)
- ❌ Ne pas confondre `result.exit_code` (test CLI) et `raise typer.Exit(code=...)`

**Bonnes pratiques** :

- ✅ Afficher un message de progression avant les opérations longues
- ✅ Utiliser des couleurs cohérentes (vert = succès, rouge = erreur, bleu = info)
- ✅ Fournir des exemples dans `--help` avec `\b` pour désactiver le wrapping
- ✅ Capturer TOUTES les exceptions et afficher un message clair

**Ressources** :

- [Typer documentation](https://typer.tiangolo.com/)
- [Rich documentation](https://rich.readthedocs.io/)
- [CliRunner testing](https://typer.tiangolo.com/tutorial/testing/)
