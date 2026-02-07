# Epic 13 - Story 4: Integrer la commande `run-macro` dans le CLI et completer `__init__.py`

**Statut** : A faire

**Priorite** : P0 (CLI-001) + P2 (INIT-001, INIT-002)

**En tant qu'** utilisateur de xlManage
**Je veux** la commande `run-macro` dans le CLI et des exports complets dans `__init__.py`
**Afin de** pouvoir executer des macros VBA depuis la ligne de commande et importer tous les modules publics

## Contexte

L'audit du 2026-02-06 a identifie que 20 des 21 commandes CLI sont implementees (95%). Seule `run-macro` est absente. Cette commande est deja specifiee dans l'Epic 12 Story 3 (statut: A faire).

De plus, `__init__.py` est incomplet : 8 exports manquent (VBAManager, VBAModuleInfo, MacroRunner, MacroResult, ExcelOptimizer, ScreenOptimizer, CalculationOptimizer, OptimizationState), et les aliases confus (`WorkbookInfoClass`, `WorksheetInfoData`, `TableInfoData`) doivent etre supprimes.

**Reference architecture** : section 4.9 (`_dev/architecture.md`, lignes 1335-1388)

## Anomalies a corriger

| ID       | Severite  | Description                                          |
| -------- | --------- | ---------------------------------------------------- |
| CLI-001  | Critique  | Commande `run-macro` manquante                       |
| INIT-001 | Important | 8 exports manquants dans `__init__.py`               |
| INIT-002 | Mineur    | Aliases inutiles dans `__init__.py`                  |

## Taches techniques

### Tache 4.1 : Implementer la commande `run-macro` dans `cli.py` (CLI-001)

**Fichier** : `src/xlmanage/cli.py`

Suivre la specification de l'Epic 12 Story 3 (`_dev/stories/epic12/epic12-story03.md`).

La commande doit etre ajoutee au niveau racine de l'application (pas dans un sous-groupe) :

```python
@app.command()
def run_macro(
    macro_name: str = typer.Argument(
        ...,
        help="Nom de la macro VBA (ex: 'Module1.MySub')"
    ),
    workbook: Path | None = typer.Option(
        None,
        "--workbook", "-w",
        help="Classeur contenant la macro (actif si omis)"
    ),
    args: str | None = typer.Option(
        None,
        "--args", "-a",
        help="Arguments CSV (ex: '\"hello\",42,true')"
    ),
    timeout: int = typer.Option(
        60,
        "--timeout", "-t",
        help="Timeout en secondes (defaut: 60)"
    ),
) -> None:
    """Execute une macro VBA (Sub ou Function).

    Exemples:

        xlmanage run-macro "Module1.SayHello"

        xlmanage run-macro "Module1.GetSum" --args "10,20"

        xlmanage run-macro "Module1.Process" -w "data.xlsm" -a '"Report",true'
    """
    ...
```

**Pattern d'implementation** : meme pattern que les autres commandes (couche mince, try/except, affichage Rich).

**Imports a ajouter** dans `cli.py` :
```python
from .macro_runner import MacroRunner, MacroResult, _format_return_value
```

### Tache 4.2 : Implementer `_display_macro_result()` dans `cli.py`

Fonction helper pour l'affichage Rich du resultat d'execution :
- Sub VBA (return_value=None) : Panel vert "Executee avec succes"
- Function VBA (return_value!=None) : Panel vert avec type et valeur formatee
- Erreur VBA (success=False) : Panel rouge avec message d'erreur

Voir `_dev/stories/epic12/epic12-story03.md` Tache 3.2 pour la specification complete.

### Tache 4.3 : Completer les exports dans `__init__.py` (INIT-001)

**Fichier** : `src/xlmanage/__init__.py`

Ajouter les imports et exports manquants :

```python
from .vba_manager import VBAManager, VBAModuleInfo
from .macro_runner import MacroRunner, MacroResult
from .excel_optimizer import ExcelOptimizer, OptimizationState
from .screen_optimizer import ScreenOptimizer
from .calculation_optimizer import CalculationOptimizer
```

Ajouter dans `__all__` :
```python
"VBAManager",
"VBAModuleInfo",
"MacroRunner",
"MacroResult",
"ExcelOptimizer",
"ScreenOptimizer",
"CalculationOptimizer",
"OptimizationState",
```

### Tache 4.4 : Supprimer les aliases inutiles dans `__init__.py` (INIT-002)

**Fichier** : `src/xlmanage/__init__.py`

**Avant** :
```python
from .table_manager import TableInfo as TableInfoData
from .workbook_manager import WorkbookInfo as WorkbookInfoClass
from .worksheet_manager import WorksheetInfo as WorksheetInfoData

WorkbookInfo = WorkbookInfoClass
WorksheetInfo = WorksheetInfoData
TableInfo = TableInfoData
```

**Apres** :
```python
from .table_manager import TableInfo
from .workbook_manager import WorkbookInfo
from .worksheet_manager import WorksheetInfo
```

Import direct sans aliases confus.

### Tache 4.5 : Ajouter l'entete licence GPL a `__init__.py`

**Fichier** : `src/xlmanage/__init__.py`

Le fichier `__init__.py` n'a pas l'entete GPL v3 requise par CLAUDE.md. Ajouter :

```python
"""
xlmanage package initialization - Exports publics et version.

This file is part of xlManage.

xlManage is free software: ...
"""
```

### Tache 4.6 : Tests CLI pour `run-macro`

**Fichier** : `tests/test_cli_run_macro.py` (nouveau)

Suivre la specification de l'Epic 12 Story 3 Tache 3.4 pour les tests.

Minimum 7 tests :
- Execution reussie Sub (pas de retour)
- Execution reussie Function (avec retour)
- Erreur VBA runtime
- Avec workbook specifie
- Workbook introuvable
- Macro introuvable
- Affichage de l'aide (`--help`)

## Criteres d'acceptation

1. [ ] La commande `run-macro` est ajoutee au CLI avec 4 parametres (macro_name, --workbook, --args, --timeout)
2. [ ] Le resultat est affiche avec Rich (vert=succes, rouge=erreur)
3. [ ] `__init__.py` exporte les 8 elements manquants
4. [ ] Les aliases inutiles sont supprimes
5. [ ] L'entete GPL est presente dans `__init__.py`
6. [ ] Tous les tests CLI `run-macro` passent
7. [ ] 21/21 commandes sont implementees

## Dependances

- Epic 12 Stories 1 et 2 (MacroRunner, _parse_macro_args) - deja terminees
- Story 3 de cet epic (optimizers) pour les imports corrects dans `__init__.py`

## Definition of Done

- [ ] Commande `run-macro` fonctionnelle
- [ ] `__init__.py` complet et propre
- [ ] Tests CLI run-macro passent (>= 7 tests)
- [ ] `xlmanage run-macro --help` affiche l'aide complete
