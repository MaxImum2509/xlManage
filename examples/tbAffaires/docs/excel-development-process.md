# Processus de développement Excel (VBA) avec xlManage

## Objectif

Mettre en place une boucle de développement rapide et fiable pour l'application TBFR (VBA/Excel) en utilisant `xlManage` comme interface d'import/export et d'exécution (macros + orchestration), avec des tests VBA qui écrivent leurs résultats dans la feuille `test results`.

Ce document décrit une méthode de travail souple (adaptable au besoin) et puissante (automatisable, reproductible, traçable).

## Principes

1. **Le code source est dans Git**: vous écrivez les modules VBA directement dans `src/` (fichiers `.bas/.cls/.frm`). Le classeur `.xlsm` est une cible d'exécution, pas la source.
2. **Import idempotent**: l'import vers Excel doit être rejouable (overwrite) et sans actions manuelles dans l'IDE VBA.
3. **Tests comme contrat**: chaque changement VBA s'accompagne d'un ou plusieurs tests. Les tests s'exécutent via macro, et produisent un résultat machine-lisible.
4. **Résultats observables**: les tests écrivent dans `test results` (idéalement sous forme de Table/ListObject), puis un parseur détecte les échecs.
5. **RAII et non-régression Excel**: les tests/restaurations doivent laisser Excel dans un état propre (pas de boîtes de dialogue, pas de fichiers ouverts en erreur, pas d'instances zombies).

## Contraintes à respecter (rappels)

### VBA

- `.bas/.cls/.frm` en **Windows-1252 + CRLF**.
- **Chemins en `/`** dans les scripts/commandes.
- `Option Explicit` obligatoire, pas de `Select`/`Activate`.
- Conventions: Français + PascalCase, modules `modXXX.bas` / classes `clsYYY.cls`.

### Python / outillage

- Utiliser `poetry` (jamais `pip`) et exécuter via `poetry run`.
- Ne jamais utiliser `\\` comme séparateur de chemin (utiliser `/` ou `pathlib`).

## Boucle de travail (cycle court)

### (Optionnel) Découvrir xlManage (CLI + API)

Avant d'automatiser, identifiez les commandes disponibles et les signatures côté Python.

```bash
poetry run xlmanage --help
poetry run xlmanage vba --help
poetry run xlmanage run-macro --help

# Introspection API (méthodologie similaire à docs/python-documentation-discovery.md)
poetry run python -c "from xlmanage.vba_manager import VBAManager; help(VBAManager)"
poetry run python -c "from xlmanage.macro_runner import MacroRunner; help(MacroRunner)"
```

### 0) Préparation (une fois par poste)

1. Activer dans Excel: "Trust access to the VBA project object model" (sinon l'import VBA échoue).
2. Vérifier que le classeur cible est au format macro (`.xlsm`).
3. Installer l'environnement:

```bash
poetry install
poetry run xlmanage version
```

### 1) Écrire le code VBA dans `src/`

1. Créer/modifier les modules dans `src/`.
2. Respecter les en-têtes module/procédure (GPL + headers procs) et les patterns (ListObject, arrays bulk, RAII, gestion d'erreurs `On Error GoTo` + CleanUp + Log).
3. Garder les procédures petites et testables (SRP), et externaliser les effets de bord (I/O Excel, fichiers) derrière des fonctions claires.

### 2) Importer les modules dans le classeur (via xlManage)

Approche minimale (import fichier par fichier):

```bash
poetry run xlmanage start --new --visible
poetry run xlmanage workbook open tbAffaires.xlsm

poetry run xlmanage vba import src/modl_Utils.bas --workbook tbAffaires.xlsm --overwrite
poetry run xlmanage vba import src/clsApplicationState.cls --workbook tbAffaires.xlsm --overwrite
```

Approche robuste (batch):

- Écrire un petit script Python (dans `scripts/`) qui liste les fichiers `src/**/*.bas`, `src/**/*.cls`, `src/**/*.frm` et appelle `xlmanage.vba_manager.VBAManager.import_module(..., overwrite=True)`.
- Avantage: un seul point d'entrée, logs uniformes, et possibilité d'ajouter des règles (ordre d'import, exclusions, vérifications d'encodage).

### 3) Concevoir des tests VBA importables

Objectif: des tests exécutables via une macro unique (ex: `modTestsRunner.RunAllTests`) qui écrit des lignes dans `test results`.

Recommandations de conception:

- Mettre les tests dans des modules dédiés (ex: `src/modTestsRunner.bas`, `src/modTests_*.bas`).
- Un test = une `Sub` sans paramètre (ex: `Public Sub Test_CalculTotal_Nominal()`).
- Ne pas dépendre de l'ordre d'exécution (tests indépendants).
- Nettoyer ce qui est créé (feuilles, tables, noms, fichiers) dans un bloc `CleanUp`.
- Envelopper chaque test avec l'équivalent d'un "application state" (RAII) pour restaurer Calculation/ScreenUpdating/DisplayAlerts.

### 4) Standardiser la feuille `test results`

Nom de feuille: `test results`.

Format recommandé: une Table/ListObject (ex: `tbTestResults`) avec au minimum les colonnes:

- `RunId` (string): identifiant d'exécution (timestamp ou GUID)
- `TestName` (string)
- `Status` (string): `PASS` / `FAIL` / `ERROR` / `SKIP`
- `Message` (string): détail lisible
- `DurationMs` (long)
- `Timestamp` (date)

Bonnes pratiques:

- Le runner efface/archive les anciennes lignes au début d'un run (au choix: clear ou append avec `RunId`).
- Les asserts n'affichent pas de MsgBox: ils écrivent dans `test results` et laissent le runner continuer (sauf si erreur fatale).
- En cas d'exception VBA non gérée, le runner doit capturer l'erreur (`Err.Number`, `Err.Description`) et écrire `ERROR`.

### 5) Importer les tests et exécuter le runner

```bash
poetry run xlmanage vba import src/modTestsRunner.bas --workbook tbAffaires.xlsm --overwrite
poetry run xlmanage vba import src/modTests_Foo.bas --workbook tbAffaires.xlsm --overwrite

poetry run xlmanage run-macro "modTestsRunner.RunAllTests" --workbook tbAffaires.xlsm --timeout 120
```

### 6) Parser `test results` pour détecter les dysfonctionnements

Deux stratégies compatibles:

1. **Parsing côté VBA**: une macro `modTestsRunner.ExportResultsCsv(outputPath)` exporte la table `tbTestResults` vers un CSV dans `_dev/`, puis un script Python lit le CSV.
2. **Parsing côté Python via COM**: un script Python utilise `xlManage` pour démarrer Excel, ouvrir le classeur, puis lit `Worksheets("test results")` et la table/range par COM.

Exemple de parsing côté Python (à adapter):

```python
from __future__ import annotations

from pathlib import Path

from xlmanage.excel_manager import ExcelManager
from xlmanage.workbook_manager import WorkbookManager


def read_test_results(workbook_path: Path) -> list[dict[str, object]]:
    with ExcelManager(visible=False) as mgr:
        mgr.start(new=True)
        WorkbookManager(mgr).open(workbook_path, read_only=True)

        wb = mgr.app.Workbooks(workbook_path.name)
        ws = wb.Worksheets("test results")

        # Recommandé: résultats dans une Table "tbTestResults"
        lo = ws.ListObjects("tbTestResults")
        values = lo.DataBodyRange.Value  # 2D tuple
        headers = [c.Value for c in lo.HeaderRowRange.Value[0]]

        rows: list[dict[str, object]] = []
        for row in values:
            rows.append(dict(zip(headers, row, strict=False)))
        return rows


def main() -> int:
    rows = read_test_results(Path("tbAffaires.xlsm"))
    failures = [r for r in rows if str(r.get("Status", "")) not in {"PASS"}]
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
```

L'objectif est d'avoir un signal binaire exploitable (0 = OK, 1 = KO) et une sortie lisible (liste des tests en échec) pour accélérer la correction.

## Stratégie de développement (souple, par niveaux)

### Niveau 1: itération locale rapide

- Import uniquement les modules modifiés.
- Lancer un sous-ensemble de tests (suite cible) via une macro `RunSuite("SuiteName")`.
- Parser uniquement les lignes du dernier `RunId`.

### Niveau 2: régression régulière

- Batch import complet (tous modules `src/`).
- `RunAllTests`.
- Export/parse (CSV ou COM) et blocage si échec.

### Niveau 3: stabilisation avant merge

- Lancer la batterie complète + scénarios manuels si nécessaire.
- Vérifier que le classeur s'ouvre sans dialogues, et qu'aucune instance Excel zombie ne reste active.

## Checklists

### Avant de lancer les tests

- Le classeur cible est `.xlsm` et accessible.
- L'accès VBProject est autorisé (Trust Center).
- Aucun `MsgBox`/dialogue bloquant n'est attendu dans le code testé.

### Avant commit

- Les tests écrivent bien dans `test results`.
- Le parseur détecte correctement les FAIL/ERROR.
- Message de commit conforme Conventional Commits (ex: `feat(vba): ...`, `test(vba): ...`, `docs: ...`).

## Dépannage rapide

- Import VBA échoue: vérifier Trust Center + encodage Windows-1252.
- Macro introuvable: vérifier nom complet (`Module.Procedure`) et que le module a bien été importé dans le bon classeur.
- Dialogues Excel: s'assurer que le runner désactive/restore `DisplayAlerts` et que les opérations (delete sheet/table) sont faites sans confirmation.
- Instances Excel zombies: utiliser `poetry run xlmanage status` puis `poetry run xlmanage stop` (en dernier recours seulement, `force_kill` existe côté API).
