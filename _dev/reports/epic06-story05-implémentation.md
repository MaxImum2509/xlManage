# Epic 6 - Story 5: Implémentation de WorkbookManager.create()

## Vue d'ensemble

Cette story implémente la méthode `create()` pour la classe `WorkbookManager`, permettant aux utilisateurs de créer de nouveaux classeurs Excel avec ou sans template. Cette fonctionnalité complète le cycle de vie des classeurs en permettant la création programmatique de fichiers Excel.

## Statut

**✅ TERMINÉ** - Tous les critères d'acceptation sont remplis

## Critères d'acceptation remplis

1. ✅ Méthode `create()` implémentée
2. ✅ Support création avec et sans template
3. ✅ Détection automatique du format (xlsx/xlsm/xls/xlsb)
4. ✅ Sauvegarde immédiate au bon format
5. ✅ Retourne WorkbookInfo
6. ✅ Tests couvrent templates, formats, erreurs

## Implémentation technique

### 1. Méthode create()

```python
def create(self, path: Path, template: Path | None = None) -> WorkbookInfo:
    """Create a new workbook.

    Creates a new Excel workbook and saves it to the specified path.
    Optionally uses a template file as starting point.

    Args:
        path: Destination path for the new workbook
        template: Optional path to a template file (.xltx, .xltm, or .xlsx/.xlsm)

    Returns:
        WorkbookInfo with details about the created workbook

    Raises:
        WorkbookNotFoundError: If template file doesn't exist
        WorkbookSaveError: If save operation fails
        ExcelConnectionError: If COM connection fails
    """
```

### 2. Séquence d'exécution

**Validation du template (si fourni)** :
```python
if template is not None:
    if not template.exists():
        raise WorkbookNotFoundError(
            template,
            f"Template file not found: {template}",
        )
```

**Détection du format de fichier** :
```python
try:
    file_format = _detect_file_format(path)
except ValueError as e:
    raise WorkbookSaveError(
        path,
        message=f"Invalid file extension: {str(e)}",
    ) from e
```

**Création du classeur** :
```python
if template is None:
    # Create blank workbook
    wb = app.Workbooks.Add()
else:
    # Create from template
    wb = app.Workbooks.Add(str(template.resolve()))
```

**Sauvegarde avec gestion d'erreur** :
```python
try:
    wb.SaveAs(abs_path, FileFormat=file_format)
except Exception as e:
    # Clean up the unsaved workbook
    try:
        wb.Close(SaveChanges=False)
        del wb
    except Exception:
        pass

    # Raise save error
    if hasattr(e, "hresult"):
        raise WorkbookSaveError(
            path,
            hresult=getattr(e, "hresult"),
            message=f"Failed to save workbook: {str(e)}",
        ) from e
    else:
        raise WorkbookSaveError(
            path,
            message=f"Failed to save workbook: {str(e)}",
        ) from e
```

**Construction et retour de WorkbookInfo** :
```python
info = WorkbookInfo(
    name=wb.Name,
    full_path=Path(wb.FullName),
    read_only=wb.ReadOnly,
    saved=wb.Saved,
    sheets_count=wb.Worksheets.Count,
)
return info
```

### 3. Gestion d'erreur complète

| Type d'erreur | Exception levée | Condition |
|--------------|----------------|-----------|
| Template introuvable | `WorkbookNotFoundError` | `not template.exists()` |
| Extension invalide | `WorkbookSaveError` | `_detect_file_format()` lève `ValueError` |
| Échec de sauvegarde | `WorkbookSaveError` | `wb.SaveAs()` échoue |
| Erreur COM générale | `ExcelConnectionError` | Autres erreurs COM |

### 4. Import nécessaire ajouté

```python
from .exceptions import (
    ExcelConnectionError,
    WorkbookAlreadyOpenError,
    WorkbookNotFoundError,
    WorkbookSaveError,  # Ajouté pour Story 5
)
```

## Tests implémentés

### Classe TestWorkbookManagerCreate

```python
class TestWorkbookManagerCreate:
    """Tests for WorkbookManager.create() method."""
```

**9 tests complets couvrant tous les scénarios :**

1. **test_create_blank_workbook** : Création d'un classeur vide
2. **test_create_with_template** : Création à partir d'un template
3. **test_create_xlsm_format** : Création de classeur macro-enabled (.xlsm)
4. **test_create_xls_legacy_format** : Création de format hérité (.xls)
5. **test_create_template_not_found** : Template introuvable
6. **test_create_invalid_extension** : Extension de fichier invalide
7. **test_create_save_fails** : Échec de sauvegarde avec cleanup
8. **test_create_com_error** : Erreur COM pendant la création
9. **test_create_cleanup_fails_silently** : Échec silencieux du cleanup

**Résultats des tests :**
- ✅ 9/9 tests passant (100% succès)
- ✅ Tous les scénarios couverts (succès et erreurs)
- ✅ Gestion d'erreur complète validée

## Intégration et compatibilité

### Dépendances existantes utilisées

- `_detect_file_format()` : Détection automatique du format (Story 2)
- `WorkbookInfo` : Structure de données pour les informations (Story 2)
- `WorkbookNotFoundError`, `WorkbookSaveError` : Exceptions (Story 1)
- `ExcelConnectionError` : Erreurs COM (Story 1)

### Fonctionnalités supportées

| Format | Code | Supporté |
|--------|------|----------|
| .xlsx | 51 | ✅ |
| .xlsm | 52 | ✅ |
| .xls | 56 | ✅ |
| .xlsb | 50 | ✅ |
| .xltx | 54 | ✅ (templates)|

## Couverture de code

```
Name                               Stmts   Miss  Cover   Missing
----------------------------------------------------------------
src\xlmanage\workbook_manager.py      83     30    64%    26-27, 121-142, 192-233, 341
```

**Analyse des lignes couvertes :**
- Méthode `create()` : 100% couverture
- Gestion d'erreur : Tous les chemins testés
- Cleanup : Vérifié dans les tests d'échec

## Définition of Done

- ✅ Méthode create() implémentée avec toutes les fonctionnalités
- ✅ Support template optionnel fonctionnel
- ✅ Détection automatique du format
- ✅ Cleanup en cas d'erreur validé
- ✅ Tous les tests passent (9 nouveaux tests)
- ✅ Intégration avec les composants existants
- ✅ Documentation complète avec exemples
- ✅ Gestion d'erreur complète et appropriée
- ✅ Suivi des conventions du projet (PEP 8, docstrings en anglais)

## Prochaines étapes

Cette implémentation prépare le terrain pour les stories suivantes de l'Epic 6 :

- **Story 6** : Implémenter `WorkbookManager.close()` et `save()`
- **Story 7** : Implémenter `WorkbookManager.list()` et intégration CLI

La méthode `create()` est maintenant prête pour une utilisation complète et peut être intégrée avec l'interface CLI dans la Story 7.

## Commandes de vérification

```bash
# Lancer les tests spécifiques
poetry run pytest tests/test_workbook_manager.py::TestWorkbookManagerCreate -v

# Vérifier la couverture
poetry run pytest tests/test_workbook_manager.py::TestWorkbookManagerCreate --cov=src/xlmanage/workbook_manager --cov-report=term-missing

# Linting
poetry run ruff check src/xlmanage/workbook_manager.py

# Type checking
poetry run mypy src/xlmanage/workbook_manager.py

# Test d'import
poetry run python -c "from xlmanage import WorkbookManager; print('Import successful')"
```

## Exemples d'utilisation

```python
# Créer un classeur vide
with ExcelManager() as excel_mgr:
    wb_mgr = WorkbookManager(excel_mgr)
    info = wb_mgr.create(Path("C:/data/new.xlsx"))
    print(f"Created {info.name} with {info.sheets_count} sheets")

# Créer à partir d'un template
with ExcelManager() as excel_mgr:
    wb_mgr = WorkbookManager(excel_mgr)
    template = Path("C:/templates/report.xltx")
    info = wb_mgr.create(Path("C:/data/report.xlsx"), template=template)
    print(f"Created from template: {info.name}")

# Création avec différents formats
formats = [
    Path("data.xlsx"),    # Format standard
    Path("macro.xlsm"),    # Macro-enabled
    Path("legacy.xls"),    # Excel 97-2003
    Path("binary.xlsb"),   # Format binaire
]

with ExcelManager() as excel_mgr:
    wb_mgr = WorkbookManager(excel_mgr)
    for path in formats:
        info = wb_mgr.create(path)
        print(f"Created {path.name}: {info.sheets_count} sheets")
```

## Conclusion

L'implémentation de la Story 5 fournit une fonctionnalité complète pour la création de classeurs Excel dans xlManage. La méthode `create()` suit les meilleures pratiques de développement, offre une gestion d'erreur complète avec cleanup approprié, et est prête pour les extensions futures. Les tests exhaustifs garantissent la fiabilité et la maintenabilité du code, tandis que l'intégration avec les composants existants assure la cohérence de l'API.

Cette implémentation représente une étape clé dans la gestion complète du cycle de vie des classeurs Excel, permettant aux utilisateurs de créer des fichiers Excel par programmation avec une API simple et robuste.
