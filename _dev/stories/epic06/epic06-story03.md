# Epic 6 - Story 3 : Implémenter la fonction utilitaire _find_open_workbook

## Vue d'ensemble

**En tant que** développeur
**Je veux** une fonction pour chercher un classeur déjà ouvert dans Excel
**Afin de** éviter d'ouvrir deux fois le même fichier

## Critères d'acceptation

1. ✅ Fonction `_find_open_workbook()` implémentée
2. ✅ Recherche d'abord par FullName (chemin complet)
3. ✅ Recherche ensuite par Name (nom de fichier) en fallback
4. ✅ Retourne None si non trouvé
5. ✅ Tests couvrent tous les scénarios

## Tâches techniques

### Tâche 3.1 : Implémenter _find_open_workbook

**Fichier** : `src/xlmanage/workbook_manager.py`

Ajouter cette fonction après `_detect_file_format()` :

```python
def _find_open_workbook(app: CDispatch, path: Path) -> CDispatch | None:
    """Find an open workbook by path.

    Searches for a workbook in the Excel instance by comparing paths.
    First tries to match by FullName (complete path), then falls back
    to matching by Name (filename only).

    Args:
        app: Excel Application COM object
        path: Path to the workbook to find

    Returns:
        Workbook COM object if found, None otherwise

    Note:
        The search is case-insensitive on Windows.
        Paths are resolved to absolute paths before comparison.

    Examples:
        >>> app = win32com.client.Dispatch("Excel.Application")
        >>> wb = _find_open_workbook(app, Path("C:/data/test.xlsx"))
        >>> if wb:
        ...     print(f"Found: {wb.Name}")
    """
    # Resolve to absolute path for comparison
    resolved_path = path.resolve()
    filename = path.name

    # Iterate through all open workbooks
    for wb in app.Workbooks:
        try:
            # Method 1: Compare by full path (most reliable)
            wb_full_path = Path(wb.FullName).resolve()
            if wb_full_path == resolved_path:
                return wb

            # Method 2: Compare by filename only (fallback)
            # This handles cases where the path might be different
            # but it's actually the same file (network paths, etc.)
            if wb.Name.lower() == filename.lower():
                return wb

        except Exception:
            # If we can't read wb.FullName or wb.Name, skip this workbook
            continue

    return None
```

**Points d'attention** :

1. **Ordre de recherche** : FullName d'abord, puis Name
   - FullName est plus fiable car c'est le chemin complet
   - Name est un fallback pour les cas où le chemin peut différer (réseau, liens symboliques)

2. **Resolution des chemins** : utiliser `path.resolve()` pour normaliser
   - Convertit les chemins relatifs en absolus
   - Résout les `..` et `.`
   - Normalise les séparateurs

3. **Comparaison case-insensitive** : Windows ne distingue pas majuscules/minuscules
   - `wb.Name.lower() == filename.lower()`

4. **Gestion d'erreur** : si `wb.FullName` ou `wb.Name` raise (classeur corrompu, etc.)
   - On continue avec le prochain classeur
   - On ne fait pas échouer toute la recherche

5. **Retour None** : si aucun classeur ne correspond
   - Permet à l'appelant de distinguer "non trouvé" de "erreur"

### Tâche 3.2 : Écrire les tests

**Fichier** : `tests/test_workbook_manager.py`

Ajouter cette classe de tests :

```python
class TestFindOpenWorkbook:
    """Tests for _find_open_workbook function."""

    def test_find_workbook_by_full_path(self):
        """Test finding workbook by complete path."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\test.xlsx"
        mock_wb.Name = "test.xlsx"
        mock_app.Workbooks = [mock_wb]

        # Test
        path = Path("C:/data/test.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Verify
        assert result == mock_wb

    def test_find_workbook_by_name(self):
        """Test finding workbook by filename when paths differ."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock with different path representation
        mock_app = Mock()
        mock_wb = Mock()
        # FullName uses different path format (UNC path, for example)
        mock_wb.FullName = "\\\\server\\share\\test.xlsx"
        mock_wb.Name = "test.xlsx"
        mock_app.Workbooks = [mock_wb]

        # Test with local path
        path = Path("C:/local/test.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Should find by name since full paths don't match
        assert result == mock_wb

    def test_find_workbook_not_found(self):
        """Test returning None when workbook is not found."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.FullName = "C:\\data\\other.xlsx"
        mock_wb.Name = "other.xlsx"
        mock_app.Workbooks = [mock_wb]

        # Test
        path = Path("C:/data/missing.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Verify
        assert result is None

    def test_find_workbook_case_insensitive(self):
        """Test that search is case-insensitive."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_wb = Mock()
        mock_wb.FullName = "C:\\DATA\\TEST.XLSX"
        mock_wb.Name = "TEST.XLSX"
        mock_app.Workbooks = [mock_wb]

        # Test with different case
        path = Path("c:/data/test.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Should find it
        assert result == mock_wb

    def test_find_workbook_multiple_workbooks(self):
        """Test finding correct workbook among multiple open workbooks."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mocks
        mock_app = Mock()

        mock_wb1 = Mock()
        mock_wb1.FullName = "C:\\data\\file1.xlsx"
        mock_wb1.Name = "file1.xlsx"

        mock_wb2 = Mock()
        mock_wb2.FullName = "C:\\data\\file2.xlsx"
        mock_wb2.Name = "file2.xlsx"

        mock_wb3 = Mock()
        mock_wb3.FullName = "C:\\data\\file3.xlsx"
        mock_wb3.Name = "file3.xlsx"

        mock_app.Workbooks = [mock_wb1, mock_wb2, mock_wb3]

        # Test finding the middle one
        path = Path("C:/data/file2.xlsx")
        result = _find_open_workbook(mock_app, path)

        assert result == mock_wb2

    def test_find_workbook_with_exception(self):
        """Test handling exception when accessing workbook properties."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock with one broken workbook
        mock_app = Mock()

        mock_wb1 = Mock()
        # This workbook raises exception when accessing FullName
        type(mock_wb1).FullName = Mock(side_effect=Exception("Access denied"))

        mock_wb2 = Mock()
        mock_wb2.FullName = "C:\\data\\good.xlsx"
        mock_wb2.Name = "good.xlsx"

        mock_app.Workbooks = [mock_wb1, mock_wb2]

        # Should skip the broken workbook and find the good one
        path = Path("C:/data/good.xlsx")
        result = _find_open_workbook(mock_app, path)

        assert result == mock_wb2

    def test_find_workbook_empty_collection(self):
        """Test with no open workbooks."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_app.Workbooks = []

        # Test
        path = Path("C:/data/any.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Should return None
        assert result is None

    def test_find_workbook_relative_path(self):
        """Test with relative path (should be resolved to absolute)."""
        from xlmanage.workbook_manager import _find_open_workbook

        # Setup mock
        mock_app = Mock()
        mock_wb = Mock()
        # Absolute path in Excel
        mock_wb.FullName = str(Path("data/test.xlsx").resolve())
        mock_wb.Name = "test.xlsx"
        mock_app.Workbooks = [mock_wb]

        # Test with relative path
        path = Path("data/test.xlsx")
        result = _find_open_workbook(mock_app, path)

        # Should resolve and find it
        assert result == mock_wb
```

**Commande de test** :
```bash
poetry run pytest tests/test_workbook_manager.py::TestFindOpenWorkbook -v
```

### Définition of Done

- [x] Fonction `_find_open_workbook()` implémentée
- [x] Recherche par FullName puis Name
- [x] Gestion des exceptions lors de l'itération
- [x] Tous les tests passent (minimum 8 tests)
- [x] Couverture de code 100%
- [x] Documentation complète avec exemples

**Statut** : ✅ TERMINÉ - Commit 711b5d7
**Rapport** : [_dev/reports/epic06-story03-implémentation.md](_dev/reports/epic06-story03-implémentation.md)
