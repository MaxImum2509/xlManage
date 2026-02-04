# Epic 6 - Story 2 : Implémenter la dataclass WorkbookInfo et le mapping FileFormat

## Vue d'ensemble

**En tant que** développeur
**Je veux** une structure de données pour représenter les informations d'un classeur
**Afin de** retourner des informations typées aux utilisateurs de l'API

## Critères d'acceptation

1. ✅ WorkbookInfo dataclass créée avec 5 champs
2. ✅ FILE_FORMAT_MAP constant défini avec les 4 formats supportés
3. ✅ Fonction `_detect_file_format()` implémentée
4. ✅ Tests unitaires couvrent tous les formats et les erreurs

## Tâches techniques

### Tâche 2.1 : Créer le fichier workbook_manager.py

**Fichier** : `src/xlmanage/workbook_manager.py`

Commencer par les imports et la structure de base :

```python
"""
Workbook lifecycle management for xlmanage.

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

from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    from win32com.client import CDispatch
except ImportError:
    CDispatch = Any

from .excel_manager import ExcelManager
from .exceptions import (
    WorkbookNotFoundError,
    WorkbookAlreadyOpenError,
    WorkbookSaveError,
)
```

### Tâche 2.2 : Définir la constante FILE_FORMAT_MAP

```python
# Excel file format constants
# See: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat
FILE_FORMAT_MAP: dict[str, int] = {
    ".xlsx": 51,   # xlOpenXMLWorkbook
    ".xlsm": 52,   # xlOpenXMLWorkbookMacroEnabled
    ".xls": 56,    # xlExcel8 (Excel 97-2003 format)
    ".xlsb": 50,   # xlExcel12 (Excel binary workbook)
}
```

**Points d'attention** :
- Les clés sont des extensions avec le point (`.xlsx`, pas `xlsx`)
- Les valeurs sont les codes numériques utilisés par Excel COM API
- Ces codes sont documentés par Microsoft et ne doivent jamais changer

### Tâche 2.3 : Créer la dataclass WorkbookInfo

```python
@dataclass
class WorkbookInfo:
    """Information about an Excel workbook.

    Attributes:
        name: Workbook filename (e.g., "data.xlsx")
        full_path: Full path to the workbook file
        read_only: Whether the workbook is opened in read-only mode
        saved: Whether all changes have been saved
        sheets_count: Number of worksheets in the workbook
    """

    name: str
    full_path: Path
    read_only: bool
    saved: bool
    sheets_count: int
```

**Points d'attention** :
- `name` : nom du fichier uniquement (ex: "data.xlsx")
- `full_path` : chemin complet de type `Path`, pas `str`
- `read_only` : `True` si ouvert en lecture seule
- `saved` : `True` si toutes les modifications sont sauvegardées
- `sheets_count` : nombre de feuilles dans le classeur

### Tâche 2.4 : Implémenter _detect_file_format()

```python
def _detect_file_format(path: Path) -> int:
    """Detect Excel file format from file extension.

    Args:
        path: Path to the Excel file

    Returns:
        Excel FileFormat code (51, 52, 56, or 50)

    Raises:
        ValueError: If the file extension is not recognized

    Examples:
        >>> _detect_file_format(Path("data.xlsx"))
        51
        >>> _detect_file_format(Path("macro.xlsm"))
        52
    """
    extension = path.suffix.lower()

    if extension not in FILE_FORMAT_MAP:
        supported = ", ".join(FILE_FORMAT_MAP.keys())
        raise ValueError(
            f"Unsupported file extension '{extension}'. "
            f"Supported formats: {supported}"
        )

    return FILE_FORMAT_MAP[extension]
```

**Points d'attention** :
- Utiliser `path.suffix` pour extraire l'extension
- Convertir en minuscules avec `.lower()` pour gérer `.XLSX`, `.XlSx`, etc.
- Le message d'erreur liste tous les formats supportés pour guider l'utilisateur
- Cette fonction est privée (préfixe `_`) car elle n'est utilisée qu'en interne

### Tâche 2.5 : Écrire les tests

**Fichier** : `tests/test_workbook_manager.py` (nouveau fichier)

```python
"""
Tests for WorkbookManager functionality.
"""

import pytest
from pathlib import Path
from unittest.mock import Mock

from xlmanage.workbook_manager import (
    WorkbookInfo,
    FILE_FORMAT_MAP,
    _detect_file_format,
)


class TestWorkbookInfo:
    """Tests for WorkbookInfo dataclass."""

    def test_workbook_info_creation(self):
        """Test creating WorkbookInfo instance."""
        info = WorkbookInfo(
            name="test.xlsx",
            full_path=Path("C:/data/test.xlsx"),
            read_only=False,
            saved=True,
            sheets_count=3,
        )

        assert info.name == "test.xlsx"
        assert info.full_path == Path("C:/data/test.xlsx")
        assert info.read_only is False
        assert info.saved is True
        assert info.sheets_count == 3

    def test_workbook_info_fields(self):
        """Test all fields are accessible."""
        info = WorkbookInfo(
            name="data.xlsm",
            full_path=Path("D:/projects/data.xlsm"),
            read_only=True,
            saved=False,
            sheets_count=5,
        )

        # Verify all fields
        assert isinstance(info.name, str)
        assert isinstance(info.full_path, Path)
        assert isinstance(info.read_only, bool)
        assert isinstance(info.saved, bool)
        assert isinstance(info.sheets_count, int)


class TestFileFormatMap:
    """Tests for FILE_FORMAT_MAP constant."""

    def test_file_format_map_keys(self):
        """Test FILE_FORMAT_MAP has all expected extensions."""
        expected_extensions = {".xlsx", ".xlsm", ".xls", ".xlsb"}
        assert set(FILE_FORMAT_MAP.keys()) == expected_extensions

    def test_file_format_map_values(self):
        """Test FILE_FORMAT_MAP values are correct."""
        assert FILE_FORMAT_MAP[".xlsx"] == 51
        assert FILE_FORMAT_MAP[".xlsm"] == 52
        assert FILE_FORMAT_MAP[".xls"] == 56
        assert FILE_FORMAT_MAP[".xlsb"] == 50


class TestDetectFileFormat:
    """Tests for _detect_file_format function."""

    def test_detect_xlsx_format(self):
        """Test detecting .xlsx format."""
        path = Path("C:/data/file.xlsx")
        assert _detect_file_format(path) == 51

    def test_detect_xlsm_format(self):
        """Test detecting .xlsm format."""
        path = Path("D:/projects/macro.xlsm")
        assert _detect_file_format(path) == 52

    def test_detect_xls_format(self):
        """Test detecting .xls format."""
        path = Path("E:/legacy/old.xls")
        assert _detect_file_format(path) == 56

    def test_detect_xlsb_format(self):
        """Test detecting .xlsb format."""
        path = Path("F:/binary/data.xlsb")
        assert _detect_file_format(path) == 50

    def test_detect_format_case_insensitive(self):
        """Test format detection is case-insensitive."""
        assert _detect_file_format(Path("test.XLSX")) == 51
        assert _detect_file_format(Path("test.XlSm")) == 52
        assert _detect_file_format(Path("test.XLS")) == 56
        assert _detect_file_format(Path("test.XLSB")) == 50

    def test_detect_format_unsupported_extension(self):
        """Test ValueError is raised for unsupported extensions."""
        with pytest.raises(ValueError) as exc_info:
            _detect_file_format(Path("document.docx"))

        assert "Unsupported file extension" in str(exc_info.value)
        assert ".docx" in str(exc_info.value)
        assert ".xlsx" in str(exc_info.value)  # Lists supported formats

    def test_detect_format_no_extension(self):
        """Test ValueError for files without extension."""
        with pytest.raises(ValueError) as exc_info:
            _detect_file_format(Path("file_without_extension"))

        assert "Unsupported file extension" in str(exc_info.value)

    def test_detect_format_wrong_extension(self):
        """Test ValueError for wrong extensions."""
        invalid_files = [
            Path("data.csv"),
            Path("data.txt"),
            Path("data.pdf"),
            Path("data.xml"),
        ]

        for path in invalid_files:
            with pytest.raises(ValueError):
                _detect_file_format(path)
```

**Commande de test** :
```bash
poetry run pytest tests/test_workbook_manager.py::TestWorkbookInfo -v
poetry run pytest tests/test_workbook_manager.py::TestFileFormatMap -v
poetry run pytest tests/test_workbook_manager.py::TestDetectFileFormat -v
```

### Définition of Done

- [ ] WorkbookInfo dataclass créée avec 5 champs
- [ ] FILE_FORMAT_MAP défini avec les 4 formats
- [ ] `_detect_file_format()` implémentée avec gestion d'erreur
- [ ] Tous les tests passent (minimum 15 tests)
- [ ] Couverture de code 100% pour la dataclass et la fonction
- [ ] Documentation complète (docstrings avec exemples)
