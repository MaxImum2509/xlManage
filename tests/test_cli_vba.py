"""
Tests for VBA CLI commands.

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

from pathlib import Path
from unittest.mock import Mock, patch

import pytest
from typer.testing import CliRunner

from xlmanage.cli import app
from xlmanage.exceptions import (
    VBAExportError,
    VBAImportError,
    VBAModuleAlreadyExistsError,
    VBAModuleNotFoundError,
    VBAProjectAccessError,
    VBAWorkbookFormatError,
)
from xlmanage.vba_manager import VBAModuleInfo

runner = CliRunner()


class TestVBAImport:
    """Tests for vba import command."""

    def test_vba_import_success(self, tmp_path):
        """Test vba import command success."""
        bas_file = tmp_path / "Module1.bas"
        bas_file.write_text("Sub Test()\nEnd Sub", encoding="windows-1252")

        mock_info = VBAModuleInfo(
            name="Module1", module_type="standard", lines_count=2, has_predeclared_id=False
        )

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.import_module.return_value = mock_info
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "import", str(bas_file)])

            assert result.exit_code == 0
            assert "importé avec succès" in result.stdout
            assert "Module1" in result.stdout
            mock_vba.import_module.assert_called_once()

    def test_vba_import_with_options(self, tmp_path):
        """Test vba import with all options."""
        cls_file = tmp_path / "MyClass.cls"
        cls_file.write_text("", encoding="windows-1252")
        workbook = tmp_path / "test.xlsm"

        mock_info = VBAModuleInfo(
            name="MyClass", module_type="class", lines_count=10, has_predeclared_id=True
        )

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.import_module.return_value = mock_info
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(
                app,
                [
                    "vba",
                    "import",
                    str(cls_file),
                    "--type",
                    "class",
                    "--workbook",
                    str(workbook),
                    "--overwrite",
                ],
            )

            assert result.exit_code == 0
            assert "MyClass" in result.stdout
            mock_vba.import_module.assert_called_once_with(
                module_file=Path(str(cls_file)),
                module_type="class",
                workbook=Path(str(workbook)),
                overwrite=True,
            )

    def test_vba_import_project_access_error(self, tmp_path):
        """Test vba import with VBAProjectAccessError."""
        bas_file = tmp_path / "Module1.bas"
        bas_file.write_text("", encoding="windows-1252")

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.import_module.side_effect = VBAProjectAccessError("test.xlsm")
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "import", str(bas_file)])

            assert result.exit_code == 1
            assert "Erreur d'accès VBA" in result.stdout
            assert "Trust access" in result.stdout

    def test_vba_import_workbook_format_error(self, tmp_path):
        """Test vba import with VBAWorkbookFormatError."""
        bas_file = tmp_path / "Module1.bas"
        bas_file.write_text("", encoding="windows-1252")

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.import_module.side_effect = VBAWorkbookFormatError("test.xlsx")
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "import", str(bas_file)])

            assert result.exit_code == 1
            assert "Format de classeur invalide" in result.stdout
            assert ".xlsm" in result.stdout

    def test_vba_import_module_exists_error(self, tmp_path):
        """Test vba import with VBAModuleAlreadyExistsError."""
        bas_file = tmp_path / "Module1.bas"
        bas_file.write_text("", encoding="windows-1252")

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.import_module.side_effect = VBAModuleAlreadyExistsError(
                "Module1", "test.xlsm"
            )
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "import", str(bas_file)])

            assert result.exit_code == 1
            assert "Module existant" in result.stdout
            assert "Module1" in result.stdout
            assert "--overwrite" in result.stdout

    def test_vba_import_generic_error(self, tmp_path):
        """Test vba import with VBAImportError."""
        bas_file = tmp_path / "Module1.bas"
        bas_file.write_text("", encoding="windows-1252")

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.import_module.side_effect = VBAImportError(
                "Module1.bas", "Invalid encoding"
            )
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "import", str(bas_file)])

            assert result.exit_code == 1
            assert "Erreur d'import" in result.stdout


class TestVBAExport:
    """Tests for vba export command."""

    def test_vba_export_success(self, tmp_path):
        """Test vba export command success."""
        output_file = tmp_path / "Module1.bas"

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.export_module.return_value = output_file
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "export", "Module1", str(output_file)])

            assert result.exit_code == 0
            assert "exporté avec succès" in result.stdout
            assert "Module1" in result.stdout
            mock_vba.export_module.assert_called_once()

    def test_vba_export_with_workbook(self, tmp_path):
        """Test vba export with workbook option."""
        output_file = tmp_path / "ThisWorkbook.cls"
        workbook = tmp_path / "test.xlsm"

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.export_module.return_value = output_file
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(
                app,
                ["vba", "export", "ThisWorkbook", str(output_file), "--workbook", str(workbook)],
            )

            assert result.exit_code == 0
            mock_vba.export_module.assert_called_once_with(
                module_name="ThisWorkbook",
                output_file=Path(str(output_file)),
                workbook=Path(str(workbook)),
            )

    def test_vba_export_module_not_found(self, tmp_path):
        """Test vba export with module not found."""
        output_file = tmp_path / "Module1.bas"

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.export_module.side_effect = VBAModuleNotFoundError(
                "Module1", "test.xlsm"
            )
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "export", "Module1", str(output_file)])

            assert result.exit_code == 1
            assert "Module introuvable" in result.stdout

    def test_vba_export_error(self, tmp_path):
        """Test vba export with VBAExportError."""
        output_file = tmp_path / "Module1.bas"

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.export_module.side_effect = VBAExportError(
                "Module1", str(output_file), "Permission denied"
            )
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "export", "Module1", str(output_file)])

            assert result.exit_code == 1
            assert "Erreur d'export" in result.stdout


class TestVBAList:
    """Tests for vba list command."""

    def test_vba_list_success(self):
        """Test vba list command success."""
        mock_modules = [
            VBAModuleInfo("Module1", "standard", 42, False),
            VBAModuleInfo("MyClass", "class", 15, True),
            VBAModuleInfo("ThisWorkbook", "document", 8, False),
        ]

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.list_modules.return_value = mock_modules
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "list"])

            assert result.exit_code == 0
            assert "Module1" in result.stdout
            assert "MyClass" in result.stdout
            assert "ThisWorkbook" in result.stdout
            assert "Total : 3 module(s)" in result.stdout

    def test_vba_list_empty(self):
        """Test vba list with no modules."""
        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.list_modules.return_value = []
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "list"])

            assert result.exit_code == 0
            assert "Aucun module VBA trouvé" in result.stdout

    def test_vba_list_with_workbook(self, tmp_path):
        """Test vba list with workbook option."""
        workbook = tmp_path / "test.xlsm"
        mock_modules = [VBAModuleInfo("Module1", "standard", 10, False)]

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.list_modules.return_value = mock_modules
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "list", "--workbook", str(workbook)])

            assert result.exit_code == 0
            assert "test.xlsm" in result.stdout
            mock_vba.list_modules.assert_called_once_with(workbook=Path(str(workbook)))

    def test_vba_list_project_access_error(self):
        """Test vba list with VBAProjectAccessError."""
        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.list_modules.side_effect = VBAProjectAccessError("test.xlsm")
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "list"])

            assert result.exit_code == 1
            assert "Erreur d'accès VBA" in result.stdout
            assert "Trust access" in result.stdout


class TestVBADelete:
    """Tests for vba delete command."""

    def test_vba_delete_success(self):
        """Test vba delete command success."""
        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.delete_module.return_value = None
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "delete", "Module1"])

            assert result.exit_code == 0
            assert "supprimé avec succès" in result.stdout
            assert "Module1" in result.stdout
            mock_vba.delete_module.assert_called_once()

    def test_vba_delete_with_options(self, tmp_path):
        """Test vba delete with all options."""
        workbook = tmp_path / "test.xlsm"

        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.delete_module.return_value = None
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(
                app, ["vba", "delete", "Module1", "--workbook", str(workbook), "--force"]
            )

            assert result.exit_code == 0
            mock_vba.delete_module.assert_called_once_with(
                module_name="Module1", workbook=Path(str(workbook)), force=True
            )

    def test_vba_delete_module_not_found(self):
        """Test vba delete with module not found."""
        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.delete_module.side_effect = VBAModuleNotFoundError(
                "Module1", "test.xlsm"
            )
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "delete", "Module1"])

            assert result.exit_code == 1
            assert "Erreur" in result.stdout

    def test_vba_delete_document_module(self):
        """Test vba delete with document module (not deletable)."""
        with patch("xlmanage.cli.ExcelManager") as mock_mgr_class, patch(
            "xlmanage.cli.VBAManager"
        ) as mock_vba_class:
            mock_mgr = Mock()
            mock_mgr_class.return_value.__enter__.return_value = mock_mgr

            mock_vba = Mock()
            mock_vba.delete_module.side_effect = VBAModuleNotFoundError(
                "ThisWorkbook", "test.xlsm", reason="Cannot delete document module"
            )
            mock_vba_class.return_value = mock_vba

            result = runner.invoke(app, ["vba", "delete", "ThisWorkbook"])

            assert result.exit_code == 1
            assert "Erreur" in result.stdout
            assert "modules de document" in result.stdout
