"""
Tests for CLI commands.

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

from typer.testing import CliRunner

from xlmanage.cli import app
from xlmanage.excel_manager import InstanceInfo
from xlmanage.exceptions import ExcelConnectionError, ExcelManageError
from xlmanage.table_manager import TableInfo
from xlmanage.workbook_manager import WorkbookInfo
from xlmanage.worksheet_manager import WorksheetInfo

runner = CliRunner()


class TestVersionCommand:
    """Test version command."""

    def test_version_command(self):
        """Test version command output."""
        result = runner.invoke(app, ["version"])
        assert result.exit_code == 0
        assert "xlmanage" in result.stdout
        assert "0.1.0" in result.stdout


class TestStartCommand:
    """Test start command."""

    @patch("xlmanage.cli.ExcelManager")
    def test_start_command_default(self, mock_manager_class):
        """Test start command with default options."""
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=False, workbooks_count=0, hwnd=5678)
        mock_manager.start.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["start"])

        # Assertions
        assert result.exit_code == 0
        mock_manager_class.assert_called_once_with(visible=False)
        mock_manager.start.assert_called_once_with(new=False)
        assert "Excel instance started successfully" in result.stdout
        assert "1234" in result.stdout
        assert "5678" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_start_command_visible(self, mock_manager_class):
        """Test start command with --visible option."""
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=0, hwnd=5678)
        mock_manager.start.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["start", "--visible"])

        # Assertions
        assert result.exit_code == 0
        mock_manager_class.assert_called_once_with(visible=True)
        mock_manager.start.assert_called_once_with(new=False)
        assert "visible" in result.stdout.lower()

    @patch("xlmanage.cli.ExcelManager")
    def test_start_command_new(self, mock_manager_class):
        """Test start command with --new option."""
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=False, workbooks_count=0, hwnd=5678)
        mock_manager.start.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["start", "--new"])

        # Assertions
        assert result.exit_code == 0
        mock_manager.start.assert_called_once_with(new=True)
        assert "new" in result.stdout.lower()

    @patch("xlmanage.cli.ExcelManager")
    def test_start_command_visible_and_new(self, mock_manager_class):
        """Test start command with both --visible and --new options."""
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=0, hwnd=5678)
        mock_manager.start.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["start", "--visible", "--new"])

        # Assertions
        assert result.exit_code == 0
        mock_manager_class.assert_called_once_with(visible=True)
        mock_manager.start.assert_called_once_with(new=True)

    @patch("xlmanage.cli.ExcelManager")
    def test_start_command_connection_error(self, mock_manager_class):
        """Test start command with connection error."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.start.side_effect = ExcelConnectionError(
            0x80080005, "Excel not installed"
        )
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["start"])

        # Assertions
        assert result.exit_code == 1
        assert "Connection Error" in result.stdout
        assert "Excel not installed" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_start_command_manage_error(self, mock_manager_class):
        """Test start command with ExcelManageError."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.start.side_effect = ExcelManageError("Management error")
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["start"])

        # Assertions
        assert result.exit_code == 1
        assert "Excel management error" in result.stdout
        assert "Management error" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_start_command_generic_error(self, mock_manager_class):
        """Test start command with generic error."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.start.side_effect = Exception("Unexpected error")
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["start"])

        # Assertions
        assert result.exit_code == 1
        assert "Unexpected Error" in result.stdout
        assert "Unexpected error" in result.stdout


class TestStopCommand:
    """Test stop command."""

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_default(self, mock_manager_class):
        """Test stop command with default options (no --force)."""
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)
        mock_manager.get_running_instance.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        # Run command without --force (confirm with y)
        result = runner.invoke(app, ["stop"], input="y\n")

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop_instance.assert_called_once_with(1234, save=True)
        assert "arrêtée avec succès" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_no_save(self, mock_manager_class):
        """Test stop command with --no-save option."""
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)
        mock_manager.get_running_instance.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        # Run command without --force, with --no-save (confirm with y)
        result = runner.invoke(app, ["stop", "--no-save"], input="y\n")

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop_instance.assert_called_once_with(1234, save=False)

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_with_confirmation_yes(self, mock_manager_class):
        """Test stop command with user confirmation (yes)."""
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)
        mock_manager.get_running_instance.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        # Run command with confirmation input
        result = runner.invoke(app, ["stop"], input="y\n")

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop_instance.assert_called_once_with(1234, save=True)

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_no_active_instance(self, mock_manager_class):
        """Test stop command when no instance is active."""
        # Setup mock - get_running_instance returns None
        mock_manager = Mock()
        mock_manager.get_running_instance.return_value = None
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["stop"])

        # Assertions
        assert result.exit_code == 0
        assert "Aucune instance" in result.stdout
        mock_manager.stop_instance.assert_not_called()

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_all_no_instances(self, mock_manager_class):
        """Test stop command with --all when no instances are running."""
        # Setup mock
        mock_manager = Mock()
        mock_manager.list_running_instances.return_value = []
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["stop", "--all"])

        # Assertions
        assert result.exit_code == 0
        assert "Aucune instance" in result.stdout
        mock_manager.stop_all.assert_not_called()

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_all_with_instances(self, mock_manager_class):
        """Test stop command with --all when instances are running."""
        # Setup mock
        mock_manager = Mock()
        instances = [
            InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678),
            InstanceInfo(pid=5678, visible=False, workbooks_count=0, hwnd=9012),
        ]
        mock_manager.list_running_instances.return_value = instances
        mock_manager.stop_all.return_value = [1234, 5678]
        mock_manager_class.return_value = mock_manager

        # Run command with --all (no --force = clean shutdown via stop_all)
        result = runner.invoke(app, ["stop", "--all"])

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop_all.assert_called_once_with(save=True)
        assert "arrêtée" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_all_with_no_save(self, mock_manager_class):
        """Test stop command with --all --no-save."""
        # Setup mock
        mock_manager = Mock()
        instances = [InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)]
        mock_manager.list_running_instances.return_value = instances
        mock_manager.stop_all.return_value = [1234]
        mock_manager_class.return_value = mock_manager

        # Run command with --all --no-save
        result = runner.invoke(app, ["stop", "--all", "--no-save"])

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop_all.assert_called_once_with(save=False)

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_all_with_partial_failure(self, mock_manager_class):
        """Test stop command with --all when some instances fail to stop."""
        # Setup mock
        mock_manager = Mock()
        instances = [
            InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678),
            InstanceInfo(pid=5678, visible=False, workbooks_count=0, hwnd=9012),
        ]
        mock_manager.list_running_instances.return_value = instances
        # stop_all returns only first PID (second failed)
        mock_manager.stop_all.return_value = [1234]
        mock_manager_class.return_value = mock_manager

        # Run command with --all (clean shutdown)
        result = runner.invoke(app, ["stop", "--all"])

        # Assertions
        assert result.exit_code == 0
        assert "1234" in result.stdout
        assert "échec" in result.stdout.lower() or "5678" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_connection_error(self, mock_manager_class):
        """Test stop command with connection error."""
        # Setup mock to raise exception on get_running_instance
        mock_manager = Mock()
        mock_manager.get_running_instance.side_effect = ExcelConnectionError(
            0x80080005, "Connection failed"
        )
        mock_manager_class.return_value = mock_manager

        # Run command without --force (calls _stop_active_instance)
        result = runner.invoke(app, ["stop"])

        # Assertions - exception propagates to stop() handler
        assert result.exit_code == 1

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_manage_error(self, mock_manager_class):
        """Test stop command with ExcelManageError."""
        # Setup mock to raise on stop_instance
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)
        mock_manager.get_running_instance.return_value = mock_info
        mock_manager.stop_instance.side_effect = ExcelManageError("Management error")
        mock_manager_class.return_value = mock_manager

        # Run command without --force (calls _stop_active_instance)
        result = runner.invoke(app, ["stop"])

        # Assertions - exception propagates to stop() handler
        assert result.exit_code == 1

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_generic_error(self, mock_manager_class):
        """Test stop command with generic error."""
        # Setup mock to raise on stop_instance
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)
        mock_manager.get_running_instance.return_value = mock_info
        mock_manager.stop_instance.side_effect = Exception("Unexpected error")
        mock_manager_class.return_value = mock_manager

        # Run command without --force
        result = runner.invoke(app, ["stop"])

        # Assertions - exception propagates to stop() generic handler
        assert result.exit_code == 1
        assert "Erreur" in result.stdout


class TestStatusCommand:
    """Test status command."""

    @patch("xlmanage.cli.ExcelManager")
    def test_status_command_no_instances(self, mock_manager_class):
        """Test status command when no instances are running."""
        # Setup mock
        mock_manager = Mock()
        mock_manager.list_running_instances.return_value = []
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["status"])

        # Assertions
        assert result.exit_code == 0
        assert "No running Excel instances found" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_status_command_with_instances(self, mock_manager_class):
        """Test status command when instances are running."""
        # Setup mock
        mock_manager = Mock()
        instances = [
            InstanceInfo(pid=1234, visible=True, workbooks_count=2, hwnd=5678),
            InstanceInfo(pid=5678, visible=False, workbooks_count=0, hwnd=9012),
        ]
        mock_manager.list_running_instances.return_value = instances
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["status"])

        # Assertions
        assert result.exit_code == 0
        assert "Running Excel Instances" in result.stdout
        assert "1234" in result.stdout
        assert "5678" in result.stdout
        assert "9012" in result.stdout
        assert "2 found" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_status_command_connection_error(self, mock_manager_class):
        """Test status command with connection error."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.list_running_instances.side_effect = ExcelConnectionError(
            0x80080005, "Connection failed"
        )
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["status"])

        # Assertions
        assert result.exit_code == 1
        assert "Connection Error" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_status_command_manage_error(self, mock_manager_class):
        """Test status command with ExcelManageError."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.list_running_instances.side_effect = ExcelManageError(
            "Management error"
        )
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["status"])

        # Assertions
        assert result.exit_code == 1
        assert "Excel management error" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_status_command_generic_error(self, mock_manager_class):
        """Test status command with generic error."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.list_running_instances.side_effect = Exception("Unexpected error")
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["status"])

        # Assertions
        assert result.exit_code == 1
        assert "Unexpected Error" in result.stdout


class TestCLIIntegration:
    """Test CLI integration scenarios."""

    @patch("xlmanage.cli.ExcelManager")
    def test_start_and_status_workflow(self, mock_manager_class):
        """Test workflow: start instance then check status."""
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=False, workbooks_count=0, hwnd=5678)
        mock_manager.start.return_value = mock_info
        mock_manager.list_running_instances.return_value = [mock_info]
        mock_manager_class.return_value = mock_manager

        # Start instance
        result1 = runner.invoke(app, ["start"])
        assert result1.exit_code == 0

        # Check status
        result2 = runner.invoke(app, ["status"])
        assert result2.exit_code == 0
        assert "1234" in result2.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_start_and_stop_workflow(self, mock_manager_class):
        """Test workflow: start instance then stop it."""
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=False, workbooks_count=0, hwnd=5678)
        mock_manager.start.return_value = mock_info
        mock_manager.get_running_instance.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        result1 = runner.invoke(app, ["start"])
        assert result1.exit_code == 0

        # stop without --force calls _stop_active_instance → stop_instance
        result2 = runner.invoke(app, ["stop"])
        assert result2.exit_code == 0
        mock_manager.stop_instance.assert_called_once_with(1234, save=True)


class TestWorkbookCommands:
    """Tests for workbook CLI commands."""

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_open_command(self, mock_wb_class, mock_mgr_class):
        """Test workbook open command."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_info = WorkbookInfo(
            name="test.xlsx",
            full_path=test_file,
            read_only=False,
            saved=True,
            sheets_count=3,
        )
        mock_wb_mgr.open.return_value = mock_info

        result = runner.invoke(app, ["workbook", "open", str(test_file)])

        assert result.exit_code == 0
        assert "test.xlsx" in result.stdout
        assert "lecture/écriture" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_open_read_only_command(self, mock_wb_class, mock_mgr_class):
        """Test workbook open command with --read-only."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_info = WorkbookInfo(
            name="test.xlsx",
            full_path=test_file,
            read_only=True,
            saved=True,
            sheets_count=3,
        )
        mock_wb_mgr.open.return_value = mock_info

        result = runner.invoke(app, ["workbook", "open", str(test_file), "--read-only"])

        assert result.exit_code == 0
        assert "lecture seule" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_open_not_found(self, mock_wb_class, mock_mgr_class):
        """Test workbook open command with file not found."""
        from xlmanage.exceptions import WorkbookNotFoundError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/missing.xlsx")
        mock_wb_mgr.open.side_effect = WorkbookNotFoundError(test_file, "Not found")

        result = runner.invoke(app, ["workbook", "open", str(test_file)])

        assert result.exit_code == 1
        assert "Fichier introuvable" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_open_already_open_error(self, mock_wb_class, mock_mgr_class):
        """Test workbook open command when workbook already open."""
        from xlmanage.exceptions import WorkbookAlreadyOpenError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_wb_mgr.open.side_effect = WorkbookAlreadyOpenError(
            "test.xlsx", test_file, "Already open"
        )

        result = runner.invoke(app, ["workbook", "open", str(test_file)])

        assert result.exit_code == 1
        assert "déjà ouvert" in result.stdout
        assert "test.xlsx" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_open_excel_manage_error(self, mock_wb_class, mock_mgr_class):
        """Test workbook open command with ExcelManageError."""
        from xlmanage.exceptions import ExcelManageError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_wb_mgr.open.side_effect = ExcelManageError("Generic error")

        result = runner.invoke(app, ["workbook", "open", str(test_file)])

        assert result.exit_code == 1
        assert "Erreur" in result.stdout
        assert "Generic error" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_create_command(self, mock_wb_class, mock_mgr_class):
        """Test workbook create command."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/new.xlsx")
        mock_info = WorkbookInfo(
            name="new.xlsx",
            full_path=test_file,
            read_only=False,
            saved=True,
            sheets_count=1,
        )
        mock_wb_mgr.create.return_value = mock_info

        result = runner.invoke(app, ["workbook", "create", str(test_file)])

        assert result.exit_code == 0
        assert "new.xlsx" in result.stdout
        assert "Vierge" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_create_with_template(self, mock_wb_class, mock_mgr_class):
        """Test workbook create command with template."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/from_template.xlsx")
        template = Path("/tmp/template.xltx")
        mock_info = WorkbookInfo(
            name="from_template.xlsx",
            full_path=test_file,
            read_only=False,
            saved=True,
            sheets_count=3,
        )
        mock_wb_mgr.create.return_value = mock_info

        result = runner.invoke(
            app, ["workbook", "create", str(test_file), "--template", str(template)]
        )

        assert result.exit_code == 0
        assert "Basé sur" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_close_command(self, mock_wb_class, mock_mgr_class):
        """Test workbook close command."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/test.xlsx")
        result = runner.invoke(app, ["workbook", "close", str(test_file)])

        assert result.exit_code == 0
        mock_wb_mgr.close.assert_called_once_with(test_file, save=True, force=False)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_close_no_save(self, mock_wb_class, mock_mgr_class):
        """Test workbook close command with --no-save."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/test.xlsx")
        result = runner.invoke(app, ["workbook", "close", str(test_file), "--no-save"])

        assert result.exit_code == 0
        mock_wb_mgr.close.assert_called_once_with(test_file, save=False, force=False)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_save_command(self, mock_wb_class, mock_mgr_class):
        """Test workbook save command."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/test.xlsx")
        result = runner.invoke(app, ["workbook", "save", str(test_file)])

        assert result.exit_code == 0
        mock_wb_mgr.save.assert_called_once_with(test_file, output=None)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_save_as_command(self, mock_wb_class, mock_mgr_class):
        """Test workbook save command with --as (SaveAs)."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        test_file = Path("/tmp/test.xlsx")
        output = Path("/tmp/save_as.xlsx")
        result = runner.invoke(
            app, ["workbook", "save", str(test_file), "--as", str(output)]
        )

        assert result.exit_code == 0
        mock_wb_mgr.save.assert_called_once_with(test_file, output=output)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_list_command_empty(self, mock_wb_class, mock_mgr_class):
        """Test workbook list command with no workbooks."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        mock_wb_mgr.list.return_value = []

        result = runner.invoke(app, ["workbook", "list"])

        assert result.exit_code == 0
        assert "Aucun classeur ouvert" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorkbookManager")
    def test_workbook_list_command(self, mock_wb_class, mock_mgr_class):
        """Test workbook list command with workbooks."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_wb_mgr = Mock()
        mock_wb_class.return_value = mock_wb_mgr

        mock_wb_mgr.list.return_value = [
            WorkbookInfo(
                name="file1.xlsx",
                full_path=Path("/tmp/file1.xlsx"),
                read_only=False,
                saved=True,
                sheets_count=2,
            ),
            WorkbookInfo(
                name="file2.xlsx",
                full_path=Path("/tmp/file2.xlsx"),
                read_only=True,
                saved=False,
                sheets_count=5,
            ),
        ]

        result = runner.invoke(app, ["workbook", "list"])

        assert result.exit_code == 0
        assert "file1.xlsx" in result.stdout
        assert "file2.xlsx" in result.stdout
        assert "Classeurs ouverts" in result.stdout


class TestWorksheetCommands:
    """Tests for worksheet CLI commands."""

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_create_command(self, mock_ws_class, mock_mgr_class):
        """Test worksheet create command."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_info = WorksheetInfo(
            name="NewSheet",
            index=2,
            visible=True,
            rows_used=0,
            columns_used=0,
        )
        mock_ws_mgr.create.return_value = mock_info

        result = runner.invoke(app, ["worksheet", "create", "NewSheet"])

        assert result.exit_code == 0
        assert "NewSheet" in result.stdout
        assert "créée avec succès" in result.stdout
        mock_ws_mgr.create.assert_called_once_with("NewSheet", workbook=None)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_create_with_workbook(self, mock_ws_class, mock_mgr_class):
        """Test worksheet create command with --workbook option."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_info = WorksheetInfo(
            name="NewSheet",
            index=2,
            visible=True,
            rows_used=0,
            columns_used=0,
        )
        mock_ws_mgr.create.return_value = mock_info

        result = runner.invoke(
            app, ["worksheet", "create", "NewSheet", "--workbook", str(test_file)]
        )

        assert result.exit_code == 0
        assert "NewSheet" in result.stdout
        mock_ws_mgr.create.assert_called_once_with("NewSheet", workbook=test_file)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_create_name_error(self, mock_ws_class, mock_mgr_class):
        """Test worksheet create command with invalid name."""
        from xlmanage.exceptions import WorksheetNameError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.create.side_effect = WorksheetNameError(
            "Invalid/Name", "contains forbidden character '/'"
        )

        result = runner.invoke(app, ["worksheet", "create", "Invalid/Name"])

        assert result.exit_code == 1
        assert "Nom de feuille invalide" in result.stdout
        assert "Invalid/Name" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_create_already_exists(self, mock_ws_class, mock_mgr_class):
        """Test worksheet create command when worksheet already exists."""
        from xlmanage.exceptions import WorksheetAlreadyExistsError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.create.side_effect = WorksheetAlreadyExistsError(
            "Sheet1", "test.xlsx"
        )

        result = runner.invoke(app, ["worksheet", "create", "Sheet1"])

        assert result.exit_code == 1
        assert "déjà existante" in result.stdout
        assert "Sheet1" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_delete_command(self, mock_ws_class, mock_mgr_class):
        """Test worksheet delete command with --force."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        result = runner.invoke(app, ["worksheet", "delete", "OldSheet", "--force"])

        assert result.exit_code == 0
        assert "supprimée avec succès" in result.stdout
        assert "OldSheet" in result.stdout
        mock_ws_mgr.delete.assert_called_once_with("OldSheet", workbook=None)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_delete_with_confirmation_yes(self, mock_ws_class, mock_mgr_class):
        """Test worksheet delete command with user confirmation (yes)."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        result = runner.invoke(app, ["worksheet", "delete", "OldSheet"], input="y\n")

        assert result.exit_code == 0
        assert "supprimée avec succès" in result.stdout
        mock_ws_mgr.delete.assert_called_once_with("OldSheet", workbook=None)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_delete_with_confirmation_no(self, mock_ws_class, mock_mgr_class):
        """Test worksheet delete command with user confirmation (no)."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        result = runner.invoke(app, ["worksheet", "delete", "OldSheet"], input="n\n")

        assert result.exit_code == 0
        assert "annulée" in result.stdout
        mock_ws_mgr.delete.assert_not_called()

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_delete_not_found(self, mock_ws_class, mock_mgr_class):
        """Test worksheet delete command when worksheet not found."""
        from xlmanage.exceptions import WorksheetNotFoundError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.delete.side_effect = WorksheetNotFoundError(
            "MissingSheet", "test.xlsx"
        )

        result = runner.invoke(app, ["worksheet", "delete", "MissingSheet", "--force"])

        assert result.exit_code == 1
        assert "Feuille introuvable" in result.stdout
        assert "MissingSheet" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_delete_error(self, mock_ws_class, mock_mgr_class):
        """Test worksheet delete command with delete error."""
        from xlmanage.exceptions import WorksheetDeleteError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.delete.side_effect = WorksheetDeleteError(
            "LastSheet", "cannot delete the last visible worksheet"
        )

        result = runner.invoke(app, ["worksheet", "delete", "LastSheet", "--force"])

        assert result.exit_code == 1
        assert "Suppression impossible" in result.stdout
        assert "LastSheet" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_list_command_empty(self, mock_ws_class, mock_mgr_class):
        """Test worksheet list command with no worksheets."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.list.return_value = []

        result = runner.invoke(app, ["worksheet", "list"])

        assert result.exit_code == 0
        assert "Aucune feuille trouvée" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_list_command(self, mock_ws_class, mock_mgr_class):
        """Test worksheet list command with worksheets."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.list.return_value = [
            WorksheetInfo(
                name="Sheet1",
                index=1,
                visible=True,
                rows_used=10,
                columns_used=5,
            ),
            WorksheetInfo(
                name="Sheet2",
                index=2,
                visible=False,
                rows_used=0,
                columns_used=0,
            ),
            WorksheetInfo(
                name="Data",
                index=3,
                visible=True,
                rows_used=100,
                columns_used=20,
            ),
        ]

        result = runner.invoke(app, ["worksheet", "list"])

        assert result.exit_code == 0
        assert "Sheet1" in result.stdout
        assert "Sheet2" in result.stdout
        assert "Data" in result.stdout
        assert "3 trouvée(s)" in result.stdout
        assert "Feuilles du classeur" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_list_with_workbook(self, mock_ws_class, mock_mgr_class):
        """Test worksheet list command with --workbook option."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_ws_mgr.list.return_value = [
            WorksheetInfo(
                name="Sheet1",
                index=1,
                visible=True,
                rows_used=10,
                columns_used=5,
            ),
        ]

        result = runner.invoke(app, ["worksheet", "list", "--workbook", str(test_file)])

        assert result.exit_code == 0
        assert "test.xlsx" in result.stdout
        mock_ws_mgr.list.assert_called_once_with(workbook=test_file)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_copy_command(self, mock_ws_class, mock_mgr_class):
        """Test worksheet copy command."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_info = WorksheetInfo(
            name="Sheet1_Copy",
            index=2,
            visible=True,
            rows_used=10,
            columns_used=5,
        )
        mock_ws_mgr.copy.return_value = mock_info

        result = runner.invoke(app, ["worksheet", "copy", "Sheet1", "Sheet1_Copy"])

        assert result.exit_code == 0
        assert "copiée avec succès" in result.stdout
        assert "Sheet1" in result.stdout
        assert "Sheet1_Copy" in result.stdout
        mock_ws_mgr.copy.assert_called_once_with("Sheet1", "Sheet1_Copy", workbook=None)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_copy_with_workbook(self, mock_ws_class, mock_mgr_class):
        """Test worksheet copy command with --workbook option."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_info = WorksheetInfo(
            name="Copy",
            index=2,
            visible=True,
            rows_used=0,
            columns_used=0,
        )
        mock_ws_mgr.copy.return_value = mock_info

        result = runner.invoke(
            app,
            ["worksheet", "copy", "Source", "Copy", "--workbook", str(test_file)],
        )

        assert result.exit_code == 0
        assert "copiée avec succès" in result.stdout
        mock_ws_mgr.copy.assert_called_once_with("Source", "Copy", workbook=test_file)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_copy_source_not_found(self, mock_ws_class, mock_mgr_class):
        """Test worksheet copy command when source not found."""
        from xlmanage.exceptions import WorksheetNotFoundError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.copy.side_effect = WorksheetNotFoundError(
            "MissingSource", "test.xlsx"
        )

        result = runner.invoke(app, ["worksheet", "copy", "MissingSource", "Copy"])

        assert result.exit_code == 1
        assert "Feuille source introuvable" in result.stdout
        assert "MissingSource" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_copy_destination_exists(self, mock_ws_class, mock_mgr_class):
        """Test worksheet copy command when destination already exists."""
        from xlmanage.exceptions import WorksheetAlreadyExistsError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.copy.side_effect = WorksheetAlreadyExistsError(
            "ExistingSheet", "test.xlsx"
        )

        result = runner.invoke(app, ["worksheet", "copy", "Source", "ExistingSheet"])

        assert result.exit_code == 1
        assert "déjà existante" in result.stdout
        assert "ExistingSheet" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.WorksheetManager")
    def test_worksheet_copy_invalid_destination_name(self, mock_ws_class, mock_mgr_class):
        """Test worksheet copy command with invalid destination name."""
        from xlmanage.exceptions import WorksheetNameError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_ws_mgr = Mock()
        mock_ws_class.return_value = mock_ws_mgr

        mock_ws_mgr.copy.side_effect = WorksheetNameError(
            "Invalid/Name", "contains forbidden character '/'"
        )

        result = runner.invoke(app, ["worksheet", "copy", "Source", "Invalid/Name"])

        assert result.exit_code == 1
        assert "Nom de destination invalide" in result.stdout
        assert "Invalid/Name" in result.stdout


class TestTableCommands:
    """Tests for table CLI commands."""

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_command(self, mock_table_class, mock_mgr_class):
        """Test table create command."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_info = TableInfo(
            name="tbl_Sales",
            worksheet_name="Data",
            range_address="$A$1:$D$100",
            columns=["A", "B", "C", "D"],
            header_row="$A$1:$D$1",
            rows_count=99,
        )
        mock_table_mgr.create.return_value = mock_info

        result = runner.invoke(app, ["table", "create", "tbl_Sales", "A1:D100"])

        assert result.exit_code == 0
        assert "tbl_Sales" in result.stdout
        assert "créée avec succès" in result.stdout
        assert "99" in result.stdout
        mock_table_mgr.create.assert_called_once_with(
            "tbl_Sales", "A1:D100", worksheet=None, workbook=None
        )

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_with_worksheet(self, mock_table_class, mock_mgr_class):
        """Test table create command with --worksheet option."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_info = TableInfo(
            name="tbl_Data",
            worksheet_name="Sheet1",
            range_address="$A$1:$E$50",
            columns=["A", "B", "C", "D", "E"],
            header_row="$A$1:$E$1",
            rows_count=49,
        )
        mock_table_mgr.create.return_value = mock_info

        result = runner.invoke(
            app, ["table", "create", "tbl_Data", "A1:E50", "--worksheet", "Sheet1"]
        )

        assert result.exit_code == 0
        assert "tbl_Data" in result.stdout
        assert "Sheet1" in result.stdout
        mock_table_mgr.create.assert_called_once_with(
            "tbl_Data", "A1:E50", worksheet="Sheet1", workbook=None
        )

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_with_workbook(self, mock_table_class, mock_mgr_class):
        """Test table create command with --workbook option."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_info = TableInfo(
            name="tbl_Test",
            worksheet_name="Data",
            range_address="$A$1:$C$20",
            columns=["A", "B", "C"],
            header_row="$A$1:$C$1",
            rows_count=19,
        )
        mock_table_mgr.create.return_value = mock_info

        result = runner.invoke(
            app,
            ["table", "create", "tbl_Test", "A1:C20", "--workbook", str(test_file)],
        )

        assert result.exit_code == 0
        assert "tbl_Test" in result.stdout
        mock_table_mgr.create.assert_called_once_with(
            "tbl_Test", "A1:C20", worksheet=None, workbook=test_file
        )

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_name_error(self, mock_table_class, mock_mgr_class):
        """Test table create command with invalid name."""
        from xlmanage.exceptions import TableNameError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.create.side_effect = TableNameError(
            "1Invalid", "must start with letter or underscore"
        )

        result = runner.invoke(app, ["table", "create", "1Invalid", "A1:D10"])

        assert result.exit_code == 1
        assert "Nom de table invalide" in result.stdout
        assert "1Invalid" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_range_error(self, mock_table_class, mock_mgr_class):
        """Test table create command with invalid range."""
        from xlmanage.exceptions import TableRangeError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.create.side_effect = TableRangeError(
            "A1:Z", "invalid range syntax"
        )

        result = runner.invoke(app, ["table", "create", "tbl_Test", "A1:Z"])

        assert result.exit_code == 1
        assert "Plage invalide" in result.stdout
        assert "A1:Z" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_already_exists(self, mock_table_class, mock_mgr_class):
        """Test table create command when table already exists."""
        from xlmanage.exceptions import TableAlreadyExistsError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.create.side_effect = TableAlreadyExistsError(
            "tbl_Sales", "test.xlsx"
        )

        result = runner.invoke(app, ["table", "create", "tbl_Sales", "A1:D100"])

        assert result.exit_code == 1
        assert "Table déjà existante" in result.stdout
        assert "tbl_Sales" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_delete_command(self, mock_table_class, mock_mgr_class):
        """Test table delete command with --force."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        result = runner.invoke(app, ["table", "delete", "tbl_Old", "--force"])

        assert result.exit_code == 0
        assert "supprimée avec succès" in result.stdout
        assert "tbl_Old" in result.stdout
        mock_table_mgr.delete.assert_called_once_with(
            "tbl_Old", worksheet=None, workbook=None
        )

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_delete_with_confirmation_yes(self, mock_table_class, mock_mgr_class):
        """Test table delete command with user confirmation (yes)."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        result = runner.invoke(app, ["table", "delete", "tbl_Old"], input="y\n")

        assert result.exit_code == 0
        assert "supprimée avec succès" in result.stdout
        mock_table_mgr.delete.assert_called_once_with(
            "tbl_Old", worksheet=None, workbook=None
        )

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_delete_with_confirmation_no(self, mock_table_class, mock_mgr_class):
        """Test table delete command with user confirmation (no)."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        result = runner.invoke(app, ["table", "delete", "tbl_Old"], input="n\n")

        assert result.exit_code == 0
        assert "annulée" in result.stdout
        mock_table_mgr.delete.assert_not_called()

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_delete_with_worksheet(self, mock_table_class, mock_mgr_class):
        """Test table delete command with --worksheet option."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        result = runner.invoke(
            app, ["table", "delete", "tbl_Data", "--worksheet", "Sheet1", "--force"]
        )

        assert result.exit_code == 0
        assert "supprimée avec succès" in result.stdout
        mock_table_mgr.delete.assert_called_once_with(
            "tbl_Data", worksheet="Sheet1", workbook=None
        )

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_delete_not_found(self, mock_table_class, mock_mgr_class):
        """Test table delete command when table not found."""
        from xlmanage.exceptions import TableNotFoundError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.delete.side_effect = TableNotFoundError(
            "tbl_Missing", "any worksheet"
        )

        result = runner.invoke(app, ["table", "delete", "tbl_Missing", "--force"])

        assert result.exit_code == 1
        assert "Table introuvable" in result.stdout
        assert "tbl_Missing" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_list_command_empty(self, mock_table_class, mock_mgr_class):
        """Test table list command with no tables."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.list.return_value = []

        result = runner.invoke(app, ["table", "list"])

        assert result.exit_code == 0
        assert "Aucune table trouvée" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_list_command(self, mock_table_class, mock_mgr_class):
        """Test table list command with tables."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.list.return_value = [
            TableInfo(
                name="tbl_Sales",
                worksheet_name="Data",
                range_address="$A$1:$D$100",
                columns=["A", "B", "C", "D"],
                header_row="$A$1:$D$1",
                rows_count=99,
            ),
            TableInfo(
                name="tbl_Products",
                worksheet_name="Products",
                range_address="$A$1:$E$50",
                columns=["A", "B", "C", "D", "E"],
                header_row="$A$1:$E$1",
                rows_count=49,
            ),
        ]

        result = runner.invoke(app, ["table", "list"])

        assert result.exit_code == 0
        assert "tbl_Sales" in result.stdout
        assert "tbl_Products" in result.stdout
        assert "2 trouvée(s)" in result.stdout
        assert "Tables" in result.stdout
        mock_table_mgr.list.assert_called_once_with(worksheet=None, workbook=None)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_list_with_worksheet(self, mock_table_class, mock_mgr_class):
        """Test table list command with --worksheet option."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.list.return_value = [
            TableInfo(
                name="tbl_Data",
                worksheet_name="Sheet1",
                range_address="$A$1:$C$20",
                columns=["A", "B", "C"],
                header_row="$A$1:$C$1",
                rows_count=19,
            ),
        ]

        result = runner.invoke(app, ["table", "list", "--worksheet", "Sheet1"])

        assert result.exit_code == 0
        assert "tbl_Data" in result.stdout
        assert "Sheet1" in result.stdout
        mock_table_mgr.list.assert_called_once_with(worksheet="Sheet1", workbook=None)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_list_with_workbook(self, mock_table_class, mock_mgr_class):
        """Test table list command with --workbook option."""
        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        test_file = Path("/tmp/test.xlsx")
        mock_table_mgr.list.return_value = [
            TableInfo(
                name="tbl_Test",
                worksheet_name="Data",
                range_address="$A$1:$B$10",
                columns=["A", "B"],
                header_row="$A$1:$B$1",
                rows_count=9,
            ),
        ]

        result = runner.invoke(app, ["table", "list", "--workbook", str(test_file)])

        assert result.exit_code == 0
        assert "test.xlsx" in result.stdout
        assert "tbl_Test" in result.stdout
        mock_table_mgr.list.assert_called_once_with(worksheet=None, workbook=test_file)

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_worksheet_not_found(self, mock_table_class, mock_mgr_class):
        """Test table create when worksheet doesn't exist."""
        from xlmanage.exceptions import WorksheetNotFoundError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.create.side_effect = WorksheetNotFoundError(
            "MissingSheet", "test.xlsx"
        )

        result = runner.invoke(
            app, ["table", "create", "tbl_Test", "A1:D10", "--worksheet", "MissingSheet"]
        )

        assert result.exit_code == 1
        assert "Feuille introuvable" in result.stdout
        assert "MissingSheet" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_workbook_not_found(self, mock_table_class, mock_mgr_class):
        """Test table create when workbook not found."""
        from xlmanage.exceptions import WorkbookNotFoundError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        test_file = Path("/tmp/missing.xlsx")
        mock_table_mgr.create.side_effect = WorkbookNotFoundError(test_file, "Not found")

        result = runner.invoke(
            app, ["table", "create", "tbl_Test", "A1:D10", "--workbook", str(test_file)]
        )

        assert result.exit_code == 1
        assert "Classeur non trouvé" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_create_excel_manage_error(self, mock_table_class, mock_mgr_class):
        """Test table create with ExcelManageError."""
        from xlmanage.exceptions import ExcelManageError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.create.side_effect = ExcelManageError("Generic error")

        result = runner.invoke(app, ["table", "create", "tbl_Test", "A1:D10"])

        assert result.exit_code == 1
        assert "Erreur" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_delete_workbook_not_found(self, mock_table_class, mock_mgr_class):
        """Test table delete when workbook not found."""
        from xlmanage.exceptions import WorkbookNotFoundError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        test_file = Path("/tmp/missing.xlsx")
        mock_table_mgr.delete.side_effect = WorkbookNotFoundError(test_file, "Not found")

        result = runner.invoke(
            app, ["table", "delete", "tbl_Test", "--workbook", str(test_file), "--force"]
        )

        assert result.exit_code == 1
        assert "Classeur non trouvé" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_delete_excel_manage_error(self, mock_table_class, mock_mgr_class):
        """Test table delete with ExcelManageError."""
        from xlmanage.exceptions import ExcelManageError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.delete.side_effect = ExcelManageError("Delete failed")

        result = runner.invoke(app, ["table", "delete", "tbl_Test", "--force"])

        assert result.exit_code == 1
        assert "Erreur" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_list_workbook_not_found(self, mock_table_class, mock_mgr_class):
        """Test table list when workbook not found."""
        from xlmanage.exceptions import WorkbookNotFoundError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        test_file = Path("/tmp/missing.xlsx")
        mock_table_mgr.list.side_effect = WorkbookNotFoundError(test_file, "Not found")

        result = runner.invoke(app, ["table", "list", "--workbook", str(test_file)])

        assert result.exit_code == 1
        assert "Classeur non trouvé" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    @patch("xlmanage.cli.TableManager")
    def test_table_list_excel_manage_error(self, mock_table_class, mock_mgr_class):
        """Test table list with ExcelManageError."""
        from xlmanage.exceptions import ExcelManageError

        mock_mgr = Mock()
        mock_mgr_class.return_value.__enter__.return_value = mock_mgr

        mock_table_mgr = Mock()
        mock_table_class.return_value = mock_table_mgr

        mock_table_mgr.list.side_effect = ExcelManageError("List failed")

        result = runner.invoke(app, ["table", "list"])

        assert result.exit_code == 1
        assert "Erreur" in result.stdout
