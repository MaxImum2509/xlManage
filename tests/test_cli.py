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

from unittest.mock import Mock, patch

from typer.testing import CliRunner

from xlmanage.cli import app
from xlmanage.excel_manager import InstanceInfo
from xlmanage.exceptions import ExcelConnectionError, ExcelManageError

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
        """Test stop command with default options."""
        # Setup mock
        mock_manager = Mock()
        mock_manager_class.return_value = mock_manager

        # Run command with force to skip confirmation
        result = runner.invoke(app, ["stop", "--force"])

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop.assert_called_once_with(save=True)
        assert "stopped successfully" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_no_save(self, mock_manager_class):
        """Test stop command with --no-save option."""
        # Setup mock
        mock_manager = Mock()
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["stop", "--force", "--no-save"])

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop.assert_called_once_with(save=False)

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_with_confirmation_yes(self, mock_manager_class):
        """Test stop command with user confirmation (yes)."""
        # Setup mock
        mock_manager = Mock()
        mock_manager_class.return_value = mock_manager

        # Run command with confirmation input
        result = runner.invoke(app, ["stop"], input="y\n")

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop.assert_called_once_with(save=True)

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_with_confirmation_no(self, mock_manager_class):
        """Test stop command with user confirmation (no)."""
        # Setup mock
        mock_manager = Mock()
        mock_manager_class.return_value = mock_manager

        # Run command with confirmation input
        result = runner.invoke(app, ["stop"], input="n\n")

        # Assertions
        assert result.exit_code == 0
        mock_manager.stop.assert_not_called()
        assert "cancelled" in result.stdout.lower()

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
        assert "No running Excel instances found" in result.stdout
        mock_manager.stop.assert_not_called()

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
        mock_manager_class.return_value = mock_manager

        # Run command with force
        result = runner.invoke(app, ["stop", "--all", "--force"])

        # Assertions
        assert result.exit_code == 0
        assert "Stopped" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_all_with_confirmation_no(self, mock_manager_class):
        """Test stop command with --all and user cancels."""
        # Setup mock
        mock_manager = Mock()
        instances = [
            InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)
        ]
        mock_manager.list_running_instances.return_value = instances
        mock_manager_class.return_value = mock_manager

        # Run command with confirmation input (no)
        result = runner.invoke(app, ["stop", "--all"], input="n\n")

        # Assertions
        assert result.exit_code == 0
        assert "cancelled" in result.stdout.lower()
        mock_manager.stop.assert_not_called()

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
        # First call succeeds, second call fails
        mock_manager.stop.side_effect = [None, Exception("Stop failed")]
        mock_manager_class.return_value = mock_manager

        # Run command with force
        result = runner.invoke(app, ["stop", "--all", "--force"])

        # Assertions
        assert result.exit_code == 0
        assert "Stopped" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_connection_error(self, mock_manager_class):
        """Test stop command with connection error."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.stop.side_effect = ExcelConnectionError(
            0x80080005, "Connection failed"
        )
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["stop", "--force"])

        # Assertions
        assert result.exit_code == 1
        assert "Connection Error" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_manage_error(self, mock_manager_class):
        """Test stop command with ExcelManageError."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.stop.side_effect = ExcelManageError("Management error")
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["stop", "--force"])

        # Assertions
        assert result.exit_code == 1
        assert "Excel management error" in result.stdout

    @patch("xlmanage.cli.ExcelManager")
    def test_stop_command_generic_error(self, mock_manager_class):
        """Test stop command with generic error."""
        # Setup mock to raise exception
        mock_manager = Mock()
        mock_manager.stop.side_effect = Exception("Unexpected error")
        mock_manager_class.return_value = mock_manager

        # Run command
        result = runner.invoke(app, ["stop", "--force"])

        # Assertions
        assert result.exit_code == 1
        assert "Unexpected Error" in result.stdout


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
        # Setup mock
        mock_manager = Mock()
        mock_info = InstanceInfo(pid=1234, visible=False, workbooks_count=0, hwnd=5678)
        mock_manager.start.return_value = mock_info
        mock_manager_class.return_value = mock_manager

        # Start instance
        result1 = runner.invoke(app, ["start"])
        assert result1.exit_code == 0

        # Stop instance
        result2 = runner.invoke(app, ["stop", "--force"])
        assert result2.exit_code == 0
        mock_manager.stop.assert_called_once()
