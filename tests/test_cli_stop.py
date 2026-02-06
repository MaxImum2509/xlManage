"""
Tests for CLI stop command.

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

import pytest
from typer.testing import CliRunner
from unittest.mock import Mock, patch

from xlmanage.cli import app
from xlmanage.excel_manager import InstanceInfo
from xlmanage.exceptions import ExcelInstanceNotFoundError, ExcelRPCError

runner = CliRunner()


@pytest.fixture
def mock_manager():
    """Mock ExcelManager for testing."""
    with patch("xlmanage.cli.ExcelManager") as mock_cls:
        mock_instance = Mock()
        mock_cls.return_value = mock_instance
        yield mock_instance


def test_stop_active_instance(mock_manager):
    """Test stopping the active instance without arguments."""
    # Setup
    mock_info = InstanceInfo(pid=12345, hwnd=67890, workbooks_count=2, visible=True)
    mock_manager.get_running_instance.return_value = mock_info
    mock_manager.stop_instance.return_value = None

    # Execute
    result = runner.invoke(app, ["stop"])

    # Assert
    assert result.exit_code == 0
    mock_manager.get_running_instance.assert_called_once()
    mock_manager.stop_instance.assert_called_once_with(12345, save=True)
    assert "12345" in result.stdout
    assert "Arrêt Excel" in result.stdout


def test_stop_active_instance_none(mock_manager):
    """Test stopping when no active instance exists."""
    # Setup
    mock_manager.get_running_instance.return_value = None

    # Execute
    result = runner.invoke(app, ["stop"])

    # Assert
    assert result.exit_code == 0
    assert "Aucune instance Excel active" in result.stdout


def test_stop_specific_pid(mock_manager):
    """Test stopping a specific instance by PID."""
    # Setup
    mock_manager.stop_instance.return_value = None

    # Execute
    result = runner.invoke(app, ["stop", "54321"])

    # Assert
    assert result.exit_code == 0
    mock_manager.stop_instance.assert_called_once_with(54321, save=True)
    assert "54321" in result.stdout


def test_stop_no_save(mock_manager):
    """Test stopping without saving."""
    # Setup
    mock_info = InstanceInfo(pid=12345, hwnd=67890, workbooks_count=1, visible=True)
    mock_manager.get_running_instance.return_value = mock_info

    # Execute
    result = runner.invoke(app, ["stop", "--no-save"])

    # Assert
    assert result.exit_code == 0
    mock_manager.stop_instance.assert_called_once_with(12345, save=False)
    assert "Non" in result.stdout  # Sauvegarde : Non


def test_stop_all(mock_manager):
    """Test stopping all instances."""
    # Setup
    mock_instances = [
        InstanceInfo(pid=111, hwnd=1, workbooks_count=1, visible=True),
        InstanceInfo(pid=222, hwnd=2, workbooks_count=2, visible=True),
        InstanceInfo(pid=333, hwnd=3, workbooks_count=0, visible=False),
    ]
    mock_manager.list_running_instances.return_value = mock_instances
    mock_manager.stop_all.return_value = [111, 222, 333]

    # Execute
    result = runner.invoke(app, ["stop", "--all"])

    # Assert
    assert result.exit_code == 0
    mock_manager.stop_all.assert_called_once_with(save=True)
    assert "3 instance(s) arrêtée(s) avec succès" in result.stdout
    assert "111" in result.stdout
    assert "222" in result.stdout
    assert "333" in result.stdout


def test_stop_all_with_failures(mock_manager):
    """Test stopping all instances with some failures."""
    # Setup
    mock_instances = [
        InstanceInfo(pid=111, hwnd=1, workbooks_count=1, visible=True),
        InstanceInfo(pid=222, hwnd=2, workbooks_count=2, visible=True),
    ]
    mock_manager.list_running_instances.return_value = mock_instances
    mock_manager.stop_all.return_value = [111]  # Only 111 stopped, 222 failed

    # Execute
    result = runner.invoke(app, ["stop", "--all"])

    # Assert
    assert result.exit_code == 0
    assert "1 instance(s) arrêtée(s) avec succès" in result.stdout
    assert "1 instance(s) en échec" in result.stdout
    assert "Échec" in result.stdout


def test_stop_all_no_instances(mock_manager):
    """Test stopping all when no instances exist."""
    # Setup
    mock_manager.list_running_instances.return_value = []

    # Execute
    result = runner.invoke(app, ["stop", "--all"])

    # Assert
    assert result.exit_code == 0
    assert "Aucune instance Excel active" in result.stdout


def test_stop_force_single(mock_manager):
    """Test force killing a specific instance."""
    # Setup
    mock_manager.force_kill.return_value = None

    # Execute
    result = runner.invoke(app, ["stop", "99999", "--force"])

    # Assert
    assert result.exit_code == 0
    mock_manager.force_kill.assert_called_once_with(99999)
    assert "99999" in result.stdout
    assert "Force Kill" in result.stdout
    assert "ATTENTION" in result.stdout


def test_stop_force_all(mock_manager):
    """Test force killing all instances."""
    # Setup
    mock_instances = [
        InstanceInfo(pid=111, hwnd=1, workbooks_count=1, visible=True),
        InstanceInfo(pid=222, hwnd=2, workbooks_count=2, visible=True),
    ]
    mock_manager.list_running_instances.return_value = mock_instances
    mock_manager.force_kill.return_value = None

    # Execute
    result = runner.invoke(app, ["stop", "--all", "--force"])

    # Assert
    assert result.exit_code == 0
    assert mock_manager.force_kill.call_count == 2
    mock_manager.force_kill.assert_any_call(111)
    mock_manager.force_kill.assert_any_call(222)
    assert "ATTENTION" in result.stdout


def test_stop_force_active(mock_manager):
    """Test force killing the active instance."""
    # Setup
    mock_info = InstanceInfo(pid=77777, hwnd=1, workbooks_count=1, visible=True)
    mock_manager.get_running_instance.return_value = mock_info
    mock_manager.force_kill.return_value = None

    # Execute
    result = runner.invoke(app, ["stop", "--force"])

    # Assert
    assert result.exit_code == 0
    mock_manager.force_kill.assert_called_once_with(77777)
    assert "77777" in result.stdout


def test_stop_all_and_pid_error(mock_manager):
    """Test error when both --all and PID are specified."""
    # Execute
    result = runner.invoke(app, ["stop", "12345", "--all"])

    # Assert
    assert result.exit_code == 1
    assert "Impossible de spécifier --all ET un PID" in result.stdout


def test_stop_invalid_pid(mock_manager):
    """Test error with invalid PID format."""
    # Execute
    result = runner.invoke(app, ["stop", "not-a-number"])

    # Assert
    assert result.exit_code == 1
    assert "PID invalide" in result.stdout


def test_stop_instance_not_found(mock_manager):
    """Test error when instance is not found."""
    # Setup
    mock_manager.stop_instance.side_effect = ExcelInstanceNotFoundError(
        "12345", "Instance not found"
    )

    # Execute
    result = runner.invoke(app, ["stop", "12345"])

    # Assert
    assert result.exit_code == 1
    assert "Instance introuvable" in result.stdout


def test_stop_rpc_error(mock_manager):
    """Test error when RPC communication fails."""
    # Setup
    mock_manager.stop_instance.side_effect = ExcelRPCError(
        0x800706BE, "The remote procedure call failed"
    )

    # Execute
    result = runner.invoke(app, ["stop", "12345"])

    # Assert
    assert result.exit_code == 1
    assert "Erreur RPC" in result.stdout
    assert "Utilisez --force" in result.stdout


def test_stop_generic_error(mock_manager):
    """Test generic error handling."""
    # Setup
    mock_manager.stop_instance.side_effect = RuntimeError("Something went wrong")

    # Execute
    result = runner.invoke(app, ["stop", "12345"])

    # Assert
    assert result.exit_code == 1
    assert "Erreur" in result.stdout
    assert "Something went wrong" in result.stdout
