"""
Tests for Excel force_kill method.

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

import logging
import pytest
import subprocess
from unittest.mock import Mock, patch

from xlmanage.excel_manager import ExcelManager
from xlmanage.exceptions import ExcelInstanceNotFoundError


def test_force_kill_success():
    """Test successful force kill."""
    mock_result = Mock()
    mock_result.stdout = "SUCCESS: The process with PID 12345 has been terminated."
    mock_result.stderr = ""

    with patch("subprocess.run", return_value=mock_result) as mock_run:
        mgr = ExcelManager()
        mgr.force_kill(12345)

        # Verify call to taskkill
        mock_run.assert_called_once_with(
            ["taskkill", "/f", "/pid", "12345"],
            capture_output=True,
            text=True,
            check=True,
            timeout=10,
        )


def test_force_kill_process_not_found():
    """Test error when process doesn't exist."""
    error = subprocess.CalledProcessError(
        returncode=128, cmd=["taskkill"], output="", stderr="ERROR: The process '99999' not found."
    )
    error.stdout = ""
    error.stderr = "ERROR: The process '99999' not found."

    with patch("subprocess.run", side_effect=error):
        mgr = ExcelManager()

        with pytest.raises(ExcelInstanceNotFoundError) as exc_info:
            mgr.force_kill(99999)

        assert "99999" in str(exc_info.value.instance_id)


def test_force_kill_access_denied():
    """Test error when access is denied."""
    error = subprocess.CalledProcessError(
        returncode=1, cmd=["taskkill"], output="ERROR: Access denied", stderr="ERROR: Access denied"
    )
    error.stdout = "ERROR: Access denied"
    error.stderr = ""

    with patch("subprocess.run", side_effect=error):
        mgr = ExcelManager()

        with pytest.raises(RuntimeError, match="Access denied"):
            mgr.force_kill(12345)


def test_force_kill_timeout():
    """Test timeout handling."""
    with patch("subprocess.run", side_effect=subprocess.TimeoutExpired("taskkill", 10)):
        mgr = ExcelManager()

        with pytest.raises(RuntimeError, match="Timeout"):
            mgr.force_kill(12345)


def test_force_kill_command_not_found():
    """Test error when taskkill is not available."""
    with patch("subprocess.run", side_effect=FileNotFoundError()):
        mgr = ExcelManager()

        with pytest.raises(RuntimeError, match="taskkill command not found"):
            mgr.force_kill(12345)


def test_force_kill_no_success_in_output():
    """Test error when SUCCESS is not in stdout."""
    mock_result = Mock()
    mock_result.stdout = "ERROR: Something went wrong"

    with patch("subprocess.run", return_value=mock_result):
        mgr = ExcelManager()

        with pytest.raises(RuntimeError, match="taskkill failed"):
            mgr.force_kill(12345)


def test_force_kill_logs_warning(caplog):
    """Test that force_kill logs a warning."""
    mock_result = Mock()
    mock_result.stdout = "SUCCESS"

    with patch("subprocess.run", return_value=mock_result):
        with caplog.at_level(logging.WARNING):
            mgr = ExcelManager()
            mgr.force_kill(12345)

            # Verify warning was logged
            assert "Force killing" in caplog.text
            assert "12345" in caplog.text


def test_force_kill_logs_success(caplog):
    """Test that force_kill logs success info."""
    mock_result = Mock()
    mock_result.stdout = "SUCCESS"

    with patch("subprocess.run", return_value=mock_result):
        with caplog.at_level(logging.INFO):
            mgr = ExcelManager()
            mgr.force_kill(12345)

            # Verify success was logged
            assert "Successfully force-killed" in caplog.text
            assert "12345" in caplog.text
