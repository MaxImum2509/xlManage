"""
Tests for Excel instance enumeration functions.

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
import subprocess
from unittest.mock import Mock, patch, MagicMock

try:
    import pythoncom
    import pywintypes
except ImportError:
    pythoncom = None
    pywintypes = None

from xlmanage.excel_manager import (
    enumerate_excel_instances,
    enumerate_excel_pids,
    connect_by_hwnd,
    _get_instance_info_from_app,
    ExcelManager,
    InstanceInfo,
)


@pytest.mark.skipif(pythoncom is None, reason="pywin32 not available")
def test_enumerate_excel_instances_rot_error(mocker):
    """Test fallback when ROT is inaccessible."""
    with patch("xlmanage.excel_manager.pythoncom.GetRunningObjectTable", side_effect=Exception("ROT error")):
        instances = enumerate_excel_instances()
        assert instances == []


def test_enumerate_excel_pids_success(mocker):
    """Test enumeration via tasklist."""
    tasklist_output = '"EXCEL.EXE","12345","Console","1","100,000 K"\n"EXCEL.EXE","67890","Console","1","120,000 K"\n'

    mock_result = Mock()
    mock_result.stdout = tasklist_output

    with patch("subprocess.run", return_value=mock_result):
        pids = enumerate_excel_pids()

        assert len(pids) == 2
        assert 12345 in pids
        assert 67890 in pids


def test_enumerate_excel_pids_no_instances(mocker):
    """Test tasklist with no Excel instances."""
    mock_result = Mock()
    mock_result.stdout = "INFO: No tasks found.\n"

    with patch("subprocess.run", return_value=mock_result):
        pids = enumerate_excel_pids()
        assert pids == []


def test_enumerate_excel_pids_empty_output(mocker):
    """Test tasklist with empty output."""
    mock_result = Mock()
    mock_result.stdout = ""

    with patch("subprocess.run", return_value=mock_result):
        pids = enumerate_excel_pids()
        assert pids == []


def test_enumerate_excel_pids_timeout(mocker):
    """Test timeout handling."""
    with patch("subprocess.run", side_effect=subprocess.TimeoutExpired("tasklist", 10)):
        with pytest.raises(RuntimeError, match="Timeout"):
            enumerate_excel_pids()


def test_enumerate_excel_pids_command_not_found(mocker):
    """Test error when tasklist command doesn't exist."""
    with patch("subprocess.run", side_effect=FileNotFoundError()):
        with pytest.raises(RuntimeError, match="tasklist introuvable"):
            enumerate_excel_pids()


def test_enumerate_excel_pids_called_process_error(mocker):
    """Test error when tasklist fails."""
    with patch("subprocess.run", side_effect=subprocess.CalledProcessError(1, "tasklist")):
        with pytest.raises(RuntimeError, match="Échec de tasklist"):
            enumerate_excel_pids()


@pytest.mark.skipif(pythoncom is None, reason="pywin32 not available")
def test_list_running_instances_via_rot(mocker):
    """Test list_running_instances using ROT."""
    mock_info = InstanceInfo(pid=12345, visible=True, workbooks_count=2, hwnd=9999)

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(Mock(), mock_info)]

        mgr = ExcelManager()
        instances = mgr.list_running_instances()

        assert len(instances) == 1
        assert instances[0].pid == 12345


@pytest.mark.skipif(pythoncom is None, reason="pywin32 not available")
def test_list_running_instances_fallback_tasklist(mocker):
    """Test list_running_instances fallback to tasklist."""
    with patch("xlmanage.excel_manager.enumerate_excel_instances", return_value=[]), \
         patch("xlmanage.excel_manager.enumerate_excel_pids", return_value=[12345, 67890]):

        mgr = ExcelManager()
        instances = mgr.list_running_instances()

        assert len(instances) == 2
        assert instances[0].pid == 12345
        assert instances[1].pid == 67890
        # Info limitée avec tasklist
        assert instances[0].visible is False
        assert instances[0].workbooks_count == 0


@pytest.mark.skipif(pythoncom is None, reason="pywin32 not available")
def test_list_running_instances_both_fail(mocker):
    """Test list_running_instances when both methods fail."""
    with patch("xlmanage.excel_manager.enumerate_excel_instances", return_value=[]), \
         patch("xlmanage.excel_manager.enumerate_excel_pids", side_effect=RuntimeError()):

        mgr = ExcelManager()
        instances = mgr.list_running_instances()

        assert instances == []


@pytest.mark.skipif(pythoncom is None, reason="pywin32 not available")
def test_connect_by_hwnd_failure(mocker):
    """Test connection failure by HWND."""
    with patch("ctypes.windll.oleacc.AccessibleObjectFromWindow", return_value=1):
        app = connect_by_hwnd(12345)
        assert app is None


@pytest.mark.skipif(pythoncom is None, reason="pywin32 not available")
def test_connect_by_hwnd_exception(mocker):
    """Test connection handles exceptions."""
    with patch("ctypes.windll.oleacc.AccessibleObjectFromWindow", side_effect=Exception()):
        app = connect_by_hwnd(12345)
        assert app is None
