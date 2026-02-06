"""
Tests for Excel instance stop methods.

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
from unittest.mock import Mock, MagicMock, patch, PropertyMock

try:
    import pywintypes
except ImportError:
    pywintypes = None

from xlmanage.excel_manager import ExcelManager, InstanceInfo
from xlmanage.exceptions import ExcelInstanceNotFoundError, ExcelRPCError


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_success(mocker):
    """Test successful stop of managed instance."""
    mock_wb = Mock()
    mock_app = Mock()
    mock_app.Workbooks = [mock_wb]
    mock_app.DisplayAlerts = True

    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app

    mgr.stop(save=True)

    # Vérifier le protocole
    assert mock_app.DisplayAlerts is False
    mock_wb.Close.assert_called_once_with(SaveChanges=True)
    assert mgr._app is None


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_no_save(mocker):
    """Test stop without saving."""
    mock_wb = Mock()
    mock_app = Mock()
    mock_app.Workbooks = [mock_wb]

    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app

    mgr.stop(save=False)

    mock_wb.Close.assert_called_once_with(SaveChanges=False)
    assert mgr._app is None


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_already_stopped():
    """Test stop when already stopped."""
    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = None

    # Ne doit pas lever d'erreur
    mgr.stop()

    assert mgr._app is None


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_with_rpc_error(mocker):
    """Test stop handles RPC errors gracefully."""
    mock_app = Mock()
    mock_app.DisplayAlerts = True
    mock_app.Workbooks = Mock(side_effect=Exception("RPC error"))

    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app

    # Ne doit pas lever d'erreur
    mgr.stop()

    assert mgr._app is None


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_multiple_workbooks(mocker):
    """Test stopping with multiple workbooks."""
    mock_wb1 = Mock()
    mock_wb2 = Mock()
    mock_wb3 = Mock()

    mock_app = Mock()
    mock_app.Workbooks = [mock_wb1, mock_wb2, mock_wb3]

    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app

    mgr.stop(save=True)

    # Vérifier que tous les classeurs sont fermés
    mock_wb1.Close.assert_called_once_with(SaveChanges=True)
    mock_wb2.Close.assert_called_once_with(SaveChanges=True)
    mock_wb3.Close.assert_called_once_with(SaveChanges=True)
    assert mgr._app is None


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_workbook_close_error(mocker):
    """Test stop continues when one workbook fails to close."""
    mock_wb1 = Mock()
    mock_wb1.Close.side_effect = Exception("Close failed")

    mock_wb2 = Mock()

    mock_app = Mock()
    mock_app.Workbooks = [mock_wb1, mock_wb2]

    mgr = ExcelManager.__new__(ExcelManager)
    mgr._app = mock_app

    # Ne doit pas lever d'erreur
    mgr.stop(save=True)

    # Le 2ème classeur doit être fermé malgré l'erreur sur le 1er
    mock_wb2.Close.assert_called_once_with(SaveChanges=True)
    assert mgr._app is None


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_instance_success(mocker):
    """Test stopping a specific instance by PID."""
    mock_wb = Mock()
    mock_app = Mock()
    mock_app.Workbooks = [mock_wb]
    mock_app.DisplayAlerts = True

    mock_info = InstanceInfo(pid=12345, visible=True, workbooks_count=1, hwnd=9999)

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(mock_app, mock_info)]

        mgr = ExcelManager()
        mgr.stop_instance(12345, save=True)

        assert mock_app.DisplayAlerts is False
        mock_wb.Close.assert_called_once()


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_instance_not_found(mocker):
    """Test error when PID doesn't exist."""
    with patch("xlmanage.excel_manager.enumerate_excel_instances", return_value=[]), \
         patch("xlmanage.excel_manager.enumerate_excel_pids", return_value=[]):

        mgr = ExcelManager()

        with pytest.raises(ExcelInstanceNotFoundError) as exc_info:
            mgr.stop_instance(99999)

        assert "99999" in str(exc_info.value.instance_id)


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_instance_disconnected(mocker):
    """Test error when instance is disconnected."""
    with patch("xlmanage.excel_manager.enumerate_excel_instances", return_value=[]), \
         patch("xlmanage.excel_manager.enumerate_excel_pids", return_value=[12345]):

        mgr = ExcelManager()

        with pytest.raises(ExcelRPCError) as exc_info:
            mgr.stop_instance(12345)

        assert "12345" in str(exc_info.value.message)


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_instance_rpc_error_during_close(mocker):
    """Test RPC error during instance shutdown."""
    mock_app = Mock()
    mock_app.Workbooks = []  # Rendre Workbooks itérable

    # Configurer DisplayAlerts pour lever une exception lors de l'assignment
    type(mock_app).DisplayAlerts = PropertyMock(side_effect=pywintypes.com_error(-2147417848, "RPC error", None, None))

    mock_info = InstanceInfo(pid=12345, visible=True, workbooks_count=0, hwnd=9999)

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(mock_app, mock_info)]

        mgr = ExcelManager()

        with pytest.raises(ExcelRPCError):
            mgr.stop_instance(12345)


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_all_success(mocker):
    """Test stopping all Excel instances."""
    mock_wb1 = Mock()
    mock_app1 = Mock()
    mock_app1.Workbooks = [mock_wb1]
    mock_app1.DisplayAlerts = True

    mock_wb2 = Mock()
    mock_app2 = Mock()
    mock_app2.Workbooks = [mock_wb2]
    mock_app2.DisplayAlerts = True

    mock_info1 = InstanceInfo(pid=12345, visible=True, workbooks_count=1, hwnd=1111)
    mock_info2 = InstanceInfo(pid=67890, visible=False, workbooks_count=1, hwnd=2222)

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(mock_app1, mock_info1), (mock_app2, mock_info2)]

        mgr = ExcelManager()
        stopped = mgr.stop_all(save=False)

        assert len(stopped) == 2
        assert 12345 in stopped
        assert 67890 in stopped
        assert mock_app1.DisplayAlerts is False
        assert mock_app2.DisplayAlerts is False


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_all_with_errors(mocker):
    """Test stop_all continues when one instance fails."""
    mock_app1 = Mock()
    mock_app1.DisplayAlerts = Mock(side_effect=Exception("Error"))

    mock_wb2 = Mock()
    mock_app2 = Mock()
    mock_app2.Workbooks = [mock_wb2]
    mock_app2.DisplayAlerts = True

    mock_info1 = InstanceInfo(pid=12345, visible=True, workbooks_count=0, hwnd=1111)
    mock_info2 = InstanceInfo(pid=67890, visible=True, workbooks_count=1, hwnd=2222)

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(mock_app1, mock_info1), (mock_app2, mock_info2)]

        mgr = ExcelManager()
        stopped = mgr.stop_all()

        # Seule la 2ème instance a été arrêtée
        assert stopped == [67890]


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_all_no_instances(mocker):
    """Test stop_all with no running instances."""
    with patch("xlmanage.excel_manager.enumerate_excel_instances", return_value=[]):
        mgr = ExcelManager()
        stopped = mgr.stop_all()

        assert stopped == []


@pytest.mark.skipif(pywintypes is None, reason="pywin32 not available")
def test_stop_all_multiple_workbooks_per_instance(mocker):
    """Test stop_all with multiple workbooks per instance."""
    mock_wb1 = Mock()
    mock_wb2 = Mock()
    mock_app = Mock()
    mock_app.Workbooks = [mock_wb1, mock_wb2]
    mock_app.DisplayAlerts = True

    mock_info = InstanceInfo(pid=12345, visible=True, workbooks_count=2, hwnd=9999)

    with patch("xlmanage.excel_manager.enumerate_excel_instances") as mock_enum:
        mock_enum.return_value = [(mock_app, mock_info)]

        mgr = ExcelManager()
        stopped = mgr.stop_all(save=True)

        assert stopped == [12345]
        mock_wb1.Close.assert_called_once_with(SaveChanges=True)
        mock_wb2.Close.assert_called_once_with(SaveChanges=True)
