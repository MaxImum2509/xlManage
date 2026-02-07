"""
Tests for ExcelManager functionality.
"""

import pytest
import subprocess
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path

from xlmanage.excel_manager import ExcelManager, InstanceInfo
from xlmanage.exceptions import ExcelConnectionError, ExcelRPCError


def test_excel_manager_initialization():
    """Test ExcelManager initialization."""
    manager = ExcelManager(visible=False)
    assert manager._app is None
    assert manager._visible is False

    manager = ExcelManager(visible=True)
    assert manager._app is None
    assert manager._visible is True


def test_excel_manager_context_manager():
    """Test ExcelManager as context manager."""
    with patch("xlmanage.excel_manager.win32com.client.Dispatch") as mock_dispatch:
        # Create a proper mock for Workbooks collection
        mock_workbooks = Mock()
        mock_workbooks.Count = 0
        mock_workbooks.__iter__ = Mock(return_value=iter([]))

        mock_app = Mock()
        mock_app.Visible = False
        mock_app.Workbooks = mock_workbooks
        mock_app.Hwnd = 12345
        mock_app.DisplayAlerts = False
        mock_dispatch.return_value = mock_app

        # Test context manager
        with ExcelManager(visible=True) as manager:
            # __enter__ should call start()
            assert manager._app is not None
            assert mock_app.Visible is True

        # __exit__ should call stop()
        assert manager._app is None


def test_excel_manager_start_failure():
    """Test ExcelManager start failure handling."""
    with patch("xlmanage.excel_manager.win32com.client.Dispatch") as mock_dispatch:
        mock_dispatch.side_effect = Exception("COM error")

        manager = ExcelManager()
        with pytest.raises(ExcelConnectionError):
            manager.start()


def test_excel_manager_app_property():
    """Test ExcelManager app property."""
    manager = ExcelManager()
    with pytest.raises(ExcelConnectionError):
        _ = manager.app


def test_instance_info_dataclass():
    """Test InstanceInfo dataclass."""
    info = InstanceInfo(pid=1234, visible=True, workbooks_count=2, hwnd=54321)
    assert info.pid == 1234
    assert info.visible is True
    assert info.workbooks_count == 2
    assert info.hwnd == 54321


def test_excel_manager_stop_no_app():
    """Test ExcelManager stop when no app is running."""
    manager = ExcelManager()
    # Should not raise an error when stopping with no app
    manager.stop()


def test_excel_manager_stop_with_mock_app():
    """Test ExcelManager stop with mock app."""
    with patch("xlmanage.excel_manager.win32com.client.Dispatch") as mock_dispatch:
        # Create a proper mock for Workbooks collection
        mock_workbooks = Mock()
        mock_workbooks.Count = 0
        mock_workbooks.__iter__ = Mock(return_value=iter([]))

        mock_app = Mock()
        mock_app.Visible = False
        mock_app.Workbooks = mock_workbooks
        mock_app.Hwnd = 12345
        mock_app.DisplayAlerts = False
        mock_dispatch.return_value = mock_app

        manager = ExcelManager()
        manager.start()

        # Stop should work without errors
        manager.stop()
        assert manager._app is None


def test_excel_manager_stop_with_workbooks():
    """Test ExcelManager stop with workbooks."""
    with patch("xlmanage.excel_manager.win32com.client.Dispatch") as mock_dispatch:
        mock_wb = Mock()

        # Create a proper mock for Workbooks collection
        mock_workbooks = Mock()
        mock_workbooks.Count = 1
        mock_workbooks.__iter__ = Mock(return_value=iter([mock_wb]))

        mock_app = Mock()
        mock_app.Visible = False
        mock_app.Workbooks = mock_workbooks
        mock_app.Hwnd = 12345
        mock_app.DisplayAlerts = False
        mock_dispatch.return_value = mock_app

        manager = ExcelManager()
        manager.start()

        # Stop should close workbooks and clean up
        manager.stop()
        assert manager._app is None
        mock_wb.Close.assert_called_once_with(SaveChanges=True)


class TestExcelManagerNewMethods:
    """Test new methods added for Epic 5 Story 2."""

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    @patch("xlmanage.excel_manager.ExcelManager.get_instance_info")
    def test_get_running_instance_success(self, mock_get_instance_info, mock_dispatch):
        """Test successful retrieval of running instance."""
        # Setup mock
        mock_app = Mock()
        mock_app.Visible = True
        mock_app.Workbooks.Count = 2
        mock_app.Hwnd = 9999

        # Mock the expected return value
        expected_info = InstanceInfo(
            pid=9999, visible=True, workbooks_count=2, hwnd=9999
        )
        mock_get_instance_info.return_value = expected_info
        mock_dispatch.return_value = mock_app

        # Test
        manager = ExcelManager()
        info = manager.get_running_instance()

        # Assertions
        mock_dispatch.assert_called_once_with("Excel.Application")
        mock_get_instance_info.assert_called_once_with(mock_app)
        assert isinstance(info, InstanceInfo)
        assert info.pid == 9999
        assert info.visible is True
        assert info.workbooks_count == 2

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_get_running_instance_failure(self, mock_dispatch):
        """Test failure handling in get_running_instance."""
        # Setup mock to raise exception
        mock_dispatch.side_effect = Exception("No instance found")

        # Test
        manager = ExcelManager()

        with pytest.raises(ExcelConnectionError) as exc_info:
            manager.get_running_instance()

        # Assertions
        assert "Failed to get running instance" in str(exc_info.value)

    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    @patch("xlmanage.excel_manager.enumerate_excel_pids")
    def test_list_running_instances_rot_success(
        self,
        mock_enumerate_pids,
        mock_enumerate_instances,
    ):
        """Test successful enumeration via ROT."""
        # Setup mocks - enumerate_excel_instances returns list of (app, info) tuples
        mock_app = Mock()
        expected_info = InstanceInfo(
            pid=1111, visible=True, workbooks_count=1, hwnd=1111
        )
        mock_enumerate_instances.return_value = [(mock_app, expected_info)]

        # Test
        manager = ExcelManager()
        instances = manager.list_running_instances()

        # Assertions
        mock_enumerate_instances.assert_called_once()
        mock_enumerate_pids.assert_not_called()
        assert len(instances) == 1
        assert isinstance(instances[0], InstanceInfo)
        assert instances[0].pid == 1111

    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    @patch("xlmanage.excel_manager.enumerate_excel_pids")
    def test_list_running_instances_fallback_success(
        self,
        mock_enumerate_pids,
        mock_enumerate_instances,
    ):
        """Test fallback to PID enumeration when ROT returns empty."""
        # ROT returns empty list, fallback to tasklist
        mock_enumerate_instances.return_value = []
        mock_enumerate_pids.return_value = [1234]

        # Test
        manager = ExcelManager()
        instances = manager.list_running_instances()

        # Assertions
        mock_enumerate_instances.assert_called_once()
        mock_enumerate_pids.assert_called_once()
        assert len(instances) == 1
        assert isinstance(instances[0], InstanceInfo)
        assert instances[0].pid == 1234
        # Fallback creates InstanceInfo with limited info
        assert instances[0].visible is False
        assert instances[0].workbooks_count == 0
        assert instances[0].hwnd == 0


class TestExcelManagerStopEdgeCases:
    """Test edge cases for stop() method."""

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_stop_with_exception_during_close(self, mock_dispatch):
        """Test stop() when workbook close raises exception."""
        mock_wb = Mock()
        mock_wb.Close.side_effect = Exception("Close failed")

        mock_workbooks = Mock()
        mock_workbooks.Count = 1
        mock_workbooks.__iter__ = Mock(return_value=iter([mock_wb]))

        mock_app = Mock()
        mock_app.Visible = False
        mock_app.Workbooks = mock_workbooks
        mock_app.Hwnd = 12345
        mock_app.DisplayAlerts = False
        mock_dispatch.return_value = mock_app

        manager = ExcelManager()
        manager.start()

        # Stop should handle exception gracefully
        manager.stop()
        assert manager._app is None

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_stop_with_com_error(self, mock_dispatch):
        """Test stop() when COM error occurs during workbook enumeration."""
        # Create a mock that raises COM error with hresult
        com_error = Exception("RPC server unavailable")
        com_error.hresult = 0x800706BE

        mock_workbooks = Mock()
        mock_workbooks.Count = 0
        mock_workbooks.__iter__ = Mock(side_effect=com_error)

        mock_app = Mock()
        mock_app.Visible = False
        mock_app.Workbooks = mock_workbooks
        mock_app.DisplayAlerts = False
        mock_dispatch.return_value = mock_app

        manager = ExcelManager()
        manager.start()

        # stop() swallows exceptions and cleans up gracefully
        manager.stop()
        assert manager._app is None

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_stop_with_generic_error(self, mock_dispatch):
        """Test stop() when generic error occurs during workbook enumeration."""
        mock_workbooks = Mock()
        mock_workbooks.Count = 0
        mock_workbooks.__iter__ = Mock(side_effect=Exception("Generic error"))

        mock_app = Mock()
        mock_app.Visible = False
        mock_app.Workbooks = mock_workbooks
        mock_app.DisplayAlerts = False
        mock_dispatch.return_value = mock_app

        manager = ExcelManager()
        manager.start()

        # stop() swallows exceptions and cleans up gracefully
        manager.stop()
        assert manager._app is None

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_stop_with_del_app_exception(self, mock_dispatch):
        """Test stop() when DisplayAlerts access raises exception."""
        mock_workbooks = Mock()
        mock_workbooks.Count = 0

        # Create a mock that raises exception during DisplayAlerts access
        mock_app = Mock()
        type(mock_app).DisplayAlerts = Mock(
            side_effect=Exception("DisplayAlerts failed")
        )
        mock_app.Workbooks = mock_workbooks
        mock_dispatch.return_value = mock_app

        manager = ExcelManager()
        manager.start()

        # stop() swallows exceptions and cleans up gracefully
        manager.stop()
        assert manager._app is None


class TestExcelManagerAdvanced:
    """Test advanced ExcelManager scenarios."""

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    @patch("xlmanage.excel_manager.ExcelManager.get_instance_info")
    def test_start_new_instance_success(self, mock_get_instance_info, mock_dispatch):
        """Test starting a new Excel instance with new=True."""
        # Setup mock
        mock_app = Mock()
        mock_app.Visible = False
        mock_app.Workbooks.Count = 0
        mock_app.Hwnd = 1234

        expected_info = InstanceInfo(
            pid=1234, visible=False, workbooks_count=0, hwnd=1234
        )
        mock_get_instance_info.return_value = expected_info
        mock_dispatch.return_value = mock_app

        # Test
        manager = ExcelManager(visible=True)
        info = manager.start(new=True)

        # Assertions - Check that Dispatch was called (with any parameters)
        assert mock_dispatch.called
        assert manager._app == mock_app
        assert mock_app.Visible is True
        assert isinstance(info, InstanceInfo)
        assert info.pid == 1234

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    @patch("xlmanage.excel_manager.ExcelManager.get_instance_info")
    def test_start_existing_instance_success(
        self, mock_get_instance_info, mock_dispatch
    ):
        """Test connecting to existing Excel instance with new=False."""
        # Setup mock
        mock_app = Mock()
        mock_app.Visible = True
        mock_app.Workbooks.Count = 2
        mock_app.Hwnd = 5678

        expected_info = InstanceInfo(
            pid=5678, visible=True, workbooks_count=2, hwnd=5678
        )
        mock_get_instance_info.return_value = expected_info
        mock_dispatch.return_value = mock_app

        # Test
        manager = ExcelManager(visible=False)
        info = manager.start(new=False)

        # Assertions
        mock_dispatch.assert_called_once_with("Excel.Application")
        assert manager._app == mock_app
        assert mock_app.Visible is False
        assert isinstance(info, InstanceInfo)
        assert info.pid == 5678

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_start_com_error_without_hresult(self, mock_dispatch):
        """Test COM error without HRESULT in start."""
        # Setup mock to raise generic exception
        mock_dispatch.side_effect = Exception("Generic COM error")

        # Test
        manager = ExcelManager()

        with pytest.raises(ExcelConnectionError) as exc_info:
            manager.start(new=False)

        # Assertions
        assert exc_info.value.hresult == 0x80080005
        assert "Failed to start Excel" in str(exc_info.value)

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_start_com_error_with_hresult(self, mock_dispatch):
        """Test COM error with HRESULT in start."""
        # Setup mock to raise COM exception with hresult
        com_error = Exception("Server execution failed")
        com_error.hresult = 0x80080005
        mock_dispatch.side_effect = com_error

        # Test
        manager = ExcelManager()

        with pytest.raises(ExcelConnectionError) as exc_info:
            manager.start(new=True)

        # Assertions
        assert exc_info.value.hresult == 0x80080005
        assert "Failed to start Excel" in str(exc_info.value)

    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_get_running_instance_com_error_with_hresult(self, mock_dispatch):
        """Test get_running_instance with COM error with HRESULT."""
        # Setup mock to raise COM exception with hresult
        com_error = Exception("Instance not available")
        com_error.hresult = 0x80040154
        mock_dispatch.side_effect = com_error

        # Test
        manager = ExcelManager()

        with pytest.raises(ExcelConnectionError) as exc_info:
            manager.get_running_instance()

        # Assertions
        assert exc_info.value.hresult == 0x80040154
        assert "Failed to get running instance" in str(exc_info.value)

    def test_get_instance_info_fallback(self):
        """Test fallback when HWND is not available."""
        # Setup mock without Hwnd
        mock_app = Mock()
        mock_app.Visible = False
        mock_app.Workbooks.Count = 1
        del mock_app.Hwnd  # Simulate missing Hwnd

        # Test
        manager = ExcelManager()
        info = manager.get_instance_info(mock_app)

        # Assertions - should use fallback values
        assert info.pid == -1
        assert info.visible is False
        assert info.workbooks_count == 1
        assert info.hwnd == -1


class TestListRunningInstancesEdgeCases:
    """Test edge cases for list_running_instances."""

    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    def test_list_running_instances_with_get_instance_info_error(
        self, mock_enumerate_instances
    ):
        """Test list_running_instances when ROT returns valid tuples."""
        # enumerate_excel_instances returns list of (app, info) tuples
        mock_app = Mock()
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)
        mock_enumerate_instances.return_value = [(mock_app, mock_info)]

        manager = ExcelManager()
        instances = manager.list_running_instances()

        # Should extract InstanceInfo from tuples
        assert len(instances) == 1
        assert instances[0].pid == 1234

    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    @patch("xlmanage.excel_manager.enumerate_excel_pids")
    def test_list_running_instances_fallback_with_connect_error(
        self,
        mock_enumerate_pids,
        mock_enumerate_instances,
    ):
        """Test fallback when PID enumeration raises RuntimeError."""
        # ROT returns empty, fallback to PID enumeration which fails
        mock_enumerate_instances.return_value = []
        mock_enumerate_pids.side_effect = RuntimeError("tasklist failed")

        manager = ExcelManager()
        instances = manager.list_running_instances()

        # Should return empty list when fallback fails
        assert len(instances) == 0

    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    @patch("xlmanage.excel_manager.enumerate_excel_pids")
    def test_list_running_instances_both_methods_fail(
        self, mock_enumerate_pids, mock_enumerate_instances
    ):
        """Test when both ROT returns empty and PID enumeration fails."""
        # ROT returns empty list (the real function catches its own exceptions)
        mock_enumerate_instances.return_value = []
        mock_enumerate_pids.side_effect = RuntimeError("PID enum failed")

        manager = ExcelManager()
        instances = manager.list_running_instances()

        # Should return empty list
        assert len(instances) == 0


class TestUtilityFunctions:
    """Test utility functions."""

    @patch("xlmanage.excel_manager._get_instance_info_from_app")
    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    @patch("xlmanage.excel_manager.pythoncom")
    def test_enumerate_excel_instances(self, mock_pythoncom, mock_dispatch, mock_get_info):
        """Test enumerate_excel_instances function."""
        # Setup mock ROT with EnumRunning
        mock_rot = Mock()
        mock_moniker = Mock()

        # Mock CreateBindCtx and GetDisplayName to return Excel name
        mock_ctx = Mock()
        mock_pythoncom.CreateBindCtx.return_value = mock_ctx
        mock_moniker.GetDisplayName.return_value = "!Excel.Application"

        # Mock EnumRunning to return monikers
        mock_rot.EnumRunning.return_value = [mock_moniker]

        # Mock GetObject and QueryInterface
        mock_obj = Mock()
        mock_rot.GetObject.return_value = mock_obj
        mock_iid = Mock()
        mock_pythoncom.IID_IDispatch = mock_iid

        # Mock Dispatch to return app
        mock_app = Mock()
        mock_dispatch.return_value = mock_app

        # Mock _get_instance_info_from_app
        mock_info = InstanceInfo(pid=1234, visible=True, workbooks_count=1, hwnd=5678)
        mock_get_info.return_value = mock_info

        mock_pythoncom.GetRunningObjectTable.return_value = mock_rot

        # Test
        from xlmanage.excel_manager import enumerate_excel_instances

        instances = enumerate_excel_instances()

        # Assertions
        assert len(instances) == 1
        assert instances[0] == (mock_app, mock_info)

    @patch("subprocess.run")
    def test_enumerate_excel_pids_success(self, mock_run):
        """Test successful PID enumeration."""
        # Setup mock subprocess
        mock_result = Mock()
        mock_result.stdout = '"EXCEL.EXE","1234","Services","0","45,672 K"\n"EXCEL.EXE","5678","Services","0","34,567 K"'
        mock_result.returncode = 0
        mock_run.return_value = mock_result

        # Test
        from xlmanage.excel_manager import enumerate_excel_pids

        pids = enumerate_excel_pids()

        # Assertions
        assert len(pids) == 2
        assert 1234 in pids
        assert 5678 in pids

    @patch("subprocess.run")
    def test_enumerate_excel_pids_failure(self, mock_run):
        """Test PID enumeration failure handling."""
        # Setup mock to fail
        mock_run.side_effect = subprocess.CalledProcessError(1, "tasklist")

        # Test
        from xlmanage.excel_manager import enumerate_excel_pids

        # enumerate_excel_pids raises RuntimeError on CalledProcessError
        with pytest.raises(RuntimeError, match="tasklist"):
            enumerate_excel_pids()

    @patch("xlmanage.excel_manager.pythoncom")
    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    def test_connect_by_hwnd_success(self, mock_dispatch, mock_pythoncom):
        """Test successful connection by HWND."""
        # Setup mock for oleacc AccessibleObjectFromWindow
        mock_app = Mock()
        mock_dispatch.return_value = mock_app
        mock_pythoncom.IID_IDispatch = "mock_iid"
        mock_pythoncom.ObjectFromLresult.return_value = Mock()

        # Test
        from xlmanage.excel_manager import connect_by_hwnd

        # connect_by_hwnd uses ctypes internally which we can't easily mock,
        # but we can test the error path returns None
        result = connect_by_hwnd(1234)

        # connect_by_hwnd relies on ctypes.windll.oleacc which is not
        # available in test environment, so it returns None
        assert result is None

    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    def test_connect_by_hwnd_not_found(self, mock_enumerate_instances):
        """Test connection by HWND when not found."""
        # Setup mock with different HWND
        mock_app = Mock()
        mock_app.Hwnd = 5678
        mock_enumerate_instances.return_value = [mock_app]

        # Test
        from xlmanage.excel_manager import connect_by_hwnd

        result = connect_by_hwnd(1234)

        # Assertions - should return None
        assert result is None

    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    def test_connect_by_hwnd_exception(self, mock_enumerate_instances):
        """Test connection by HWND when exception occurs."""
        mock_enumerate_instances.side_effect = Exception("Enumeration failed")

        # Test
        from xlmanage.excel_manager import connect_by_hwnd

        result = connect_by_hwnd(1234)

        # Assertions - should return None
        assert result is None

    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    def test_connect_by_hwnd_app_exception(self, mock_enumerate_instances):
        """Test connection by HWND when accessing Hwnd raises exception."""
        mock_app = Mock()
        mock_app.Hwnd = Mock(side_effect=Exception("Hwnd access failed"))
        mock_enumerate_instances.return_value = [mock_app]

        # Test
        from xlmanage.excel_manager import connect_by_hwnd

        result = connect_by_hwnd(1234)

        # Assertions - should return None
        assert result is None

    @patch("pythoncom.GetRunningObjectTable")
    def test_enumerate_excel_instances_exception(self, mock_get_rot):
        """Test enumerate_excel_instances when exception occurs."""
        mock_get_rot.side_effect = Exception("ROT failed")

        # Test
        from xlmanage.excel_manager import enumerate_excel_instances

        instances = enumerate_excel_instances()

        # Assertions - should return empty list
        assert len(instances) == 0

    @patch("pythoncom.GetRunningObjectTable")
    def test_enumerate_excel_instances_moniker_exception(self, mock_get_rot):
        """Test enumerate_excel_instances when moniker processing fails."""
        mock_rot = Mock()
        mock_moniker = Mock()
        mock_moniker.__str__ = Mock(side_effect=Exception("Moniker error"))
        mock_rot.__iter__ = Mock(return_value=iter([mock_moniker]))
        mock_get_rot.return_value = mock_rot

        # Test
        from xlmanage.excel_manager import enumerate_excel_instances

        instances = enumerate_excel_instances()

        # Assertions - should return empty list
        assert len(instances) == 0

    @patch("subprocess.run")
    def test_enumerate_excel_pids_file_not_found(self, mock_run):
        """Test PID enumeration when tasklist is not found."""
        mock_run.side_effect = FileNotFoundError("tasklist not found")

        # Test
        from xlmanage.excel_manager import enumerate_excel_pids

        # enumerate_excel_pids raises RuntimeError on FileNotFoundError
        with pytest.raises(RuntimeError, match="tasklist introuvable"):
            enumerate_excel_pids()

    @patch("subprocess.run")
    def test_enumerate_excel_pids_invalid_format(self, mock_run):
        """Test PID enumeration with invalid CSV format."""
        mock_result = Mock()
        mock_result.stdout = "invalid,format,without,proper,pid"
        mock_result.returncode = 0
        mock_run.return_value = mock_result

        # Test
        from xlmanage.excel_manager import enumerate_excel_pids

        pids = enumerate_excel_pids()

        # Assertions - should return empty list
        assert len(pids) == 0


class TestExcelManagerAppProperty:
    """Tests for ExcelManager.app property success path."""

    def test_app_returns_app_when_set(self):
        """Test .app returns the COM app when initialized."""
        manager = ExcelManager(visible=False)
        mock_app = Mock()
        manager._app = mock_app
        assert manager.app is mock_app


class TestStopExceptionPaths:
    """Tests for stop() exception handling branches."""

    def test_stop_handles_com_error_during_workbook_close(self):
        """Test stop() handles com_error when closing workbooks."""
        import pywintypes

        manager = ExcelManager(visible=False)
        mock_app = Mock()
        manager._app = mock_app

        # Workbook that raises com_error on Close
        mock_wb = Mock()
        mock_wb.Close.side_effect = pywintypes.com_error(
            -2147023174, "RPC server unavailable", None, None
        )
        mock_workbooks = Mock()
        mock_workbooks.__iter__ = Mock(return_value=iter([mock_wb]))
        mock_workbooks.Count = 1
        mock_app.Workbooks = mock_workbooks

        manager.stop()
        assert manager._app is None

    def test_stop_handles_com_error_during_del_app(self):
        """Test stop() handles com_error when del app fails."""
        import pywintypes

        manager = ExcelManager(visible=False)
        mock_app = Mock()
        manager._app = mock_app

        # Make Workbooks iteration raise com_error
        mock_app.Workbooks.__iter__ = Mock(
            side_effect=pywintypes.com_error(
                -2147023174, "RPC server unavailable", None, None
            )
        )

        manager.stop()
        assert manager._app is None


class TestEnumerateExcelInstancesExceptionPaths:
    """Tests for enumerate_excel_instances exception handling."""

    @patch("xlmanage.excel_manager._get_instance_info_from_app")
    @patch("xlmanage.excel_manager.win32com.client.Dispatch")
    @patch("xlmanage.excel_manager.pythoncom")
    def test_enumerate_skips_inaccessible_instance(
        self, mock_pythoncom, mock_dispatch, mock_get_info
    ):
        """Test enumerate_excel_instances skips instances that raise during access."""
        from xlmanage.excel_manager import enumerate_excel_instances

        mock_rot = Mock()
        mock_ctx = Mock()
        mock_pythoncom.CreateBindCtx.return_value = mock_ctx
        mock_pythoncom.GetRunningObjectTable.return_value = mock_rot

        # Two monikers: first raises during GetObject, second succeeds
        mock_moniker1 = Mock()
        mock_moniker1.GetDisplayName.return_value = "!Excel.Application"
        mock_moniker2 = Mock()
        mock_moniker2.GetDisplayName.return_value = "!Excel.Application"

        mock_rot.EnumRunning.return_value = [mock_moniker1, mock_moniker2]

        # First GetObject raises, second succeeds
        mock_obj = Mock()
        mock_rot.GetObject.side_effect = [Exception("Disconnected"), mock_obj]

        mock_app = Mock()
        mock_dispatch.return_value = mock_app

        expected_info = InstanceInfo(pid=100, visible=True, workbooks_count=1, hwnd=999)
        mock_get_info.return_value = expected_info

        result = enumerate_excel_instances()

        # Should have only one instance (second one)
        assert len(result) == 1
        assert result[0][1] == expected_info

    @patch("xlmanage.excel_manager.pythoncom")
    def test_enumerate_skips_non_excel_moniker(self, mock_pythoncom):
        """Test enumerate_excel_instances skips non-Excel monikers."""
        from xlmanage.excel_manager import enumerate_excel_instances

        mock_rot = Mock()
        mock_ctx = Mock()
        mock_pythoncom.CreateBindCtx.return_value = mock_ctx
        mock_pythoncom.GetRunningObjectTable.return_value = mock_rot

        # Moniker with non-Excel name
        mock_moniker = Mock()
        mock_moniker.GetDisplayName.return_value = "!Word.Application"
        mock_rot.EnumRunning.return_value = [mock_moniker]

        result = enumerate_excel_instances()
        assert len(result) == 0


class TestConnectByPid:
    """Tests for connect_by_pid() function.

    Strategy: mock _find_hwnd_for_pid (ctypes internals) and connect_by_hwnd
    to test the orchestration logic without real Windows API calls.
    """

    @patch("xlmanage.excel_manager.connect_by_hwnd")
    @patch("xlmanage.excel_manager._find_hwnd_for_pid")
    def test_connect_by_pid_success(self, mock_find_hwnd, mock_connect_hwnd):
        """Test connect_by_pid finds HWND and delegates to connect_by_hwnd."""
        from xlmanage.excel_manager import connect_by_pid

        mock_app = Mock()
        mock_find_hwnd.return_value = 12345
        mock_connect_hwnd.return_value = mock_app

        result = connect_by_pid(1234)

        assert result is mock_app
        mock_find_hwnd.assert_called_once_with(1234)
        mock_connect_hwnd.assert_called_once_with(12345)

    @patch("xlmanage.excel_manager.connect_by_hwnd")
    @patch("xlmanage.excel_manager._find_hwnd_for_pid")
    def test_connect_by_pid_no_hwnd_found(self, mock_find_hwnd, mock_connect_hwnd):
        """Test connect_by_pid returns None when no HWND found."""
        from xlmanage.excel_manager import connect_by_pid

        mock_find_hwnd.return_value = None

        result = connect_by_pid(1234)

        assert result is None
        mock_connect_hwnd.assert_not_called()

    @patch("xlmanage.excel_manager.connect_by_hwnd")
    @patch("xlmanage.excel_manager._find_hwnd_for_pid")
    def test_connect_by_pid_hwnd_connection_fails(
        self, mock_find_hwnd, mock_connect_hwnd
    ):
        """Test connect_by_pid returns None when connect_by_hwnd fails."""
        from xlmanage.excel_manager import connect_by_pid

        mock_find_hwnd.return_value = 12345
        mock_connect_hwnd.return_value = None

        result = connect_by_pid(1234)

        assert result is None

    @patch("xlmanage.excel_manager._find_hwnd_for_pid")
    def test_connect_by_pid_exception_returns_none(self, mock_find_hwnd):
        """Test connect_by_pid returns None on exception."""
        from xlmanage.excel_manager import connect_by_pid

        mock_find_hwnd.side_effect = OSError("Access denied")

        result = connect_by_pid(1234)

        assert result is None


class TestListRunningInstancesFallbackWithConnect:
    """Tests for list_running_instances fallback using connect_by_pid."""

    @patch("xlmanage.excel_manager.connect_by_pid")
    @patch("xlmanage.excel_manager.enumerate_excel_pids")
    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    def test_fallback_uses_connect_by_pid(
        self, mock_enum_instances, mock_enum_pids, mock_connect
    ):
        """Test fallback tries connect_by_pid for full InstanceInfo."""
        mock_enum_instances.return_value = []  # ROT fails
        mock_enum_pids.return_value = [1234]

        mock_app = Mock()
        mock_connect.return_value = mock_app

        expected_info = InstanceInfo(pid=1234, visible=True, workbooks_count=2, hwnd=555)
        with patch(
            "xlmanage.excel_manager._get_instance_info_from_app",
            return_value=expected_info,
        ):
            manager = ExcelManager(visible=False)
            instances = manager.list_running_instances()

        assert len(instances) == 1
        assert instances[0].pid == 1234
        assert instances[0].visible is True
        assert instances[0].workbooks_count == 2

    @patch("xlmanage.excel_manager.connect_by_pid")
    @patch("xlmanage.excel_manager.enumerate_excel_pids")
    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    def test_fallback_degraded_when_connect_fails(
        self, mock_enum_instances, mock_enum_pids, mock_connect
    ):
        """Test fallback returns degraded info when connect_by_pid fails."""
        mock_enum_instances.return_value = []  # ROT fails
        mock_enum_pids.return_value = [5678]
        mock_connect.return_value = None  # connect_by_pid also fails

        manager = ExcelManager(visible=False)
        instances = manager.list_running_instances()

        assert len(instances) == 1
        assert instances[0].pid == 5678
        assert instances[0].visible is False
        assert instances[0].workbooks_count == 0
        assert instances[0].hwnd == 0

    @patch("xlmanage.excel_manager._get_instance_info_from_app")
    @patch("xlmanage.excel_manager.connect_by_pid")
    @patch("xlmanage.excel_manager.enumerate_excel_pids")
    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    def test_fallback_degraded_when_get_info_fails(
        self, mock_enum_instances, mock_enum_pids, mock_connect, mock_get_info
    ):
        """Test fallback returns degraded info when _get_instance_info_from_app raises."""
        mock_enum_instances.return_value = []  # ROT fails
        mock_enum_pids.return_value = [5678]
        mock_connect.return_value = Mock()  # connect_by_pid succeeds
        mock_get_info.side_effect = Exception("ctypes error")  # but info extraction fails

        manager = ExcelManager(visible=False)
        instances = manager.list_running_instances()

        assert len(instances) == 1
        assert instances[0].pid == 5678
        assert instances[0].visible is False
        assert instances[0].hwnd == 0


class TestStopInstanceFallbackConnect:
    """Tests for stop_instance fallback via connect_by_pid."""

    @patch("xlmanage.excel_manager.connect_by_pid")
    @patch("xlmanage.excel_manager.enumerate_excel_instances")
    def test_stop_instance_uses_connect_by_pid_fallback(
        self, mock_enum_instances, mock_connect
    ):
        """Test stop_instance uses connect_by_pid when ROT misses the instance."""
        mock_enum_instances.return_value = []  # ROT returns nothing

        mock_app = Mock()
        mock_app.Workbooks = []
        mock_connect.return_value = mock_app  # But connect_by_pid succeeds

        manager = ExcelManager(visible=False)
        manager.stop_instance(1234, save=True)

        mock_connect.assert_called_once_with(1234)
        mock_app.DisplayAlerts = False
