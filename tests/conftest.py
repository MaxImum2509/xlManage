"""
Global pytest fixtures and hooks for xlmanage project.
"""

from collections.abc import Generator
from unittest.mock import Mock

import pytest


@pytest.fixture(scope="session")
def mock_excel_app() -> Generator[Mock]:
    """Session-wide mock for Excel application."""
    mock_app = Mock()
    mock_app.Visible = True
    mock_app.DisplayAlerts = False
    yield mock_app
    # Cleanup would go here if needed


@pytest.fixture(scope="function")
def mock_workbook() -> Generator[Mock]:
    """Function-scoped mock for Excel workbook."""
    mock_wb = Mock()
    mock_wb.Saved = True
    yield mock_wb
    # Cleanup would go here if needed


@pytest.fixture(scope="function")
def mock_worksheet() -> Generator[Mock]:
    """Function-scoped mock for Excel worksheet."""
    mock_ws = Mock()
    mock_ws.Name = "Sheet1"
    yield mock_ws
    # Cleanup would go here if needed


@pytest.fixture(autouse=True)
def setup_timeout(request):
    """Automatically apply timeout to all tests."""
    # Timeout is configured in pytest.ini
    pass


def pytest_configure(config):
    """Pytest configuration hook."""
    # Register custom markers
    config.addinivalue_line("markers", "com: tests involving COM automation")
    config.addinivalue_line("markers", "slow: tests that are slow to run")
    config.addinivalue_line("markers", "integration: integration tests")


def pytest_runtest_setup(item):
    """Setup hook for each test."""
    # Apply timeout marker if not already set
    if not any(marker.name == "timeout" for marker in item.iter_markers()):
        item.add_marker(pytest.mark.timeout(60))


def pytest_sessionfinish(session, exitstatus):
    """Session finish hook for cleanup."""
    # Any global cleanup would go here
    pass
