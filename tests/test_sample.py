"""
Sample test to verify pytest configuration.
"""

import pytest


def test_sample_passing():
    """A simple passing test."""
    assert 1 + 1 == 2


@pytest.mark.xfail(reason="This test is expected to fail - it's just for demonstration")
def test_sample_failing():
    """A simple failing test (should fail)."""
    # This test is expected to fail - it's just for demonstration
    assert 1 + 1 == 3


@pytest.mark.com
def test_com_mock(mock_excel_app, mock_workbook):
    """Test using COM mock fixtures."""
    assert mock_excel_app.Visible is True
    assert mock_workbook.Saved is True


@pytest.mark.slow
def test_slow_test():
    """Test marked as slow."""
    import time

    time.sleep(0.1)  # Simulate slow operation
    assert True
