"""
Basic sample tests for xlManage.

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
