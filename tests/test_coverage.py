"""
Test to verify coverage reporting works.
"""

from typer.testing import CliRunner

from xlmanage.cli import app


def test_cli_help():
    """Test that CLI help works."""
    runner = CliRunner()
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "Usage:" in result.stdout


def test_main_entry():
    """Test that main_entry function works."""
    from xlmanage.cli import main_entry

    # Just test that it can be called without error
    try:
        main_entry()
    except SystemExit:
        pass  # Expected when running CLI without arguments


def test_cli_if_main():
    """Test that CLI can be run directly."""
    # This covers the if __name__ == "__main__" block
    import subprocess

    result = subprocess.run(
        ["poetry", "run", "python", "src/xlmanage/cli.py", "--help"],
        capture_output=True,
        text=True,
        timeout=10,
    )
    assert result.returncode == 0
    assert "Usage:" in result.stdout


def test_version_command():
    """Test the version command directly."""
    # Test that version function works
    import io
    from contextlib import redirect_stdout

    from xlmanage.cli import version

    f = io.StringIO()
    with redirect_stdout(f):
        version()
    output = f.getvalue()
    assert "xlmanage version 0.1.0" in output
