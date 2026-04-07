import pytest
import typer

from excel_mcp import __main__ as cli


def test_stdio_shutdown_messages_go_to_stderr(monkeypatch, capsys):
    def fake_run_stdio():
        raise KeyboardInterrupt

    monkeypatch.setattr(cli, "run_stdio", fake_run_stdio)

    with pytest.raises(typer.Exit) as exc_info:
        cli.stdio()

    assert exc_info.value.exit_code == 130
    captured = capsys.readouterr()
    assert captured.out == ""
    assert "Shutting down server..." in captured.err
    assert "Service stopped." in captured.err
