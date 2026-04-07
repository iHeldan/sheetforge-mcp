import sys
import traceback
from collections.abc import Callable

import typer

from .server import run_sse, run_stdio, run_streamable_http

app = typer.Typer(help="SheetForge MCP")


def _run_server(start_fn: Callable[[], None]) -> None:
    """Run a server command without writing protocol-breaking output to stdout."""
    exit_code = 0
    try:
        start_fn()
    except KeyboardInterrupt:
        typer.echo("Shutting down server...", err=True)
        exit_code = 130
    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        traceback.print_exc(file=sys.stderr)
        exit_code = 1
    finally:
        typer.echo("Service stopped.", err=True)

    if exit_code:
        raise typer.Exit(code=exit_code)


@app.command()
def sse():
    """Start SheetForge MCP in SSE mode."""
    _run_server(run_sse)

@app.command()
def streamable_http():
    """Start SheetForge MCP in streamable HTTP mode."""
    _run_server(run_streamable_http)

@app.command()
def stdio():
    """Start SheetForge MCP in stdio mode."""
    _run_server(run_stdio)

if __name__ == "__main__":
    app()
