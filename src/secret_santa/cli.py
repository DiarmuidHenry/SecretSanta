"""Console script for secret_santa."""

import typer
from rich.console import Console

from secret_santa import utils

app = typer.Typer()
console = Console()


@app.command()
def main():
    """Console script for secret_santa."""
    console.print("Replace this message by putting your code into "
               "secret_santa.cli.main")
    console.print("See Typer documentation at https://typer.tiangolo.com/")
    utils.do_something_useful()


if __name__ == "__main__":
    app()
