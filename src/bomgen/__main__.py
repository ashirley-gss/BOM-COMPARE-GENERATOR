"""Allow running the CLI as a module: python -m bomgen.cli"""

from .cli import app

if __name__ == "__main__":
    app()
