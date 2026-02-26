"""Test script to verify Python setup and dependencies."""

import sys
from pathlib import Path

print("=" * 60)
print("BOM Generator - Setup Test")
print("=" * 60)
print(f"\nPython version: {sys.version}")
print(f"Python executable: {sys.executable}")
print(f"Current directory: {Path.cwd()}")

# Test imports
print("\n" + "=" * 60)
print("Testing imports...")
print("=" * 60)

try:
    import streamlit
    print(f"✓ streamlit version: {streamlit.__version__}")
except ImportError as e:
    print(f"✗ streamlit not installed: {e}")
    print("  Install with: pip install streamlit")

try:
    import openpyxl
    print(f"✓ openpyxl version: {openpyxl.__version__}")
except ImportError as e:
    print(f"✗ openpyxl not installed: {e}")
    print("  Install with: pip install openpyxl")

try:
    import typer
    print(f"✓ typer installed")
except ImportError as e:
    print(f"✗ typer not installed: {e}")
    print("  Install with: pip install typer")

try:
    import pydantic
    print(f"✓ pydantic version: {pydantic.__version__}")
except ImportError as e:
    print(f"✗ pydantic not installed: {e}")
    print("  Install with: pip install pydantic")

# Test project structure
print("\n" + "=" * 60)
print("Testing project structure...")
print("=" * 60)

project_root = Path(__file__).parent
ui_file = project_root / "src" / "bomgen" / "ui.py"
cli_file = project_root / "src" / "bomgen" / "cli.py"

print(f"Project root: {project_root}")
print(f"UI file exists: {ui_file.exists()} ({ui_file})")
print(f"CLI file exists: {cli_file.exists()} ({cli_file})")

# Test imports from project
print("\n" + "=" * 60)
print("Testing project imports...")
print("=" * 60)

sys.path.insert(0, str(project_root))

try:
    from bomgen.cli import TEMPLATE_HEADERS, REQUIRED_FIELDS
    print(f"✓ Successfully imported from bomgen.cli")
    print(f"  Template headers: {len(TEMPLATE_HEADERS)} columns")
    print(f"  Required fields: {REQUIRED_FIELDS}")
except Exception as e:
    print(f"✗ Failed to import from bomgen.cli: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Setup test complete!")
print("=" * 60)
