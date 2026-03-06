# BOM Generator

A Python tool used to generate a Bill of Material .xlsx file with user defined or randomly generated parts to be used in BOM Compare.

## Features

- **Web-based UI** (Streamlit) for building BOMs that work with BOM Compare
- **Hierarchical BOM structure**: Top parent (Level 0), Level 1 components, Level 2 and Level 3 sub-components, with parent selection per group
- **Template-driven output**: Use a BOM Compare template (.xlsx) so the generated file matches the expected format and field lengths
- **Manual or random entry** at each level: enter parts by hand or auto-generate part numbers and data (with optional long part numbers, 20–50 characters, when “Use Long Part” is enabled)
- **Use Long Part** option: when enabled, PartNo and Revision allow longer values (50 and 10 characters) and you can choose to generate long random part numbers for any level or group
- **Apply parent values**: Option to apply the parent’s Revision and/or Location to all component and sub-component parts
- **One-click export**: Generate and download the .xlsx file for use in BOM Compare
- **Command-line interface** for generate, compare, and create-template (see Usage)

## Installation

### Quick Install

1. **Install Python** (if not already installed)
   - Download from https://www.python.org/downloads/
   - Make sure to check "Add Python to PATH" during installation

2. **Install dependencies**
   ```bash
   pip install streamlit openpyxl typer pydantic
   ```

3. **Run the launcher**
   ```bash
   python launch_ui.py
   ```

### Alternative Installation Methods

**Option 1: Install as package**
```bash
pip install -e .
```

**Option 2: Install dependencies individually**
```bash
pip install openpyxl typer pydantic streamlit
```

**See [INSTALL.md](INSTALL.md) for detailed installation instructions and troubleshooting.**

## Usage

### Web UI (Recommended)

**Easiest way - Use the launcher:**
```bash
cd bom_generator
python launch_ui.py
```

The launcher will:
- Check if all dependencies are installed
- Offer to install missing packages automatically
- Launch the Streamlit UI

**Alternative methods:**

**Option 1: Direct Streamlit command**
```bash
cd bom_generator
python -m streamlit run src/bomgen/ui.py
```

**Option 2: Test setup first**
```bash
cd bom_generator
python test_setup.py
```

**Troubleshooting:**
- If the browser doesn't open automatically, navigate to `http://localhost:8501`
- If you get import errors, make sure you're in the `bom_generator` directory
- Run `python test_setup.py` to diagnose issues
- See [INSTALL.md](INSTALL.md) for detailed troubleshooting

The UI provides:
- Template file upload or selection
- Interactive form for entering BOM data
- Real-time validation
- One-click BOM generation and download

### Command-Line Interface

#### Generate a BOM

```bash
bomgen generate
```

This will prompt you for:
- Parent BOM part number
- Number of child components
- Details for each child component

#### Compare two BOMs

```bash
bomgen compare bom1.xlsx bom2.xlsx -o comparison.xlsx
```

#### Create a template

```bash
bomgen create-template -o template.xlsx
```

## Project Structure

```
bom_generator/
  templates/
    BOM_COMPARE_TEMPLATE.xlsx
  src/
    bomgen/
      __init__.py
      template.py      # Excel template handling
      models.py        # Data models
      cli.py           # Command-line interface
      ui.py            # Streamlit web UI
  run_ui.py            # Helper script to launch UI
  pyproject.toml
  README.md
```

## Development

Install development dependencies:

```bash
pip install -e ".[dev]"
```

## License

MIT
