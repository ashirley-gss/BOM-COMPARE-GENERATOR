# Installation and Setup Guide

## Quick Start

1. **Make sure Python is installed**
   - Check by running: `python --version` or `python3 --version`
   - If not installed, download from https://www.python.org/downloads/

2. **Install dependencies**
   ```bash
   pip install streamlit openpyxl typer pydantic
   ```

3. **Run the UI**
   ```bash
   python launch_ui.py
   ```

## Detailed Steps

### Step 1: Verify Python Installation

Open a terminal/command prompt and run:
```bash
python --version
```

You should see something like: `Python 3.8.x` or higher.

If you get an error, Python is not in your PATH. You may need to:
- Reinstall Python and check "Add Python to PATH" during installation
- Or use the full path to Python (e.g., `C:\Python39\python.exe`)

### Step 2: Install Dependencies

Navigate to the `bom_generator` folder and run:

```bash
pip install streamlit openpyxl typer pydantic
```

Or install all at once:
```bash
pip install -e .
```

### Step 3: Run the Application

**Option A: Use the launcher (Recommended)**
```bash
python launch_ui.py
```

The launcher will:
- Check if dependencies are installed
- Offer to install missing packages
- Launch the Streamlit UI

**Option B: Run directly**
```bash
python -m streamlit run src/bomgen/ui.py
```

**Option C: Use the test script first**
```bash
python test_setup.py
```

This will show you what's installed and what's missing.

## Troubleshooting

### "Python is not recognized"
- Python is not in your PATH
- Reinstall Python and check "Add Python to PATH"
- Or use the full path: `C:\Python39\python.exe launch_ui.py`

### "streamlit is not recognized"
- Streamlit is not installed
- Run: `pip install streamlit`

### "ModuleNotFoundError: No module named 'bomgen'"
- Make sure you're in the `bom_generator` directory
- The imports should work automatically

### Browser doesn't open automatically
- Manually navigate to: `http://localhost:8501`
- Check the terminal for the exact URL

### Port 8501 is already in use
- Another Streamlit app is running
- Stop it or use a different port: `streamlit run src/bomgen/ui.py --server.port=8502`

## Getting Help

If you're still having issues:

1. Run the test script: `python test_setup.py`
2. Check the error messages in the terminal
3. Make sure you're in the `bom_generator` directory
4. Verify all dependencies are installed: `pip list`
