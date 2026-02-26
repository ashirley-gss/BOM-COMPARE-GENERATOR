"""Simple launcher for BOM Generator UI - checks dependencies and runs Streamlit."""

import sys
import subprocess
from pathlib import Path
import os

def check_dependencies():
    """Check if required packages are installed."""
    missing = []
    
    try:
        import streamlit
        print(f"[OK] streamlit found (version {streamlit.__version__})")
    except ImportError:
        missing.append("streamlit")
        print("[X] streamlit not found")
    
    try:
        import openpyxl
        print(f"[OK] openpyxl found (version {openpyxl.__version__})")
    except ImportError:
        missing.append("openpyxl")
        print("[X] openpyxl not found")
    
    return missing

def install_dependencies(missing):
    """Install missing dependencies."""
    if not missing:
        return True
    
    print(f"\nInstalling missing packages: {', '.join(missing)}")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
        print("[OK] Dependencies installed successfully")
        return True
    except subprocess.CalledProcessError:
        print("[X] Failed to install dependencies")
        print("Please install manually with: pip install streamlit openpyxl typer pydantic")
        return False

def main():
    """Main launcher function."""
    print("=" * 60)
    print("BOM Generator UI Launcher")
    print("=" * 60)
    print()
    
    # Get project root
    script_dir = Path(__file__).parent.resolve()
    os.chdir(script_dir)
    
    print(f"Project directory: {script_dir}")
    print()
    
    # Check dependencies
    print("Checking dependencies...")
    missing = check_dependencies()
    print()
    
    # Install if needed
    if missing:
        response = input(f"Install missing packages ({', '.join(missing)})? [Y/n]: ").strip().lower()
        if response != 'n':
            if not install_dependencies(missing):
                return
        else:
            print("Please install dependencies manually:")
            print(f"  pip install {' '.join(missing)}")
            return
    
    # Check if UI file exists
    ui_file = script_dir / "src" / "bomgen" / "ui.py"
    if not ui_file.exists():
        print(f"[X] UI file not found: {ui_file}")
        return
    
    print(f"[OK] UI file found: {ui_file}")
    print()
    
    # Launch Streamlit
    print("=" * 60)
    print("Starting Streamlit UI...")
    print("=" * 60)
    print()
    print("The UI will open in your browser at: http://localhost:8501")
    print("Press Ctrl+C to stop the server")
    print()
    
    try:
        # Set env to skip email prompt; headless=false so the browser opens automatically
        env = os.environ.copy()
        env["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
        cmd = [
            sys.executable, "-m", "streamlit", "run", str(ui_file),
            "--server.headless", "false",
            "--browser.gatherUsageStats", "false",
        ]
        subprocess.run(cmd, check=True, env=env)
    except KeyboardInterrupt:
        print("\n\nShutting down...")
    except subprocess.CalledProcessError as e:
        print(f"\n[X] Error running Streamlit: {e}")
        print("\nTry running manually:")
        print(f"  cd {script_dir}")
        print(f"  python -m streamlit run src/bomgen/ui.py")
    except FileNotFoundError:
        print("\n[X] Python not found in PATH")
        print("Please make sure Python is installed and added to your PATH")

if __name__ == "__main__":
    main()
