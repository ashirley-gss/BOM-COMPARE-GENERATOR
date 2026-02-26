"""Run the Streamlit UI for BOM Generator."""

import subprocess
import sys
from pathlib import Path
import os

if __name__ == "__main__":
    # Get the project root directory
    project_root = Path(__file__).parent.resolve()
    
    # Get the path to the UI module
    ui_path = project_root / "src" / "bomgen" / "ui.py"
    
    # Change to the project root directory
    os.chdir(project_root)
    
    # Run streamlit with proper arguments
    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(ui_path),
        "--server.port=8501",
        "--server.address=localhost"
    ]
    
    print(f"Starting Streamlit UI...")
    print(f"Project root: {project_root}")
    print(f"UI file: {ui_path}")
    print(f"Command: {' '.join(cmd)}")
    print("\nThe UI should open in your browser at http://localhost:8501")
    print("Press Ctrl+C to stop the server\n")
    
    try:
        subprocess.run(cmd, check=True)
    except KeyboardInterrupt:
        print("\n\nShutting down Streamlit server...")
    except Exception as e:
        print(f"\nError running Streamlit: {e}")
        print("\nTry running manually:")
        print(f"  cd {project_root}")
        print(f"  streamlit run src/bomgen/ui.py")
        sys.exit(1)
