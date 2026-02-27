@echo off
REM Windows batch script to start the BOM Generator UI

cd /d "%~dp0"
echo Starting BOM Generator UI...
echo.

REM Check if streamlit is installed
python -m streamlit --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Streamlit is not installed!
    echo Please install it with: pip install streamlit
    pause
    exit /b 1
)

REM Run streamlit
python -m streamlit run src/bomgen/ui.py

pause
