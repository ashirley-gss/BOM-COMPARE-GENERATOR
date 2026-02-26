#!/bin/bash
# Linux/Mac shell script to start the BOM Generator UI

cd "$(dirname "$0")"

echo "Starting BOM Generator UI..."
echo ""

# Check if streamlit is installed
if ! python -m streamlit --version >/dev/null 2>&1; then
    echo "ERROR: Streamlit is not installed!"
    echo "Please install it with: pip install streamlit"
    exit 1
fi

# Run streamlit
python -m streamlit run src/bomgen/ui.py
