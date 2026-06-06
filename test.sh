#!/bin/bash
# Script to test Magic Button application before building EXE

echo "=== Magic Button - Test Suite ==="
echo ""

# Check Python version
echo "1. Checking Python version..."
python3 --version

# Check installed packages
echo ""
echo "2. Checking required packages..."
python3 -m pip list | grep -E "python-docx|pandas|openpyxl"

# Syntax check
echo ""
echo "3. Running syntax check..."
python3 -m py_compile docx_tables_to_excel_app.py
if [ $? -eq 0 ]; then
    echo "✓ Syntax OK"
else
    echo "✗ Syntax errors found"
    exit 1
fi

# Check if all imports are available
echo ""
echo "4. Checking imports..."
python3 -c "
import sys
try:
    from docx import Document
    print('✓ python-docx OK')
except ImportError as e:
    print(f'✗ python-docx error: {e}')
    sys.exit(1)

try:
    import pandas as pd
    print('✓ pandas OK')
except ImportError as e:
    print(f'✗ pandas error: {e}')
    sys.exit(1)

try:
    from openpyxl import load_workbook
    print('✓ openpyxl OK')
except ImportError as e:
    print(f'✗ openpyxl error: {e}')
    sys.exit(1)

try:
    import tkinter as tk
    print('✓ tkinter OK')
except ImportError as e:
    print(f'✗ tkinter error: {e}')
    sys.exit(1)
"

if [ $? -ne 0 ]; then
    echo "Some packages are missing. Install them with:"
    echo "  pip install -r requirements.txt"
    exit 1
fi

echo ""
echo "=== All tests passed! ==="
echo ""
echo "You can now:"
echo "1. Run the app: python docx_tables_to_excel_app.py"
echo "2. Build EXE:   python -m PyInstaller --noconfirm --onefile --windowed --name MagicButton docx_tables_to_excel_app.py"
