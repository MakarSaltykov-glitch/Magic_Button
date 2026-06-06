#!/usr/bin/env python3
"""
Build script for Windows EXE using PyInstaller
Usage: python build_windows.py
Creates: dist/MagicButton.exe (standalone executable for Windows)
"""

import os
import sys
import subprocess
from pathlib import Path


def main():
    print("=" * 60)
    print("Magic Button - Windows EXE Builder")
    print("=" * 60)
    print()
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("❌ PyInstaller not found!")
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller>=6.0.0"])
    
    project_dir = Path(__file__).parent
    os.chdir(project_dir)
    
    print("📦 Building Windows EXE...")
    print()
    
    # PyInstaller command
    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", "MagicButton",
        "--icon=NONE",
        "--add-data", f".:.",
        "docx_tables_to_excel_app.py",
    ]
    
    print(f"Command: {' '.join(cmd)}")
    print()
    
    result = subprocess.run(cmd)
    
    print()
    if result.returncode == 0:
        exe_path = project_dir / "dist" / "MagicButton.exe"
        if exe_path.exists():
            print("✅ BUILD SUCCESSFUL!")
            print()
            print(f"📍 Location: {exe_path}")
            print(f"📊 Size: {exe_path.stat().st_size / (1024*1024):.1f} MB")
            print()
            print("🚀 You can now:")
            print(f"   1. Share {exe_path.name} with other Windows users")
            print("   2. They can just double-click it to run (no Python needed!)")
            print()
        else:
            print("❌ EXE file not found after build!")
            sys.exit(1)
    else:
        print("❌ BUILD FAILED!")
        sys.exit(1)


if __name__ == "__main__":
    main()
