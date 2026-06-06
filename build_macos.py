#!/usr/bin/env python3
"""
Build script for macOS APP using PyInstaller
Usage: python build_macos.py
Creates: dist/MagicButton.app (standalone app for macOS)
"""

import os
import sys
import subprocess
from pathlib import Path


def main():
    print("=" * 60)
    print("Magic Button - macOS APP Builder")
    print("=" * 60)
    print()
    
    # Check if we're on macOS
    if sys.platform != "darwin":
        print("❌ This script only works on macOS!")
        print(f"Your platform: {sys.platform}")
        sys.exit(1)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("❌ PyInstaller not found!")
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller>=6.0.0"])
    
    project_dir = Path(__file__).parent
    os.chdir(project_dir)
    
    print("🍎 Building macOS APP...")
    print()
    
    # PyInstaller command for macOS
    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", "MagicButton",
        "--osx-bundle-identifier", "com.makar.magicbutton",
        "docx_tables_to_excel_app.py",
    ]
    
    print(f"Command: {' '.join(cmd)}")
    print()
    
    result = subprocess.run(cmd)
    
    print()
    if result.returncode == 0:
        app_path = project_dir / "dist" / "MagicButton.app"
        if app_path.exists():
            print("✅ BUILD SUCCESSFUL!")
            print()
            print(f"📍 Location: {app_path}")
            print(f"📊 Size: {sum(f.stat().st_size for f in app_path.rglob('*')) / (1024*1024):.1f} MB")
            print()
            print("🚀 You can now:")
            print(f"   1. Copy {app_path.name} to Applications folder")
            print("   2. Or share it with other macOS users")
            print("   3. They can just double-click it to run (no Python needed!)")
            print()
        else:
            print("❌ APP not found after build!")
            sys.exit(1)
    else:
        print("❌ BUILD FAILED!")
        sys.exit(1)


if __name__ == "__main__":
    main()
