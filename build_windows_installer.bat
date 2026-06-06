@echo off
REM Magic Button - Windows Installer Builder
REM Creates a standalone EXE that users can just download and run
REM No Python installation required!

echo.
echo ╔════════════════════════════════════════════╗
echo ║  Magic Button - Windows Standalone Builder ║
echo ╚════════════════════════════════════════════╝
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ❌ ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from https://www.python.org/downloads/
    echo IMPORTANT: Check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo ✓ Python found
echo.

REM Check if PyInstaller is installed
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller>=6.0.0
    if errorlevel 1 (
        echo ❌ Failed to install PyInstaller
        pause
        exit /b 1
    )
)

echo ✓ PyInstaller ready
echo.

echo 🔨 Building standalone Windows EXE...
echo This may take 1-2 minutes...
echo.

REM Clean old builds
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM Run the builder
python build_windows.py
if errorlevel 1 (
    echo.
    echo ❌ BUILD FAILED
    pause
    exit /b 1
)

echo.
echo ✅ BUILD COMPLETE!
echo.
echo 📂 The EXE file is in: dist\MagicButton.exe
echo.
echo 🚀 Next steps:
echo    1. Test the EXE file on this computer
echo    2. If it works, you can:
echo       - Share it with other Windows users
echo       - Create an installer (MSI)
echo       - Upload to GitHub Releases
echo.
echo 📦 Users who receive the EXE can:
echo    - Just double-click it
echo    - No installation needed
echo    - No Python required
echo    - Works on any Windows computer
echo.

pause
