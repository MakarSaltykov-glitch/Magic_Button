#!/bin/bash
# Magic Button - macOS Standalone Builder
# Creates a standalone APP that users can just download and run
# No Python installation required!

echo ""
echo "╔════════════════════════════════════════════╗"
echo "║   Magic Button - macOS Standalone Builder  ║"
echo "╚════════════════════════════════════════════╝"
echo ""

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo "❌ ERROR: Python is not installed"
    echo ""
    echo "Please install Python from https://www.python.org/downloads/"
    echo "Or use: brew install python3"
    echo ""
    exit 1
fi

echo "✓ Python found: $(python3 --version)"
echo ""

# Check if PyInstaller is installed
if ! python3 -m pip show pyinstaller &> /dev/null; then
    echo "Installing PyInstaller..."
    python3 -m pip install pyinstaller>=6.0.0
    if [ $? -ne 0 ]; then
        echo "❌ Failed to install PyInstaller"
        exit 1
    fi
fi

echo "✓ PyInstaller ready"
echo ""

echo "🍎 Building standalone macOS APP..."
echo "This may take 1-2 minutes..."
echo ""

# Clean old builds
rm -rf dist build

# Run the builder
python3 build_macos.py
if [ $? -ne 0 ]; then
    echo ""
    echo "❌ BUILD FAILED"
    exit 1
fi

echo ""
echo "✅ BUILD COMPLETE!"
echo ""
echo "📂 The APP is in: dist/MagicButton.app"
echo ""
echo "🚀 Next steps:"
echo "   1. Test the APP on this computer"
echo "   2. If it works, you can:"
echo "      - Copy to Applications folder"
echo "      - Create a DMG installer"
echo "      - Upload to GitHub Releases"
echo ""
echo "📦 Users who receive the APP can:"
echo "   - Just double-click it"
echo "   - No installation needed"
echo "   - No Python required"
echo "   - Works on any macOS computer"
echo ""
