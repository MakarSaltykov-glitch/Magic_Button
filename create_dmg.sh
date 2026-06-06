#!/bin/bash
# Create DMG installer for macOS
# Usage: bash create_dmg.sh
# Creates a nice DMG file for distribution

set -e

echo "🍎 Creating macOS DMG Installer..."
echo ""

APP_PATH="dist/MagicButton.app"
DMG_NAME="MagicButton-installer.dmg"
DMG_FOLDER="DMG_temp"

if [ ! -d "$APP_PATH" ]; then
    echo "❌ ERROR: $APP_PATH not found!"
    echo "Please run build_macos_installer.sh first"
    exit 1
fi

# Clean old DMG
rm -f "$DMG_NAME"
rm -rf "$DMG_FOLDER"

# Create temporary folder for DMG contents
mkdir -p "$DMG_FOLDER"

# Copy app to DMG folder
cp -r "$APP_PATH" "$DMG_FOLDER/"

# Create symlink to Applications folder
ln -s /Applications "$DMG_FOLDER/Applications"

# Create DMG
echo "Creating DMG file..."
hdiutil create -volname "Magic Button" \
               -srcfolder "$DMG_FOLDER" \
               -ov -format UDZO \
               "$DMG_NAME"

# Clean up temp folder
rm -rf "$DMG_FOLDER"

echo ""
echo "✅ DMG CREATED!"
echo ""
echo "📂 Location: $DMG_NAME"
echo "📊 Size: $(ls -lh "$DMG_NAME" | awk '{print $5}')"
echo ""
echo "🚀 Users can now:"
echo "   1. Download the DMG file"
echo "   2. Double-click to open it"
echo "   3. Drag MagicButton to Applications"
echo "   4. It's installed!"
echo ""
