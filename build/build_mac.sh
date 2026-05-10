#!/bin/bash
# Build the macOS .dmg installer for Daily Activity Tracker
# Run this script on a macOS machine from the project root folder
# Requirements: Python 3.8+, pip, Homebrew (for create-dmg)

set -e
cd "$(dirname "$0")/.."

echo "==> Installing build dependencies..."
pip3 install pyinstaller pywebview flask openpyxl

echo "==> Converting icon to .icns..."
if command -v sips &>/dev/null && command -v iconutil &>/dev/null; then
    mkdir -p icon.iconset
    sips -z 16 16     icon.png --out icon.iconset/icon_16x16.png
    sips -z 32 32     icon.png --out icon.iconset/icon_16x16@2x.png
    sips -z 32 32     icon.png --out icon.iconset/icon_32x32.png
    sips -z 64 64     icon.png --out icon.iconset/icon_32x32@2x.png
    sips -z 128 128   icon.png --out icon.iconset/icon_128x128.png
    sips -z 256 256   icon.png --out icon.iconset/icon_128x128@2x.png
    sips -z 256 256   icon.png --out icon.iconset/icon_256x256.png
    sips -z 512 512   icon.png --out icon.iconset/icon_256x256@2x.png
    sips -z 512 512   icon.png --out icon.iconset/icon_512x512.png
    iconutil -c icns icon.iconset -o icon.icns
    rm -rf icon.iconset
    echo "   icon.icns created"
else
    echo "   sips/iconutil not found, skipping icon conversion"
fi

echo "==> Building with PyInstaller..."
pyinstaller ActivityTracker.spec --clean --noconfirm

APP="dist/ActivityTracker.app"
DMG_NAME="build/ActivityTracker_macOS.dmg"

echo "==> Creating .dmg..."
if command -v create-dmg &>/dev/null; then
    create-dmg \
        --volname "Daily Activity Tracker" \
        --volicon "icon.icns" \
        --window-pos 200 120 \
        --window-size 600 400 \
        --icon-size 100 \
        --icon "ActivityTracker.app" 175 190 \
        --hide-extension "ActivityTracker.app" \
        --app-drop-link 425 190 \
        "$DMG_NAME" \
        "$APP"
else
    echo "   create-dmg not found. Installing via Homebrew..."
    brew install create-dmg
    create-dmg \
        --volname "Daily Activity Tracker" \
        --window-size 600 400 \
        --icon "ActivityTracker.app" 175 190 \
        --app-drop-link 425 190 \
        "$DMG_NAME" \
        "$APP"
fi

echo ""
echo "✅  Built: $DMG_NAME"
echo "   Distribute this file — users drag ActivityTracker.app to Applications."
