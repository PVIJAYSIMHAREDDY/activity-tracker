#!/bin/bash
# Build the Linux .deb installer for Daily Activity Tracker
# Run from the project root: bash build/build_linux.sh

set -e
cd "$(dirname "$0")/.."

APP="activity-tracker"
VERSION="1.4.0"
DEB_ROOT="build/_deb"
OUT="build/${APP}_${VERSION}_all.deb"

echo "==> Cleaning previous build..."
rm -rf "$DEB_ROOT"

echo "==> Creating package structure..."
mkdir -p "$DEB_ROOT/DEBIAN"
mkdir -p "$DEB_ROOT/opt/activity-tracker/templates"
mkdir -p "$DEB_ROOT/opt/activity-tracker/static"
mkdir -p "$DEB_ROOT/usr/local/bin"
mkdir -p "$DEB_ROOT/usr/share/applications"
mkdir -p "$DEB_ROOT/usr/share/icons/hicolor/64x64/apps"

echo "==> Copying app files..."
cp app.py main.py launch.sh icon.png "$DEB_ROOT/opt/activity-tracker/"
cp -r templates/. "$DEB_ROOT/opt/activity-tracker/templates/"
cp -r static/.   "$DEB_ROOT/opt/activity-tracker/static/" 2>/dev/null || true
cp icon.png "$DEB_ROOT/usr/share/icons/hicolor/64x64/apps/activity-tracker.png"

echo "==> Writing DEBIAN/control..."
cat > "$DEB_ROOT/DEBIAN/control" << EOF
Package: $APP
Version: $VERSION
Section: utils
Priority: optional
Architecture: all
Depends: python3 (>= 3.8), python3-flask, python3-webview, python3-openpyxl
Maintainer: Activity Tracker <noreply@activitytracker.app>
Description: Daily Activity Tracker
 Track daily tasks, habits, work hours, goals and diet with progress charts.
 Export your data to Excel. Supports dark mode.
EOF

echo "==> Writing desktop entry..."
cat > "$DEB_ROOT/usr/share/applications/activity-tracker.desktop" << 'EOF'
[Desktop Entry]
Version=1.0
Type=Application
Name=Daily Activity Tracker
Comment=Track tasks, habits, work hours, goals and diet
Exec=/usr/local/bin/activity-tracker
Icon=activity-tracker
Terminal=false
Categories=Utility;Office;
StartupNotify=true
EOF

echo "==> Writing launcher..."
cat > "$DEB_ROOT/usr/local/bin/activity-tracker" << 'EOF'
#!/bin/bash
exec python3 /opt/activity-tracker/main.py "$@"
EOF
chmod 755 "$DEB_ROOT/usr/local/bin/activity-tracker"

echo "==> Writing postinst script..."
cat > "$DEB_ROOT/DEBIAN/postinst" << 'EOF'
#!/bin/bash
set -e
chmod +x /opt/activity-tracker/launch.sh
update-desktop-database /usr/share/applications/ 2>/dev/null || true
gtk-update-icon-cache -f -t /usr/share/icons/hicolor 2>/dev/null || true
EOF
chmod 755 "$DEB_ROOT/DEBIAN/postinst"

echo "==> Writing prerm script..."
cat > "$DEB_ROOT/DEBIAN/prerm" << 'EOF'
#!/bin/bash
update-desktop-database /usr/share/applications/ 2>/dev/null || true
EOF
chmod 755 "$DEB_ROOT/DEBIAN/prerm"

echo "==> Building .deb package..."
dpkg-deb --build "$DEB_ROOT" "$OUT"
echo ""
echo "✅  Built: $OUT"
echo ""
echo "Install with:"
echo "  sudo dpkg -i $OUT"
echo "  sudo apt-get install -f   # fixes any missing dependencies"
