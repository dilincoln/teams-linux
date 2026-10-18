#!/bin/sh
# Installs the app launcher in the system menu (current user).
# The file name must match "desktopName" in package.json: that's how KDE ties the
# window and the notifications to the app.
set -eu
app_dir=$(cd "$(dirname "$0")" && pwd)
apps="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
icons="${XDG_DATA_HOME:-$HOME/.local/share}/icons/hicolor"

mkdir -p "$apps"
for png in "$app_dir"/assets/icons/*x*.png; do
  size=$(basename "$png" .png)
  mkdir -p "$icons/$size/apps"
  cp "$png" "$icons/$size/apps/teams-linux.png"
done
cat > "$apps/teams-linux.desktop" <<DESKTOP
[Desktop Entry]
Type=Application
Name=Teams
GenericName=Microsoft Teams
Comment=Microsoft Teams (web) with multiple accounts
Exec="$app_dir/node_modules/electron/dist/electron" "$app_dir" %U
Icon=$app_dir/assets/icons/512x512.png
Terminal=false
Categories=Network;InstantMessaging;Chat;
StartupWMClass=teams-linux
StartupNotify=true
DESKTOP

update-desktop-database "$apps" 2>/dev/null || true
# Makes KDE reload launchers and icons (otherwise it may keep showing no icon).
kbuildsycoca6 >/dev/null 2>&1 || true
echo "Installed: $apps/teams-linux.desktop"
