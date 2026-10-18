#!/bin/sh
# Instala o atalho do app no menu do sistema (usuário atual).
# O nome do arquivo precisa bater com "desktopName" do package.json: é por ele que o
# KDE associa a janela e as notificações ao app.
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
Comment=Microsoft Teams (web) com múltiplas contas
Exec="$app_dir/node_modules/electron/dist/electron" "$app_dir" %U
Icon=$app_dir/assets/icons/512x512.png
Terminal=false
Categories=Network;InstantMessaging;Chat;
StartupWMClass=teams-linux
StartupNotify=true
DESKTOP

update-desktop-database "$apps" 2>/dev/null || true
# Faz o KDE reler atalhos e ícones (sem isso pode continuar mostrando sem ícone).
kbuildsycoca6 >/dev/null 2>&1 || true
echo "Instalado: $apps/teams-linux.desktop"
