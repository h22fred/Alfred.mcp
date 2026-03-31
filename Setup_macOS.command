#!/bin/bash
# Double-click this file in Finder to install AlFred.mcp.
# macOS will open a Terminal window and run everything automatically.

set -e

REPO_URL="https://github.com/h22fred/Alfred.mcp.git"

# Ask role first so we know which folder to install into
echo ""
echo "▶ What is your role?"
echo ""
echo "   1) SC / SSC / Manager  — Solution Consulting"
echo "   2) Sales               — Account Executive / Manager"
echo ""
printf "   Enter 1 or 2 (default: 1): "
read -r VARIANT_CHOICE </dev/tty
if [ "$VARIANT_CHOICE" = "2" ]; then
  ALFRED_VARIANT="sales"
  INSTALL_DIR="$HOME/Documents/alfred.sales"
  echo "   ✅ Installing Alfred Sales to ~/Documents/alfred.sales"
else
  ALFRED_VARIANT="sc"
  INSTALL_DIR="$HOME/Documents/alfred.sc"
  echo "   ✅ Installing Alfred SC to ~/Documents/alfred.sc"
fi
export ALFRED_VARIANT

echo ""
echo "=================================================="
echo "  AlFred.mcp — Installer"
echo "=================================================="
echo ""

# ------------------------------------------------------------
# 1. Check Git — install Xcode Command Line Tools if missing
# ------------------------------------------------------------
if ! command -v git &>/dev/null; then
  echo "   ⚠️  Git not found — installing Xcode Command Line Tools..."
  xcode-select --install 2>/dev/null || true
  echo ""
  echo "   A popup will appear asking you to install the Command Line Tools."
  echo "   Click Install, wait for it to finish, then re-run this script:"
  echo ""
  echo "   bash ~/Downloads/Setup_macOS.command"
  echo ""
  echo "   (Full instructions and download: https://github.com/h22fred/Alfred.mcp)"
  echo ""
  echo "Press any key to close..."
  read -n 1 </dev/tty
  exit 1
fi

# ------------------------------------------------------------
# 2. Clone or update the repo
# ------------------------------------------------------------
if [ -d "$INSTALL_DIR/.git" ]; then
  echo "▶ Updating existing installation..."
  git -C "$INSTALL_DIR" fetch -q origin 2>/dev/null
  git -C "$INSTALL_DIR" reset -q --hard origin/main
  echo "   ✅ Updated to latest"
else
  echo "▶ Cloning alfred.mcp..."
  git clone "$REPO_URL" "$INSTALL_DIR"
  echo "   ✅ Cloned to $INSTALL_DIR"
fi

# ------------------------------------------------------------
# 3. Run setup
# ------------------------------------------------------------
echo ""
bash "$INSTALL_DIR/setup/setup.sh"

echo ""
echo "Press any key to close this window..."
read -n 1 </dev/tty
