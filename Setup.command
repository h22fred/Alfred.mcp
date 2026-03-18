#!/bin/bash
# Double-click this file in Finder to install AlFred.mcp.
# macOS will open a Terminal window and run everything automatically.

set -e

REPO_URL="https://github.com/h22fred/alfred.mcp.git"
INSTALL_DIR="$HOME/Documents/alfred.mcp"

echo ""
echo "=================================================="
echo "  AlFred.mcp — Installer"
echo "=================================================="
echo ""

# ------------------------------------------------------------
# 1. Check Git
# ------------------------------------------------------------
if ! command -v git &>/dev/null; then
  echo "❌ Git not found. Install Xcode Command Line Tools:"
  echo "   xcode-select --install"
  echo ""
  echo "Press any key to close..."
  read -n 1
  exit 1
fi

# ------------------------------------------------------------
# 2. Clone or update the repo
# ------------------------------------------------------------
if [ -d "$INSTALL_DIR/.git" ]; then
  echo "▶ Updating existing installation..."
  git -C "$INSTALL_DIR" pull --ff-only
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
bash "$INSTALL_DIR/setup.sh"

echo ""
echo "Press any key to close this window..."
read -n 1
