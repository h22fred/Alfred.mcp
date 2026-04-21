#!/bin/bash
# Double-click this file in Finder to install AlFred.mcp.
# macOS will open a Terminal window and run everything automatically.

set -e

ZIP_URL="https://github.com/h22fred/Alfred.mcp/archive/refs/heads/main.zip"
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
# 1. Download or update the repo (no Git/Xcode required)
# ------------------------------------------------------------
# Prefer git if available (faster updates), fall back to zip download
if command -v git &>/dev/null; then
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
else
  echo "▶ Downloading alfred.mcp..."
  TMP_ZIP="/tmp/alfred-mcp-$$.zip"
  curl -fsSL "$ZIP_URL" -o "$TMP_ZIP"
  # Remove old install if it exists (non-git install)
  if [ -d "$INSTALL_DIR" ] && [ ! -d "$INSTALL_DIR/.git" ]; then
    rm -rf "$INSTALL_DIR"
  fi
  mkdir -p "$INSTALL_DIR"
  unzip -qo "$TMP_ZIP" -d /tmp
  # Move contents from extracted folder (Alfred.mcp-main/) into install dir
  cp -R /tmp/Alfred.mcp-main/* "$INSTALL_DIR/"
  rm -rf /tmp/Alfred.mcp-main "$TMP_ZIP"
  echo "   ✅ Downloaded to $INSTALL_DIR"
fi

# ------------------------------------------------------------
# 2. Run setup
# ------------------------------------------------------------
echo ""
bash "$INSTALL_DIR/setup/setup.sh"

echo ""
echo "Press any key to close this window..."
read -n 1 </dev/tty
