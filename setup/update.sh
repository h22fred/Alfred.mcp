#!/bin/bash
# Alfred — one-shot update script
# Detects installed variant from ~/.alfred-config.json, syncs to latest, rebuilds.
# No questions asked. Safe to re-run.

set -e

# Detect install dir from config
VARIANT=$(python3 -c "
import json, os
f = os.path.expanduser('~/.alfred-config.json')
if os.path.exists(f):
    d = json.load(open(f))
    print(d.get('variant', 'sc'))
else:
    print('sc')
" 2>/dev/null)

if [ "$VARIANT" = "sales" ]; then
  INSTALL_DIR="$HOME/Documents/alfred.sales"
else
  INSTALL_DIR="$HOME/Documents/alfred.sc"
fi

if [ ! -d "$INSTALL_DIR/.git" ]; then
  echo "❌ Alfred not found at $INSTALL_DIR — run the full installer first."
  exit 1
fi

# Find npm — not in PATH when running via curl | bash (nvm not loaded)
NPM_PATH=""
for p in \
  /opt/homebrew/bin/npm \
  /usr/local/bin/npm \
  "$HOME/.nvm/versions/node/$(ls "$HOME/.nvm/versions/node/" 2>/dev/null | sort -V | tail -1)/bin/npm"; do
  if [ -x "$p" ]; then NPM_PATH="$p"; break; fi
done
if [ -z "$NPM_PATH" ] && command -v npm &>/dev/null; then
  NPM_PATH="$(command -v npm)"
fi
if [ -z "$NPM_PATH" ]; then
  echo "❌ npm not found — please run the full installer to reinstall Node.js:"
  echo "   curl -fsSL https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup_macOS.command | bash"
  exit 1
fi
NODE_DIR="$(dirname "$NPM_PATH")"

echo "▶ Updating Alfred ($VARIANT) at $INSTALL_DIR..."

git -C "$INSTALL_DIR" fetch -q origin
LOCAL=$(git -C "$INSTALL_DIR" rev-parse HEAD)
REMOTE=$(git -C "$INSTALL_DIR" rev-parse origin/main)

if [ "$LOCAL" = "$REMOTE" ]; then
  echo "   ✅ Already up to date — no rebuild needed."
  exit 0
fi

git -C "$INSTALL_DIR" reset -q --hard origin/main
echo "   ✅ Code updated"

echo "▶ Installing dependencies..."
PATH="$NODE_DIR:$PATH" "$NPM_PATH" install --prefix "$INSTALL_DIR" --no-fund --silent
echo "   ✅ Dependencies ready"

echo "▶ Rebuilding..."
PATH="$NODE_DIR:$PATH" "$NPM_PATH" run build --prefix "$INSTALL_DIR" --silent
echo "   ✅ Build complete"

echo ""
echo "✅ Alfred updated! Restart Claude Desktop to load the new version."
