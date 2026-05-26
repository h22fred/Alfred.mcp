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
npm install --prefix "$INSTALL_DIR" --no-fund --silent
echo "   ✅ Dependencies ready"

echo "▶ Rebuilding..."
npm run build --prefix "$INSTALL_DIR" --silent
echo "   ✅ Build complete"

echo ""
echo "✅ Alfred updated! Restart Claude Desktop to load the new version."
