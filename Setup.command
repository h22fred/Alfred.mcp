#!/bin/bash
# Double-click this file to install SC Engagement MCP.
# macOS will open a Terminal window and run the setup automatically.

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"
bash setup.sh

echo ""
echo "Press any key to close this window..."
read -n 1
