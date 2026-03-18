#!/bin/bash
set -e

# ============================================================
# SC Engagement MCP — Setup Script
# ============================================================
# Run this once to install everything and configure Claude Desktop.
# Requirements: macOS, Google Chrome, Claude Desktop

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CLAUDE_CONFIG="$HOME/Library/Application Support/Claude/claude_desktop_config.json"
CHROMELINK_APP="$HOME/Desktop/ChromeLink.app"

echo ""
echo "=================================================="
echo "  SC Engagement MCP — Setup"
echo "=================================================="
echo ""

# ------------------------------------------------------------
# 1. Check Node.js
# ------------------------------------------------------------
echo "▶ Checking Node.js..."
NODE_PATH=""
for p in /opt/homebrew/bin/node /usr/local/bin/node; do
  if [ -x "$p" ]; then NODE_PATH="$p"; break; fi
done

if [ -z "$NODE_PATH" ]; then
  echo ""
  echo "❌ Node.js not found. Install it with Homebrew:"
  echo "   /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
  echo "   brew install node"
  exit 1
fi

NODE_DIR="$(dirname "$NODE_PATH")"
echo "   ✅ Node.js found at $NODE_PATH"

# ------------------------------------------------------------
# 2. Install dependencies and build
# ------------------------------------------------------------
echo ""
echo "▶ Installing dependencies..."
PATH="$NODE_DIR:$PATH" npm install --prefix "$SCRIPT_DIR"

echo ""
echo "▶ Building MCP server..."
PATH="$NODE_DIR:$PATH" npm run build --prefix "$SCRIPT_DIR"
echo "   ✅ Build complete"

# ------------------------------------------------------------
# 3. Configure Claude Desktop
# ------------------------------------------------------------
echo ""
echo "▶ Configuring Claude Desktop..."

DIST_PATH="$SCRIPT_DIR/dist/index.js"
MCP_ENTRY=$(cat <<EOF
{
  "command": "$NODE_PATH",
  "args": ["$DIST_PATH"]
}
EOF
)

if [ ! -f "$CLAUDE_CONFIG" ]; then
  # Create fresh config
  mkdir -p "$(dirname "$CLAUDE_CONFIG")"
  echo "{\"mcpServers\":{\"sc-engagement\":$MCP_ENTRY}}" > "$CLAUDE_CONFIG"
  echo "   ✅ Created Claude Desktop config"
else
  # Check if already configured
  if grep -q "sc-engagement" "$CLAUDE_CONFIG" 2>/dev/null; then
    echo "   ✅ Already configured in Claude Desktop"
  else
    # Merge into existing config using Python (available on all Macs)
    python3 - <<PYEOF
import json, sys

config_path = """$CLAUDE_CONFIG"""
dist_path = """$DIST_PATH"""
node_path = """$NODE_PATH"""

with open(config_path, "r") as f:
    config = json.load(f)

if "mcpServers" not in config:
    config["mcpServers"] = {}

config["mcpServers"]["sc-engagement"] = {
    "command": node_path,
    "args": [dist_path]
}

with open(config_path, "w") as f:
    json.dump(config, f, indent=2)

print("   ✅ Added sc-engagement to Claude Desktop config")
PYEOF
  fi
fi

# ------------------------------------------------------------
# 4. Create ChromeLink.app on Desktop (plain shell bundle — no AppleScript)
# ------------------------------------------------------------
echo ""
echo "▶ Creating ChromeLink.app on Desktop..."

# Remove old versions
[ -d "$CHROMELINK_APP" ] && rm -rf "$CHROMELINK_APP"
[ -f "$HOME/Desktop/ChromeLink.command" ] && rm -f "$HOME/Desktop/ChromeLink.command"

mkdir -p "$CHROMELINK_APP/Contents/MacOS"

cat > "$CHROMELINK_APP/Contents/MacOS/ChromeLink" << 'SHELLEOF'
#!/bin/bash
notify() { osascript -e "display notification \"$1\" with title \"ChromeLink\"" 2>/dev/null; }

# Already running?
if curl -s --max-time 1 http://localhost:9222/json/version > /dev/null 2>&1; then
  notify "Already running — you're good to use Claude!"
  exit 0
fi

mkdir -p ~/.chromelink-profile
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --remote-debugging-port=9222 \
  --user-data-dir=~/.chromelink-profile \
  --no-first-run \
  --no-default-browser-check \
  "https://servicenow.crm.dynamics.com" \
  "https://outlook.office.com" \
  "https://teams.microsoft.com" \
  > /dev/null 2>&1 &

# First run detection — profile dir will be nearly empty on first launch
PROFILE_SIZE=$(du -sk ~/.chromelink-profile 2>/dev/null | cut -f1)
if [ -z "$PROFILE_SIZE" ] || [ "$PROFILE_SIZE" -lt 500 ]; then
  notify "First time setup: log into Dynamics, Outlook and Teams in this window. You only do this once!"
else
  notify "Launched — ready for Claude!"
fi
SHELLEOF

chmod +x "$CHROMELINK_APP/Contents/MacOS/ChromeLink"

cat > "$CHROMELINK_APP/Contents/Info.plist" << 'PLISTEOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleExecutable</key><string>ChromeLink</string>
  <key>CFBundleIdentifier</key><string>com.servicenow.chromelink</string>
  <key>CFBundleName</key><string>ChromeLink</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleVersion</key><string>1.0</string>
  <key>LSUIElement</key><true/>
</dict>
</plist>
PLISTEOF

echo "   ✅ ChromeLink.app created on Desktop"
echo "   ℹ️  First launch: right-click → Open (one-time macOS approval)"

# ------------------------------------------------------------
# 5. Teams webhook config
# ------------------------------------------------------------
echo ""
echo "▶ Setting up Teams webhook for hygiene sweep notifications..."

CONFIG_FILE="$HOME/.sc-engagement-config.json"
EXISTING_WEBHOOK=""
if [ -f "$CONFIG_FILE" ]; then
  EXISTING_WEBHOOK=$(python3 -c "import json; d=json.load(open('$CONFIG_FILE')); print(d.get('teamsWebhook',''))" 2>/dev/null)
fi

if [ -n "$EXISTING_WEBHOOK" ]; then
  echo "   ✅ Teams webhook already configured"
  echo "      $EXISTING_WEBHOOK"
  echo ""
  printf "   Replace it? (press Enter to keep, or paste a new URL): "
  read -r NEW_WEBHOOK
  if [ -n "$NEW_WEBHOOK" ]; then
    EXISTING_WEBHOOK="$NEW_WEBHOOK"
    python3 -c "import json; f='$CONFIG_FILE'; d=json.load(open(f)) if __import__('os').path.exists(f) else {}; d['teamsWebhook']='$NEW_WEBHOOK'; json.dump(d,open(f,'w'),indent=2)"
    echo "   ✅ Webhook updated"
  fi
else
  echo ""
  echo "   To get a webhook URL:"
  echo "   Teams → any channel → ··· → Connectors → Incoming Webhook → configure → copy URL"
  echo ""
  printf "   Paste your Teams incoming webhook URL (or press Enter to skip): "
  read -r NEW_WEBHOOK
  if [ -n "$NEW_WEBHOOK" ]; then
    python3 -c "import json; f='$CONFIG_FILE'; d=json.load(open(f)) if __import__('os').path.exists(f) else {}; d['teamsWebhook']='$NEW_WEBHOOK'; json.dump(d,open(f,'w'),indent=2)"
    echo "   ✅ Webhook saved to $CONFIG_FILE"
  else
    echo "   ⏭  Skipped — hygiene sweep will run without Teams notifications"
    echo "      Run setup again anytime to add it, or ask Claude to configure_teams_webhook"
  fi
fi

# ------------------------------------------------------------
# 6. Install cron jobs
# ------------------------------------------------------------
echo ""
echo "▶ Installing cron jobs..."

HYGIENE_CMD="$NODE_PATH $SCRIPT_DIR/scripts/hygiene-sweep.mjs >> $HOME/.sc-engagement-hygiene.log 2>&1"
HYGIENE_CRON="30 9 * * 1 $HYGIENE_CMD"

MEETING_CMD="$NODE_PATH $SCRIPT_DIR/scripts/post-meeting-sweep.mjs >> $HOME/.sc-engagement-meetings.log 2>&1"
MEETING_CRON="0 14 * * 5 $MEETING_CMD"

CURRENT_CRON=$(crontab -l 2>/dev/null)

UPDATED_CRON="$CURRENT_CRON"
if echo "$CURRENT_CRON" | grep -q "hygiene-sweep"; then
  echo "   ✅ Hygiene cron already installed (Monday 9:30am)"
else
  UPDATED_CRON="$UPDATED_CRON
$HYGIENE_CRON"
  echo "   ✅ Hygiene cron installed (Monday 9:30am)"
fi

if echo "$CURRENT_CRON" | grep -q "post-meeting-sweep"; then
  echo "   ✅ Meeting review cron already installed (Friday 2:00pm)"
else
  UPDATED_CRON="$UPDATED_CRON
$MEETING_CRON"
  echo "   ✅ Meeting review cron installed (Friday 2:00pm)"
fi

echo "$UPDATED_CRON" | crontab -

# ------------------------------------------------------------
# Done
# ------------------------------------------------------------
echo ""
echo "=================================================="
echo "  ✅ Setup complete!"
echo "=================================================="
echo ""
echo "Next steps:"
echo "  1. Double-click ChromeLink.app on your Desktop"
echo "  2. Log into Dynamics, Outlook and Teams in that window"
echo "  3. Restart Claude Desktop"
echo "  4. Ask Claude anything — opportunities, calendar, hygiene sweep!"
echo ""
echo "Automated jobs:"
echo "  • Monday 9:30am — CRM hygiene sweep"
echo "  • Friday 2:00pm — Weekly meeting review"
if [ -n "$EXISTING_WEBHOOK" ] || [ -n "$NEW_WEBHOOK" ]; then
  echo "Results will be posted to your Teams channel."
else
  echo "Run setup again to add a Teams webhook for automated notifications."
fi
echo ""
