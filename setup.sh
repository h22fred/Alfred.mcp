#!/bin/bash
set -e

# ============================================================
# SC Engagement MCP — Setup Script
# ============================================================
# Run this once to install everything and configure Claude Desktop.
# Requirements: macOS, Google Chrome, Claude Desktop

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CLAUDE_CONFIG="$HOME/Library/Application Support/Claude/claude_desktop_config.json"
CHROMELINK_APP="$HOME/Desktop/Alfred.app"

echo ""
echo "=================================================="
echo "  SC Engagement MCP — Setup"
echo "=================================================="
echo ""

# ------------------------------------------------------------
# 1. Check / install Node.js
# ------------------------------------------------------------
echo "▶ Checking Node.js..."
NODE_PATH=""
for p in /opt/homebrew/bin/node /usr/local/bin/node "$HOME/.nvm/versions/node/$(ls "$HOME/.nvm/versions/node/" 2>/dev/null | sort -V | tail -1)/bin/node"; do
  if [ -x "$p" ]; then NODE_PATH="$p"; break; fi
done

# Also check if nvm is already loaded
if [ -z "$NODE_PATH" ] && command -v node &>/dev/null; then
  NODE_PATH="$(command -v node)"
fi

if [ -z "$NODE_PATH" ]; then
  echo "   ⚠️  Node.js not found — installing via nvm (no password required)..."

  # Install nvm — runs entirely in the user's home directory, no sudo needed
  export NVM_DIR="$HOME/.nvm"
  curl -fsSL https://raw.githubusercontent.com/nvm-sh/nvm/v0.40.1/install.sh | bash

  # Load nvm for the rest of this script
  [ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh"

  echo "   📦 Installing Node.js LTS..."
  nvm install --lts
  nvm use --lts

  NODE_PATH="$(command -v node 2>/dev/null)"

  if [ -z "$NODE_PATH" ]; then
    echo ""
    echo "❌ Node.js installation failed. Please install manually from https://nodejs.org"
    exit 1
  fi
  echo "   ✅ Node.js installed via nvm"
fi

NODE_DIR="$(dirname "$NODE_PATH")"
echo "   ✅ Node.js found at $NODE_PATH"

# ------------------------------------------------------------
# 2. Install dependencies and build
# ------------------------------------------------------------
echo ""
echo "▶ Installing dependencies..."
PATH="$NODE_DIR:$PATH" npm ci --prefix "$SCRIPT_DIR" --no-fund

echo ""
echo "▶ Building MCP server..."
PATH="$NODE_DIR:$PATH" npm run build --prefix "$SCRIPT_DIR"
echo "   ✅ Build complete"

# Save installed version SHA for update checks
INSTALLED_SHA=$(git -C "$SCRIPT_DIR" rev-parse --short HEAD 2>/dev/null || echo "")
if [ -n "$INSTALLED_SHA" ]; then
  CONFIG_FILE_EARLY="$HOME/.alfred-config.json"
  python3 -c "
import json, os
f = os.path.expanduser('~/.alfred-config.json')
d = json.load(open(f)) if os.path.exists(f) else {}
d['installedVersion'] = '$INSTALLED_SHA'
json.dump(d, open(f, 'w'), indent=2)
" 2>/dev/null
  chmod 600 "$HOME/.alfred-config.json" 2>/dev/null || true
  echo "   ✅ Installed version: $INSTALLED_SHA"
fi

# ------------------------------------------------------------
# 3. Configure Claude Desktop
# ------------------------------------------------------------
echo ""
echo "▶ Configuring Claude Desktop..."

DIST_PATH="$SCRIPT_DIR/dist/${ALFRED_VARIANT:-sc}/index.js"
MCP_ENTRY=$(cat <<EOF
{
  "command": "$NODE_PATH",
  "args": ["$DIST_PATH"]
}
EOF
)

if [ ! -f "$CLAUDE_CONFIG" ]; then
  mkdir -p "$(dirname "$CLAUDE_CONFIG")"
  echo "{\"mcpServers\":{\"alfred\":$MCP_ENTRY}}" > "$CLAUDE_CONFIG"
  echo "   ✅ Created Claude Desktop config"
else
  python3 - <<PYEOF
import json

config_path = """$CLAUDE_CONFIG"""
dist_path = """$DIST_PATH"""
node_path = """$NODE_PATH"""

with open(config_path, "r") as f:
    config = json.load(f)

if "mcpServers" not in config:
    config["mcpServers"] = {}

# Remove old entry if present
config["mcpServers"].pop("sc-engagement", None)

config["mcpServers"]["alfred"] = {
    "command": node_path,
    "args": [dist_path]
}

with open(config_path, "w") as f:
    json.dump(config, f, indent=2)

print("   ✅ Claude Desktop config updated")
PYEOF
fi

# ------------------------------------------------------------
# 4. Create Alfred.app on Desktop (plain shell bundle — no AppleScript)
# ------------------------------------------------------------
echo ""
echo "▶ Creating Alfred.app on Desktop..."

# Remove old versions
[ -d "$CHROMELINK_APP" ] && rm -rf "$CHROMELINK_APP"
[ -f "$HOME/Desktop/Alfred.command" ] && rm -f "$HOME/Desktop/Alfred.command"

mkdir -p "$CHROMELINK_APP/Contents/MacOS"
mkdir -p "$CHROMELINK_APP/Contents/Resources"

cat > "$CHROMELINK_APP/Contents/MacOS/Alfred" << SHELLEOF
#!/bin/bash
notify() { osascript -e "display notification \"\$1\" with title \"Alfred\"" 2>/dev/null; }

# Already running?
if pgrep -f "alfred-profile" > /dev/null 2>&1; then
  notify "Already running — you're good to use Claude!"
  open -a "Claude" 2>/dev/null || true
  exit 0
fi

mkdir -p "\$HOME/.alfred-profile"
open -na "Google Chrome" --args \
  --remote-debugging-port=9222 \
  --user-data-dir="\$HOME/.alfred-profile" \
  --no-first-run \
  --no-default-browser-check \
  --disable-extensions \
  --disable-sync \
  --disable-default-apps \
  --disable-translate \
  --disable-component-update \
  --disable-domain-reliability \
  --disable-client-side-phishing-detection \
  "$NEW_DYNAMICS_URL" \
  "https://outlook.office.com" \
  "https://teams.microsoft.com/v2/"

# First run detection — profile dir will be nearly empty on first launch
PROFILE_SIZE=\$(du -sk "\$HOME/.alfred-profile" 2>/dev/null | cut -f1)
if [ -z "\$PROFILE_SIZE" ] || [ "\$PROFILE_SIZE" -lt 500 ]; then
  notify "First time setup: log into Dynamics, Outlook and Teams in this window. You only do this once!"
else
  notify "Launched — ready for Claude!"
fi
open -a "Claude" 2>/dev/null || true

# Background update check — runs silently, never blocks startup
(
  INSTALLED=\$(python3 -c "
import json, os
f = os.path.expanduser('~/.alfred-config.json')
d = json.load(open(f)) if os.path.exists(f) else {}
print(d.get('installedVersion', ''))
" 2>/dev/null)
  if [ -z "\$INSTALLED" ]; then exit 0; fi
  LATEST=\$(curl -sf --max-time 5 \
    "https://api.github.com/repos/h22fred/Alfred.mcp/commits/main" \
    | python3 -c "import json,sys; print(json.load(sys.stdin)['sha'][:7])" 2>/dev/null)
  if [ -n "\$LATEST" ] && [ "\$INSTALLED" != "\$LATEST" ]; then
    osascript -e "display notification \"A new version of Alfred is available. Ask Claude: update Alfred\" with title \"Alfred Update Available 🆕\" sound name \"Ping\"" 2>/dev/null
  fi
) &
SHELLEOF

chmod +x "$CHROMELINK_APP/Contents/MacOS/Alfred"

# Copy icon
cp "$SCRIPT_DIR/assets/alfred.icns" "$CHROMELINK_APP/Contents/Resources/alfred.icns"

cat > "$CHROMELINK_APP/Contents/Info.plist" << 'PLISTEOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleExecutable</key><string>Alfred</string>
  <key>CFBundleIdentifier</key><string>com.servicenow.alfred</string>
  <key>CFBundleName</key><string>Alfred</string>
  <key>CFBundleIconFile</key><string>alfred</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleVersion</key><string>1.0</string>
  <key>LSUIElement</key><true/>
</dict>
</plist>
PLISTEOF

echo "   ✅ Alfred.app created on Desktop"
echo "   ℹ️  First launch: right-click → Open (one-time macOS approval)"

# ------------------------------------------------------------
# 5. Dynamics URL
# ------------------------------------------------------------
echo ""
echo "▶ Dynamics 365 instance..."

EXISTING_DYNAMICS_URL=$(python3 -c "import json,os; d=json.load(open(os.path.expanduser('~/.alfred-config.json'))) if os.path.exists(os.path.expanduser('~/.alfred-config.json')) else {}; print(d.get('dynamicsUrl',''))" 2>/dev/null)

if [ -n "$EXISTING_DYNAMICS_URL" ]; then
  echo "   ✅ Dynamics URL already set: $EXISTING_DYNAMICS_URL"
  printf "   Change company name? (press Enter to keep): "
  read -r NEW_COMPANY </dev/tty
  if [ -n "$NEW_COMPANY" ]; then
    NEW_DYNAMICS_URL="https://${NEW_COMPANY}.crm.dynamics.com"
  else
    NEW_DYNAMICS_URL="$EXISTING_DYNAMICS_URL"
  fi
else
  printf "   What is your company name? (press Enter for 'servicenow'): "
  read -r NEW_COMPANY </dev/tty
  [ -z "$NEW_COMPANY" ] && NEW_COMPANY="servicenow"
  NEW_DYNAMICS_URL="https://${NEW_COMPANY}.crm.dynamics.com"
fi

python3 -c "
import json, os
f = os.path.expanduser('~/.alfred-config.json')
d = json.load(open(f)) if os.path.exists(f) else {}
d['dynamicsUrl'] = '$NEW_DYNAMICS_URL'
json.dump(d, open(f, 'w'), indent=2)
"
chmod 600 "$HOME/.alfred-config.json"
echo "   ✅ Dynamics URL set to: $NEW_DYNAMICS_URL"

# ------------------------------------------------------------
# 5b. Teams webhook config
# ------------------------------------------------------------
echo ""
echo "▶ Setting up Teams webhook for hygiene sweep notifications..."
CONFIG_FILE="$HOME/.alfred-config.json"
EXISTING_WEBHOOK=""
if [ -f "$CONFIG_FILE" ]; then
  EXISTING_WEBHOOK=$(python3 -c "import json; d=json.load(open('$CONFIG_FILE')); print(d.get('teamsWebhook',''))" 2>/dev/null)
fi

if [ -n "$EXISTING_WEBHOOK" ]; then
  echo "   ✅ Teams webhook already configured"
  echo "      $EXISTING_WEBHOOK"
  echo ""
  printf "   Replace it? (press Enter to keep, or paste a new URL): "
  read -r NEW_WEBHOOK </dev/tty
  if [ -n "$NEW_WEBHOOK" ]; then
    EXISTING_WEBHOOK="$NEW_WEBHOOK"
    python3 -c "import json; f='$CONFIG_FILE'; d=json.load(open(f)) if __import__('os').path.exists(f) else {}; d['teamsWebhook']='$NEW_WEBHOOK'; json.dump(d,open(f,'w'),indent=2)"
    chmod 600 "$CONFIG_FILE"
    echo "   ✅ Webhook updated"
  fi
else
  echo ""
  echo "   To get a webhook URL:"
  echo "   Teams → any channel → ··· → Connectors → Incoming Webhook → configure → copy URL"
  echo ""
  printf "   Paste your Teams incoming webhook URL (or press Enter to skip): "
  read -r NEW_WEBHOOK </dev/tty
  if [ -n "$NEW_WEBHOOK" ]; then
    python3 -c "import json; f='$CONFIG_FILE'; d=json.load(open(f)) if __import__('os').path.exists(f) else {}; d['teamsWebhook']='$NEW_WEBHOOK'; json.dump(d,open(f,'w'),indent=2)"
    chmod 600 "$CONFIG_FILE"
    echo "   ✅ Webhook saved to $CONFIG_FILE"
  else
    echo "   ⏭  Skipped — hygiene sweep will run without Teams notifications"
    echo "      Run setup again anytime to add it, or ask Claude to configure_teams_webhook"
  fi
fi

# ------------------------------------------------------------
# 6. Role
# ------------------------------------------------------------
echo ""
if [ "$ALFRED_VARIANT" = "sales" ]; then
  echo "▶ What is your Sales role?"
  echo ""
  echo "   1) AE      — Account Executive (you own accounts and opportunities)"
  echo "   2) Manager — Sales Manager (you oversee a team of AEs)"
  echo ""
  printf "   Enter 1 or 2 (default: 1): "
  read -r ROLE_CHOICE </dev/tty
  case "$ROLE_CHOICE" in
    2) USER_ROLE="sales_manager"; echo "   ✅ Role set to Sales Manager — Alfred will show territory-wide pipeline" ;;
    *) USER_ROLE="sales";         echo "   ✅ Role set to AE — Alfred will default to your own pipeline" ;;
  esac
else
  echo "▶ What is your SC role?"
  echo ""
  echo "   1) SC      — Solution Consultant (you have assigned opportunities in Dynamics)"
  echo "   2) SSC     — Sales Support Consultant (you support SCs, no assigned pipeline)"
  echo "   3) Manager — SC Manager (you want to see your team's pipeline)"
  echo ""
  printf "   Enter 1, 2 or 3 (default: 1): "
  read -r ROLE_CHOICE </dev/tty
  case "$ROLE_CHOICE" in
    2) USER_ROLE="ssc";     echo "   ✅ Role set to SSC — Alfred will search all accounts by default" ;;
    3) USER_ROLE="manager"; echo "   ✅ Role set to Manager — Alfred will browse by territory/SC by default" ;;
    *) USER_ROLE="sc";      echo "   ✅ Role set to SC — Alfred will default to your own pipeline" ;;
  esac
fi
python3 -c "
import json, os
f = os.path.expanduser('~/.alfred-config.json')
d = json.load(open(f)) if os.path.exists(f) else {}
d['role'] = '$USER_ROLE'
json.dump(d, open(f, 'w'), indent=2)
"
chmod 600 "$HOME/.alfred-config.json"

# ------------------------------------------------------------
# 7. Engagement types
# ------------------------------------------------------------
echo ""
echo "▶ Which milestones do you track on your opportunities?"
echo "   (Press Enter to keep all defaults, or enter numbers separated by spaces)"
echo ""

if [ "$ALFRED_VARIANT" = "sales" ]; then
  # AE-owned milestones per NowSell AI Native Sales framework
  echo "    1) Discovery               4) Budget"
  echo "    2) Opportunity Summary     5) Implementation Plan"
  echo "    3) Mutual Plan             6) Stakeholder Alignment"
  echo ""
  printf "   Your selection (e.g. 1 2 3), or Enter for all: "
  read -r TYPE_SELECTION </dev/tty

  python3 - <<PYEOF
import json, os
all_types = [
  "Discovery", "Opportunity Summary", "Mutual Plan",
  "Budget", "Implementation Plan", "Stakeholder Alignment"
]
sel = "$TYPE_SELECTION".strip()
if sel:
    indices = [int(x)-1 for x in sel.split() if x.isdigit() and 1 <= int(x) <= len(all_types)]
    selected = [all_types[i] for i in indices] if indices else all_types
else:
    selected = all_types
f = os.path.expanduser('~/.alfred-config.json')
d = json.load(open(f)) if os.path.exists(f) else {}
d['engagementTypes'] = selected
json.dump(d, open(f, 'w'), indent=2)
print("   ✅ Milestones: " + ", ".join(selected))
PYEOF
else
  # SC/SSC/Manager-owned engagement types
  echo "    1) Business Case            6) Post Sale Engagement"
  echo "    2) Customer Business Review 7) POV"
  echo "    3) Demo                     8) RFx"
  echo "    4) Discovery                9) Technical Win"
  echo "    5) EBC                     10) Workshop"
  echo ""
  printf "   Your selection (e.g. 3 4 8 9), or Enter for all: "
  read -r TYPE_SELECTION </dev/tty

  python3 - <<PYEOF
import json, os
all_types = [
  "Business Case", "Customer Business Review", "Demo", "Discovery", "EBC",
  "Post Sale Engagement", "POV", "RFx", "Technical Win", "Workshop"
]
sel = "$TYPE_SELECTION".strip()
if sel:
    indices = [int(x)-1 for x in sel.split() if x.isdigit() and 1 <= int(x) <= len(all_types)]
    selected = [all_types[i] for i in indices] if indices else all_types
else:
    selected = all_types
f = os.path.expanduser('~/.alfred-config.json')
d = json.load(open(f)) if os.path.exists(f) else {}
d['engagementTypes'] = selected
json.dump(d, open(f, 'w'), indent=2)
print("   ✅ Engagement types: " + ", ".join(selected))
PYEOF
fi
chmod 600 "$HOME/.alfred-config.json"

# ------------------------------------------------------------
# 8. Install cron jobs
# ------------------------------------------------------------
echo ""
echo "▶ Automated jobs..."
echo ""

# Helper: parse day name to cron weekday number (1=Mon … 5=Fri, 0=Sun)
day_to_cron() {
  case "$(echo "$1" | tr '[:upper:]' '[:lower:]')" in
    mon|monday)    echo 1 ;;
    tue|tuesday)   echo 2 ;;
    wed|wednesday) echo 3 ;;
    thu|thursday)  echo 4 ;;
    fri|friday)    echo 5 ;;
    sat|saturday)  echo 6 ;;
    sun|sunday)    echo 0 ;;
    *)             echo ""  ;;
  esac
}

# Helper: parse HH:MM into separate hour/minute, stripping leading zeros
parse_time() {
  TIME_INPUT="$1"
  T_HOUR=$(echo "$TIME_INPUT" | cut -d: -f1 | sed 's/^0*//')
  T_MIN=$(echo "$TIME_INPUT"  | cut -d: -f2 | sed 's/^0*//')
  T_HOUR="${T_HOUR:-0}"
  T_MIN="${T_MIN:-0}"
}

# --- Hygiene sweep ---
printf "   Install hygiene sweep (flags missing engagements on your pipeline)? [Y/n]: "
read -r INSTALL_HYGIENE </dev/tty

HYGIENE_CRON=""
HYGIENE_SCHEDULE_DESC=""

case "$INSTALL_HYGIENE" in
  [nN]*) echo "   ⏭  Hygiene sweep skipped" ;;
  *)
    printf "   Run on which day?       [Monday]: "
    read -r HYGIENE_DAY </dev/tty
    HYGIENE_DAY="${HYGIENE_DAY:-Monday}"
    HYGIENE_CRON_DAY=$(day_to_cron "$HYGIENE_DAY")
    while [ -z "$HYGIENE_CRON_DAY" ]; do
      printf "   Unknown day — try again [Monday]: "
      read -r HYGIENE_DAY </dev/tty
      HYGIENE_DAY="${HYGIENE_DAY:-Monday}"
      HYGIENE_CRON_DAY=$(day_to_cron "$HYGIENE_DAY")
    done

    printf "   Run at what time? (HH:MM 24h) [09:30]: "
    read -r HYGIENE_TIME </dev/tty
    HYGIENE_TIME="${HYGIENE_TIME:-09:30}"
    parse_time "$HYGIENE_TIME"
    HYGIENE_HOUR=$T_HOUR; HYGIENE_MIN=$T_MIN

    HYGIENE_SCHEDULE_DESC="$HYGIENE_DAY at $HYGIENE_TIME"
    ROTATE_LOG="f=\$HOME/.alfred-hygiene.log; [ -f \"\$f\" ] && [ \$(wc -c < \"\$f\") -gt 1048576 ] && tail -500 \"\$f\" > \"\$f.tmp\" && mv \"\$f.tmp\" \"\$f\""
    HYGIENE_CMD="$NODE_PATH $SCRIPT_DIR/scripts/hygiene-sweep.mjs >> $HOME/.alfred-hygiene.log 2>&1"
    HYGIENE_CRON="$HYGIENE_MIN $HYGIENE_HOUR * * $HYGIENE_CRON_DAY $ROTATE_LOG; $HYGIENE_CMD"
    touch "$HOME/.alfred-hygiene.log"
    chmod 600 "$HOME/.alfred-hygiene.log"
    ;;
esac

# --- Meeting review ---
printf "   Install meeting review (matches this week's meetings to open opps)? [Y/n]: "
read -r INSTALL_MEETING </dev/tty

MEETING_CRON=""
MEETING_SCHEDULE_DESC=""

case "$INSTALL_MEETING" in
  [nN]*) echo "   ⏭  Meeting review skipped" ;;
  *)
    printf "   Run on which day?       [Friday]: "
    read -r MEETING_DAY </dev/tty
    MEETING_DAY="${MEETING_DAY:-Friday}"
    MEETING_CRON_DAY=$(day_to_cron "$MEETING_DAY")
    while [ -z "$MEETING_CRON_DAY" ]; do
      printf "   Unknown day — try again [Friday]: "
      read -r MEETING_DAY </dev/tty
      MEETING_DAY="${MEETING_DAY:-Friday}"
      MEETING_CRON_DAY=$(day_to_cron "$MEETING_DAY")
    done

    printf "   Run at what time? (HH:MM 24h) [14:00]: "
    read -r MEETING_TIME </dev/tty
    MEETING_TIME="${MEETING_TIME:-14:00}"
    parse_time "$MEETING_TIME"
    MEETING_HOUR=$T_HOUR; MEETING_MIN=$T_MIN

    MEETING_SCHEDULE_DESC="$MEETING_DAY at $MEETING_TIME"
    ROTATE_LOG2="f=\$HOME/.alfred-meetings.log; [ -f \"\$f\" ] && [ \$(wc -c < \"\$f\") -gt 1048576 ] && tail -500 \"\$f\" > \"\$f.tmp\" && mv \"\$f.tmp\" \"\$f\""
    MEETING_CMD="$NODE_PATH $SCRIPT_DIR/scripts/post-meeting-sweep.mjs >> $HOME/.alfred-meetings.log 2>&1"
    MEETING_CRON="$MEETING_MIN $MEETING_HOUR * * $MEETING_CRON_DAY $ROTATE_LOG2; $MEETING_CMD"
    touch "$HOME/.alfred-meetings.log"
    chmod 600 "$HOME/.alfred-meetings.log"
    ;;
esac

# Apply crontab — remove stale entries then add new ones
CURRENT_CRON=$(crontab -l 2>/dev/null)
UPDATED_CRON="$CURRENT_CRON"

if [ -n "$HYGIENE_CRON" ]; then
  UPDATED_CRON=$(echo "$UPDATED_CRON" | grep -v "hygiene-sweep")
  UPDATED_CRON="$UPDATED_CRON
$HYGIENE_CRON"
  echo "   ✅ Hygiene sweep scheduled: $HYGIENE_SCHEDULE_DESC"
fi

if [ -n "$MEETING_CRON" ]; then
  UPDATED_CRON=$(echo "$UPDATED_CRON" | grep -v "post-meeting-sweep")
  UPDATED_CRON="$UPDATED_CRON
$MEETING_CRON"
  echo "   ✅ Meeting review scheduled: $MEETING_SCHEDULE_DESC"
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
echo "  1. Double-click Alfred.app on your Desktop"
echo "  2. Log into Dynamics, Outlook and Teams in that window"
echo "  3. Restart Claude Desktop"
echo "  4. Ask Claude anything — opportunities, calendar, hygiene sweep!"
echo ""
if [ -n "$HYGIENE_SCHEDULE_DESC" ] || [ -n "$MEETING_SCHEDULE_DESC" ]; then
  echo "Automated jobs:"
  [ -n "$HYGIENE_SCHEDULE_DESC" ] && echo "  • $HYGIENE_SCHEDULE_DESC — CRM hygiene sweep"
  [ -n "$MEETING_SCHEDULE_DESC"  ] && echo "  • $MEETING_SCHEDULE_DESC — Weekly meeting review"
  if [ -n "$EXISTING_WEBHOOK" ] || [ -n "$NEW_WEBHOOK" ]; then
    echo "Results will be posted to your Teams channel."
  else
    echo "Run setup again to add a Teams webhook for automated notifications."
  fi
  echo ""
fi
