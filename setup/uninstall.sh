#!/bin/bash
# Uninstall Alfred.mcp — removes cron jobs, Claude Desktop entry, Alfred.app, and optionally config + profile

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CLAUDE_CONFIG="$HOME/Library/Application Support/Claude/claude_desktop_config.json"
CHROMELINK_APP="$HOME/Desktop/Alfred.app"
CONFIG_FILE="$HOME/.alfred-config.json"
CHROME_PROFILE="$HOME/.alfred-profile"

echo ""
echo "=================================================="
echo "  Alfred.mcp — Uninstaller"
echo "=================================================="
echo ""
echo "This will remove Alfred from your Mac."
echo ""
printf "Are you sure you want to uninstall? [y/N]: "
read -r CONFIRM
if [ "$CONFIRM" != "y" ] && [ "$CONFIRM" != "Y" ]; then
  echo "   Aborted."
  exit 0
fi

# ------------------------------------------------------------
# 1. Remove cron jobs
# ------------------------------------------------------------
echo ""
echo "▶ Removing cron jobs..."

CURRENT_CRON=$(crontab -l 2>/dev/null || true)
NEW_CRON=$(echo "$CURRENT_CRON" | grep -v "hygiene-sweep" | grep -v "post-meeting-sweep" || true)

if [ "$CURRENT_CRON" != "$NEW_CRON" ]; then
  echo "$NEW_CRON" | crontab -
  echo "   ✅ Cron jobs removed"
else
  echo "   ℹ️  No Alfred cron jobs found"
fi

# ------------------------------------------------------------
# 2. Remove from Claude Desktop config
# ------------------------------------------------------------
echo ""
echo "▶ Removing from Claude Desktop..."

if [ -f "$CLAUDE_CONFIG" ]; then
  python3 - <<PYEOF
import json, os
path = """$CLAUDE_CONFIG"""
try:
    with open(path) as f:
        config = json.load(f)
    removed = False
    for key in ("alfred", "sc-engagement"):
        if key in config.get("mcpServers", {}):
            del config["mcpServers"][key]
            removed = True
    with open(path, "w") as f:
        json.dump(config, f, indent=2)
    print("   ✅ Removed from Claude Desktop config" if removed else "   ℹ️  No Alfred entry found in Claude Desktop config")
except Exception as e:
    print(f"   ⚠️  Could not update Claude Desktop config: {e}")
PYEOF
else
  echo "   ℹ️  Claude Desktop config not found"
fi

# ------------------------------------------------------------
# 3. Remove Alfred.app from Desktop
# ------------------------------------------------------------
echo ""
echo "▶ Removing Alfred.app from Desktop..."

if [ -d "$CHROMELINK_APP" ]; then
  rm -rf "$CHROMELINK_APP"
  echo "   ✅ Alfred.app removed"
else
  echo "   ℹ️  Alfred.app not found on Desktop"
fi

# Also remove old .command shortcut if still present
[ -f "$HOME/Desktop/Alfred.command" ] && rm -f "$HOME/Desktop/Alfred.command" && echo "   ✅ Alfred.command removed"

# ------------------------------------------------------------
# 4. Remove config file (optional)
# ------------------------------------------------------------
echo ""
if [ -f "$CONFIG_FILE" ]; then
  printf "▶ Remove ~/.alfred-config.json (Dynamics URL, Teams webhook, role)? [y/N]: "
  read -r REMOVE_CONFIG
  if [ "$REMOVE_CONFIG" = "y" ] || [ "$REMOVE_CONFIG" = "Y" ]; then
    rm -f "$CONFIG_FILE"
    echo "   ✅ Config removed"
  else
    echo "   ⏭  Config kept"
  fi
fi

# ------------------------------------------------------------
# 5. Remove Chrome profile (optional)
# ------------------------------------------------------------
echo ""
if [ -d "$CHROME_PROFILE" ]; then
  printf "▶ Remove ~/.alfred-profile (Chrome session, cookies)? [y/N]: "
  read -r REMOVE_PROFILE
  if [ "$REMOVE_PROFILE" = "y" ] || [ "$REMOVE_PROFILE" = "Y" ]; then
    rm -rf "$CHROME_PROFILE"
    echo "   ✅ Chrome profile removed"
  else
    echo "   ⏭  Chrome profile kept (re-run setup anytime to use it again)"
  fi
fi

# ------------------------------------------------------------
# 6. Remove log files (optional)
# ------------------------------------------------------------
LOGS_EXIST=false
[ -f "$HOME/.alfred-hygiene.log" ] && LOGS_EXIST=true
[ -f "$HOME/.alfred-meetings.log" ] && LOGS_EXIST=true

if $LOGS_EXIST; then
  echo ""
  printf "▶ Remove log files (~/.alfred-hygiene.log, ~/.alfred-meetings.log)? [y/N]: "
  read -r REMOVE_LOGS
  if [ "$REMOVE_LOGS" = "y" ] || [ "$REMOVE_LOGS" = "Y" ]; then
    rm -f "$HOME/.alfred-hygiene.log" "$HOME/.alfred-meetings.log"
    echo "   ✅ Log files removed"
  else
    echo "   ⏭  Log files kept"
  fi
fi

# ------------------------------------------------------------
# Done
# ------------------------------------------------------------
echo ""
echo "=================================================="
echo "  ✅ Alfred uninstalled"
echo "=================================================="
echo ""
echo "The Alfred install folder ($SCRIPT_DIR) was NOT removed."
echo "Delete it manually if you no longer need it."
echo ""
echo "Restart Claude Desktop to fully apply the changes."
echo ""
