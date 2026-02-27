#!/bin/bash
# The Sidebar Uninstall Script
# Idempotent — safe to run multiple times.
#
# Stops server, removes login item, optionally removes ~/.thesidebar/

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG_DIR="$HOME/.thesidebar"
PLIST_LABEL="com.thesidebar.server"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"
SERVER_PORT=3001

echo "⚖️ The Sidebar Uninstaller"
echo "═════════════════════════"
echo ""

# ── Stop Server ──
echo "── Stopping server ──"
if lsof -ti:$SERVER_PORT &>/dev/null; then
  lsof -ti:$SERVER_PORT | xargs kill -SIGTERM 2>/dev/null || true
  sleep 1
  echo "  ✅ Server stopped"
else
  echo "  ✅ Server not running"
fi

# ── Remove Login Item ──
echo ""
echo "── Login item ──"
if [ -f "$PLIST_PATH" ]; then
  launchctl unload "$PLIST_PATH" 2>/dev/null || true
  rm -f "$PLIST_PATH"
  echo "  ✅ Login item removed"
else
  echo "  ✅ No login item found"
fi

# ── Config Directory ──
echo ""
echo "── Config directory ──"
if [ -d "$CONFIG_DIR" ]; then
  read -p "  Remove $CONFIG_DIR? (contains API keys) [y/N] " -n 1 -r
  echo
  if [[ $REPLY =~ ^[Yy]$ ]]; then
    rm -rf "$CONFIG_DIR"
    echo "  ✅ Config directory removed"
  else
    echo "  ⏭️  Kept $CONFIG_DIR"
  fi
else
  echo "  ✅ No config directory found"
fi

echo ""
echo "═════════════════════════"
echo "✅ The Sidebar uninstalled"
echo ""
echo "Note: Source files in $SCRIPT_DIR were not removed."
echo "To fully remove, delete the directory manually."
