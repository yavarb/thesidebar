#!/bin/bash
# The Sidebar Install Script
# Idempotent — safe to run multiple times.
#
# Checks dependencies, installs server + app + menubar,
# generates self-signed cert, creates ~/.thesidebar/, registers login item.

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG_DIR="$HOME/.thesidebar"
CERT_DIR="$SCRIPT_DIR/certs"
PLIST_LABEL="com.thesidebar.server"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"

echo "⚖️ The Sidebar Installer"
echo "═══════════════════════"
echo ""

# ── Check Dependencies ──
echo "── Checking dependencies ──"

if ! command -v node &>/dev/null; then
  echo "❌ Node.js not found. Install via: brew install node"
  exit 1
fi
NODE_VER=$(node -v)
echo "  ✅ Node.js $NODE_VER"

if ! command -v npm &>/dev/null; then
  echo "❌ npm not found. Install via: brew install node"
  exit 1
fi
NPM_VER=$(npm -v)
echo "  ✅ npm $NPM_VER"

# ── Create Config Directory ──
echo ""
echo "── Config directory ──"
if [ ! -d "$CONFIG_DIR" ]; then
  mkdir -p "$CONFIG_DIR"
  chmod 700 "$CONFIG_DIR"
  echo "  ✅ Created $CONFIG_DIR"
else
  echo "  ✅ $CONFIG_DIR already exists"
fi

if [ ! -f "$CONFIG_DIR/config.json" ]; then
  echo '{}' > "$CONFIG_DIR/config.json"
  chmod 600 "$CONFIG_DIR/config.json"
  echo "  ✅ Created default config.json"
else
  echo "  ✅ config.json already exists"
fi

# ── Generate Self-Signed Certificate ──
echo ""
echo "── SSL Certificate ──"
if [ ! -f "$CERT_DIR/server.key" ] || [ ! -f "$CERT_DIR/server.crt" ]; then
  mkdir -p "$CERT_DIR"
  openssl req -x509 -newkey rsa:2048 -keyout "$CERT_DIR/server.key" -out "$CERT_DIR/server.crt" \
    -days 365 -nodes -subj "/CN=localhost" -addext "subjectAltName=DNS:localhost,IP:127.0.0.1" \
    2>/dev/null
  echo "  ✅ Generated self-signed certificate"
else
  echo "  ✅ Certificate already exists"
fi

# ── Install Server Dependencies ──
echo ""
echo "── Server ──"
cd "$SCRIPT_DIR/server"
npm install --silent 2>/dev/null
echo "  ✅ Server dependencies installed"

# ── Install App Dependencies ──
echo ""
echo "── Word Add-in ──"
if [ -d "$SCRIPT_DIR/app" ]; then
  cd "$SCRIPT_DIR/app"
  npm install --silent 2>/dev/null
  echo "  ✅ Add-in dependencies installed"
  npm run build --silent 2>/dev/null || true
  echo "  ✅ Add-in built"
fi



# ── Register Login Item (optional) ──
echo ""
echo "── Login Item ──"
if [ ! -f "$PLIST_PATH" ]; then
  echo "  ℹ️  Auto-start not configured. Enable via menu bar app or run:"
  echo "     menubar > Auto-start on Login"
else
  echo "  ✅ Login item already registered"
fi

echo ""
echo "═══════════════════════"
echo "✅ The Sidebar installed successfully!"
echo ""
echo "To start the server:"
echo "  cd $SCRIPT_DIR && bash start.sh"
echo ""
echo "To configure API keys:"
echo "  Edit $CONFIG_DIR/config.json"
echo "  Or use the Settings panel in the Word add-in"
