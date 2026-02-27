#!/bin/bash
set -e
cd "$(dirname "$0")"

echo "⚖️  Rebuilding The Sidebar..."

# Kill running app
pkill -f "The Sidebar" 2>/dev/null && sleep 1 || true

# Build everything
echo "→ Building server..."
cd server && npm run build 2>&1 | tail -1
cd ..
echo "→ Building app..."
cd app && npm run build 2>&1 | tail -1
cd ..
echo "→ Building electron..."
cd electron && npx tsc 2>&1
cd ..

# Build DMG (arm64 only)
echo "→ Packaging DMG..."
rm -rf dist
npx electron-builder --mac --arm64 2>&1 | grep -E "building|completed|error" || true

# Swap app
echo "→ Installing..."
trash "/Applications/The Sidebar.app" 2>/dev/null || rm -rf "/Applications/The Sidebar.app"
DMG=$(ls dist/*arm64.dmg 2>/dev/null | head -1)
hdiutil attach "$DMG" -nobrowse -quiet
VOL=$(ls -d /Volumes/The\ Sidebar* 2>/dev/null | head -1)
cp -R "$VOL/The Sidebar.app" /Applications/
hdiutil detach "$VOL" -quiet

# Reinstall manifest
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp app/manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/thesidebar.xml 2>/dev/null || true

# Launch
echo "→ Launching..."
open "/Applications/The Sidebar.app"

echo "✅ Done! Restart Word to pick up changes."
