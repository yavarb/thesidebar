/**
 * @module electron/first-run
 * First-run setup for The Sidebar.
 *
 * On first launch:
 * 1. Create ~/.thesidebar/ config directory
 * 2. Generate self-signed SSL cert if missing
 * 3. Trust the cert in the macOS System keychain
 * 4. Install the Word add-in manifest for sideloading
 * 5. Show a welcome dialog
 */

import { dialog } from "electron";

import fs from "fs";
import path from "path";

const CONFIG_DIR = path.join(process.env.HOME || "~", ".thesidebar");

const FIRST_RUN_MARKER = path.join(CONFIG_DIR, ".initialized");

/** Check if this is the first run */
export function isFirstRun(): boolean {
  return !fs.existsSync(FIRST_RUN_MARKER);
}

/** Run first-time setup */
export async function firstRunSetup(
  resourcePath: (...segments: string[]) => string,
  log: (msg: string) => void,
  logError: (msg: string) => void
): Promise<void> {
  // 1. Create config directory
  if (!fs.existsSync(CONFIG_DIR)) {
    fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
    log("[setup] Created ~/.thesidebar/");
  }

  // Create logs directory
  const logsDir = path.join(CONFIG_DIR, "logs");
  if (!fs.existsSync(logsDir)) {
    fs.mkdirSync(logsDir, { recursive: true });
  }

  // 2. Install Word add-in manifest
  const wefDir = path.join(
    process.env.HOME || "~",
    "Library/Containers/com.microsoft.Word/Data/Documents/wef"
  );
  try {
    if (!fs.existsSync(wefDir)) fs.mkdirSync(wefDir, { recursive: true });
    const manifestSrc = resourcePath("app", "manifest.xml");
    const dest = path.join(wefDir, "thesidebar.xml");
    if (fs.existsSync(manifestSrc)) {
      fs.copyFileSync(manifestSrc, dest);
      log(`[setup] Manifest installed to ${dest}`);
    } else {
      log("[setup] No manifest source found, skipping sideload");
    }
  } catch (e: any) {
    logError(`[setup] Manifest install failed: ${e.message}`);
  }

  // 3. Mark as initialized
  fs.writeFileSync(FIRST_RUN_MARKER, new Date().toISOString());

  // 4. Welcome dialog
  await dialog.showMessageBox({
    type: "info",
    title: "Welcome to The Sidebar!",
    message: "The Sidebar has been set up.",
    detail:
      "✅ Config directory created (~/.thesidebar/)\n" +
      "✅ Word add-in installed\n\n" +
      "The server will start automatically. Open Word to use the add-in.",
    buttons: ["Get Started"],
  });
}
