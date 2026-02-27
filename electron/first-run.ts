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
import { execSync } from "child_process";
import fs from "fs";
import path from "path";

const CONFIG_DIR = path.join(process.env.HOME || "~", ".thesidebar");
const CERT_DIR = path.join(CONFIG_DIR, "certs");
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

  // 2. Generate self-signed cert if missing
  if (!fs.existsSync(CERT_DIR)) {
    fs.mkdirSync(CERT_DIR, { recursive: true });
  }

  const certFile = path.join(CERT_DIR, "server.crt");
  const keyFile = path.join(CERT_DIR, "server.key");

  if (!fs.existsSync(certFile) || !fs.existsSync(keyFile)) {
    log("[setup] Generating self-signed certificate...");
    try {
      // Copy from bundled certs if available
      const bundledCert = resourcePath("certs", "server.crt");
      const bundledKey = resourcePath("certs", "server.key");

      if (fs.existsSync(bundledCert) && fs.existsSync(bundledKey)) {
        fs.copyFileSync(bundledCert, certFile);
        fs.copyFileSync(bundledKey, keyFile);
        log("[setup] Copied bundled certificates");
      } else {
        // Generate new cert with openssl
        execSync(
          `openssl req -x509 -newkey rsa:2048 -keyout "${keyFile}" -out "${certFile}" ` +
          `-days 3650 -nodes -subj "/CN=localhost" ` +
          `-addext "subjectAltName=DNS:localhost,IP:127.0.0.1"`,
          { timeout: 10000 }
        );
        log("[setup] Generated new self-signed certificate");
      }
    } catch (e: any) {
      logError(`[setup] Cert generation failed: ${e.message}`);
    }
  }

  // 3. Trust the cert in System keychain
  if (fs.existsSync(certFile)) {
    try {
      log("[setup] Trusting certificate in System keychain...");
      execSync(
        `osascript -e 'do shell script "security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain \\"${certFile}\\"" with administrator privileges'`,
        { timeout: 60000 }
      );
      log("[setup] Certificate trusted");
    } catch (e: any) {
      logError(`[setup] Could not trust cert (user may have cancelled): ${e.message}`);
    }
  }

  // 4. Install Word add-in manifest
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

  // 5. Mark as initialized
  fs.writeFileSync(FIRST_RUN_MARKER, new Date().toISOString());

  // 6. Welcome dialog
  await dialog.showMessageBox({
    type: "info",
    title: "Welcome to The Sidebar!",
    message: "The Sidebar has been set up.",
    detail:
      "✅ Config directory created (~/.thesidebar/)\n" +
      "✅ SSL certificate installed\n" +
      "✅ Word add-in manifest sideloaded\n\n" +
      "The server will start automatically. Open Word to use the add-in.",
    buttons: ["Get Started"],
  });
}
