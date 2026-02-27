/**
 * @module electron/main
 * The Sidebar Electron main process.
 *
 * Responsibilities:
 * - Owns the server lifecycle (fork server/dist/index.js as child process)
 * - macOS menu bar tray with status, controls, and settings
 * - First-run setup (certs, manifest sideloading)
 * - Auto-update via electron-updater
 * - Logging to ~/.thesidebar/logs/
 */

import {
  app,
  Tray,
  Menu,
  nativeImage,
  dialog,
  shell,
  Notification,
} from "electron";
import { fork, ChildProcess, execSync } from "child_process";
import path from "path";
import https from "https";
import fs from "fs";
import { autoUpdater, UpdateInfo } from "electron-updater";
import { firstRunSetup, isFirstRun } from "./first-run";
import { createLogger, Logger } from "./logger";

// ── Paths ──

/** Resolve a resource path — handles both dev and packaged app */
function resourcePath(...segments: string[]): string {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, ...segments);
  }
  return path.join(__dirname, "..", ...segments);
}

// ── Config ──
const SERVER_PORT = 3001;
const STATUS_POLL_INTERVAL = 5000;
const UPDATE_CHECK_INTERVAL = 6 * 60 * 60 * 1000; // 6 hours
const CONFIG_DIR = path.join(process.env.HOME || "~", ".thesidebar");

// ── State ──
let tray: Tray | null = null;
let serverProcess: ChildProcess | null = null;
let serverRunning = false;
let connectionCount = 0;
let promptQueueSize = 0;
let statusInterval: NodeJS.Timeout | null = null;
let updateInterval: NodeJS.Timeout | null = null;
let logger: Logger;
let pendingUpdate: UpdateInfo | null = null;

// ── Logging ──
function log(msg: string): void {
  console.log(msg);
  logger?.info(msg);
}

function logError(msg: string): void {
  console.error(msg);
  logger?.error(msg);
}

// ── HTTP Helper ──

/** Fetch JSON from the The Sidebar server */
function fetchStatus(): Promise<any> {
  return new Promise((resolve, reject) => {
    const req = https.get(
      `https://localhost:${SERVER_PORT}/api/status`,
      { rejectUnauthorized: false },
      (res) => {
        let body = "";
        res.on("data", (c: Buffer) => (body += c));
        res.on("end", () => {
          try { resolve(JSON.parse(body)); } catch { reject(new Error("Invalid JSON")); }
        });
      }
    );
    req.on("error", reject);
    req.setTimeout(3000, () => { req.destroy(); reject(new Error("Timeout")); });
  });
}

// ── Server Management ──

async function checkServerStatus(): Promise<void> {
  try {
    const res = await fetchStatus();
    if (res?.ok) {
      serverRunning = true;
      const data = res.data || {};
      connectionCount = data.connected ? 1 : 0;
      promptQueueSize = data.promptQueueSize || 0;
    } else {
      serverRunning = false;
    }
  } catch {
    serverRunning = false;
  }
  updateTrayMenu();
}

/** Start the The Sidebar server by forking the compiled JS */
function startServer(): void {
  if (serverProcess) return;

  const serverEntry = resourcePath("server", "dist", "index.js");
  if (!fs.existsSync(serverEntry)) {
    dialog.showErrorBox("The Sidebar", `Server not found: ${serverEntry}\nRun 'npm run build:server' first.`);
    return;
  }

  try {
    log(`[server] Forking ${serverEntry}`);
    serverProcess = fork(serverEntry, [], {
      cwd: resourcePath("server"),
      env: { ...process.env, SIDEBAR_PORT: String(SERVER_PORT), NODE_TLS_REJECT_UNAUTHORIZED: "0" },
      silent: true,
    });

    serverProcess.stdout?.on("data", (data: Buffer) => {
      const msg = data.toString().trim();
      if (msg) log(`[server] ${msg}`);
    });
    serverProcess.stderr?.on("data", (data: Buffer) => {
      const msg = data.toString().trim();
      if (msg) logError(`[server] ${msg}`);
    });
    serverProcess.on("exit", (code) => {
      log(`[server] Exited with code ${code}`);
      serverProcess = null;
      serverRunning = false;
      updateTrayMenu();
    });
    serverProcess.on("error", (err) => {
      logError(`[server] Error: ${err.message}`);
      serverProcess = null;
      serverRunning = false;
      updateTrayMenu();
    });

    setTimeout(() => checkServerStatus(), 2000);
  } catch (e: any) {
    logError(`[server] Failed to start: ${e.message}`);
    dialog.showErrorBox("The Sidebar", `Failed to start server: ${e.message}`);
  }
}

/** Stop the The Sidebar server gracefully (SIGTERM, then SIGKILL after 5s) */
function stopServer(): void {
  if (serverProcess) {
    log("[server] Sending SIGTERM...");
    serverProcess.kill("SIGTERM");
    const forceKill = setTimeout(() => {
      if (serverProcess) {
        log("[server] Force killing...");
        serverProcess.kill("SIGKILL");
        serverProcess = null;
      }
    }, 5000);
    serverProcess.once("exit", () => { clearTimeout(forceKill); serverProcess = null; });
  }
  try { execSync(`lsof -ti:${SERVER_PORT} | xargs kill -SIGTERM 2>/dev/null || true`, { timeout: 3000 }); } catch {}
  serverRunning = false;
  connectionCount = 0;
  updateTrayMenu();
}

// ── Auto-Updater ──

function setupAutoUpdater(): void {
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.logger = {
    info: (msg: any) => log(`[updater] ${msg}`),
    warn: (msg: any) => log(`[updater] WARN: ${msg}`),
    error: (msg: any) => logError(`[updater] ERROR: ${msg}`),
    debug: (msg: any) => {},
  };

  autoUpdater.on("update-available", (info: UpdateInfo) => {
    log(`[updater] Update available: v${info.version}`);
    pendingUpdate = info;
    updateTrayMenu();
    if (Notification.isSupported()) {
      new Notification({ title: "The Sidebar Update", body: `Version ${info.version} downloading...` }).show();
    }
  });

  autoUpdater.on("update-downloaded", (info: UpdateInfo) => {
    log(`[updater] Downloaded: v${info.version}`);
    pendingUpdate = info;
    updateTrayMenu();
  });

  autoUpdater.on("error", (err: Error) => {
    logError(`[updater] ${err.message}`);
  });

  autoUpdater.checkForUpdates().catch(() => {});
  updateInterval = setInterval(() => { autoUpdater.checkForUpdates().catch(() => {}); }, UPDATE_CHECK_INTERVAL);
}

// ── Add-in Sideloading ──

function installWordAddin(): void {
  const wefDir = path.join(process.env.HOME || "~", "Library/Containers/com.microsoft.Word/Data/Documents/wef");
  try {
    if (!fs.existsSync(wefDir)) fs.mkdirSync(wefDir, { recursive: true });

    const manifestSrc = resourcePath("app", "dist", "manifest.xml");
    const src = fs.existsSync(manifestSrc) ? manifestSrc : resourcePath("app", "manifest.xml");

    if (!fs.existsSync(src)) {
      logError("[addin] No manifest found, generating...");
      generateManifest(path.join(wefDir, "thesidebar.xml"));
      return;
    }
    const dest = path.join(wefDir, "thesidebar.xml");
    fs.copyFileSync(src, dest);
    log(`[addin] Manifest installed to ${dest}`);
  } catch (e: any) {
    logError(`[addin] Failed: ${e.message}`);
  }
}

function generateManifest(destPath: string): void {
  const manifest = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>05c2e1c9-3e1d-406e-9a91-e9ac64854143</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>The Sidebar</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="The Sidebar"/>
  <Description DefaultValue="AI-powered Word assistant"/>
  <IconUrl DefaultValue="https://localhost:3001/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3001/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://github.com/yavarb/thesidebar"/>
  <AppDomains><AppDomain>https://localhost:3001</AppDomain></AppDomains>
  <Hosts><Host Name="Document"/></Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3001/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3001/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3001/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3001/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://github.com/yavarb/thesidebar"/>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3001/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3001/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Welcome to The Sidebar!"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="The Sidebar"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="The Sidebar"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="The Sidebar is ready. Click the The Sidebar button to open the task pane."/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open The Sidebar task pane"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;
  fs.writeFileSync(destPath, manifest);
  log(`[addin] Generated manifest at ${destPath}`);
}

// ── Tray ──

function updateTrayMenu(): void {
  if (!tray) return;

  const version = app.getVersion();
  const statusLabel = serverRunning
    ? `🟢 Running (${connectionCount} conn${connectionCount !== 1 ? "s" : ""})`
    : "🔴 Stopped";

  const items: Electron.MenuItemConstructorOptions[] = [
    { label: `⚖️ The Sidebar v${version}`, enabled: false },
    { type: "separator" },
    { label: statusLabel, enabled: false },
  ];

  if (serverRunning && promptQueueSize > 0) {
    items.push({ label: `  Queue: ${promptQueueSize}`, enabled: false });
  }

  items.push(
    { type: "separator" },
    { label: serverRunning ? "Stop Server" : "Start Server", click: () => serverRunning ? stopServer() : startServer() },
    { label: "Open in Browser", enabled: serverRunning, click: () => shell.openExternal(`https://localhost:${SERVER_PORT}/api/status`) },
    { type: "separator" },
    { label: "Reinstall Word Add-in", click: () => { installWordAddin(); dialog.showMessageBox({ type: "info", title: "The Sidebar", message: "Add-in manifest reinstalled. Restart Word." }); } },
    { label: "Open Config Directory", click: () => shell.openPath(CONFIG_DIR) },
  );

  if (pendingUpdate) {
    items.push({ label: `⬆ Update v${pendingUpdate.version} — click to install`, click: () => autoUpdater.quitAndInstall(false, true) });
  } else {
    items.push({ label: "Check for Updates", click: () => autoUpdater.checkForUpdates().catch((e: Error) => {
      dialog.showMessageBox({ type: "info", title: "The Sidebar", message: "No updates available." });
    })});
  }

  items.push(
    { type: "separator" },
    { label: "About The Sidebar", click: () => dialog.showMessageBox({ type: "info", title: "About The Sidebar", message: `The Sidebar v${version}`, detail: "AI-powered Word assistant.\nhttps://github.com/yavarb/thesidebar" }) },
    { label: "Quit", click: () => { stopServer(); if (statusInterval) clearInterval(statusInterval); if (updateInterval) clearInterval(updateInterval); app.quit(); } },
  );

  tray.setContextMenu(Menu.buildFromTemplate(items));
  tray.setTitle(serverRunning ? "⚖️" : "");
  tray.setToolTip(`The Sidebar — ${serverRunning ? "Running" : "Stopped"}`);
}

// ── App Lifecycle ──
app.dock?.hide();

app.whenReady().then(async () => {
  logger = createLogger(CONFIG_DIR);
  log("[app] The Sidebar starting...");

  if (isFirstRun()) {
    log("[app] First run detected");
    try {
      await firstRunSetup(resourcePath, log, logError);
    } catch (e: any) {
      logError(`[app] First-run setup failed: ${e.message}`);
    }
  }

  tray = new Tray(nativeImage.createEmpty());
  tray.setTitle("⚖️");
  updateTrayMenu();

  startServer();
  statusInterval = setInterval(() => checkServerStatus(), STATUS_POLL_INTERVAL);
  setupAutoUpdater();
});

app.on("window-all-closed", () => {});
app.on("before-quit", () => stopServer());
