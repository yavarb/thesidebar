/**
 * @module electron/logger
 * Simple rotating file logger for The Sidebar.
 * Writes to ~/.thesidebar/logs/thesidebar-YYYY-MM-DD.log
 * Rotates daily, keeps last 7 days.
 */

import fs from "fs";
import path from "path";

export interface Logger {
  info(msg: string): void;
  error(msg: string): void;
}

/** Create a file logger that writes to configDir/logs/ */
export function createLogger(configDir: string): Logger {
  const logsDir = path.join(configDir, "logs");
  if (!fs.existsSync(logsDir)) {
    fs.mkdirSync(logsDir, { recursive: true });
  }

  // Clean old logs (keep 7 days)
  try {
    const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
    for (const f of fs.readdirSync(logsDir)) {
      const fp = path.join(logsDir, f);
      const stat = fs.statSync(fp);
      if (stat.mtimeMs < cutoff) fs.unlinkSync(fp);
    }
  } catch {}

  function getLogFile(): string {
    const date = new Date().toISOString().slice(0, 10);
    return path.join(logsDir, `thesidebar-${date}.log`);
  }

  function write(level: string, msg: string): void {
    const ts = new Date().toISOString();
    const line = `${ts} [${level}] ${msg}\n`;
    try {
      fs.appendFileSync(getLogFile(), line);
    } catch {}
  }

  return {
    info: (msg: string) => write("INFO", msg),
    error: (msg: string) => write("ERROR", msg),
  };
}
