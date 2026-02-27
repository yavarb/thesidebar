/**
 * @module settings
 * Settings & configuration management for The Sidebar.
 * Stores API keys and model configs in ~/.thesidebar/config.json.
 * Provides Express route handlers for GET/POST /api/settings.
 */

import fs from "fs";
import path from "path";

// ── Types ──

/** A local OpenAI-compatible endpoint */
export interface LocalEndpointConfig {
  name: string;
  baseUrl: string;
}

/** Full configuration stored on disk */
export interface SidebarConfig {
  /** OpenClaw gateway URL (e.g., "http://10.0.0.58:18789") */
  openclawUrl?: string;
  /** OpenClaw gateway auth token */
  openclawToken?: string;
  /** Anthropic API key */
  anthropicApiKey?: string;
  /** OpenAI API key */
  openaiApiKey?: string;
  /** Local endpoints */
  localEndpoints?: LocalEndpointConfig[];
  /** Default model string */
  defaultModel?: string;
  /** Percentage of context window to use for conversation history (default 40) */
  contextBudgetPercent?: number;
  /** Session TTL in days. Default 30. 0 or -1 means forever. */
  sessionTTLDays?: number;
}

/** Sensitive field names that should be masked in GET responses */
const SENSITIVE_FIELDS: (keyof SidebarConfig)[] = ["anthropicApiKey", "openaiApiKey", "openclawToken"];

// ── Config Path ──

/** Get the config directory path (~/.thesidebar/) */
export function getConfigDir(): string {
  return path.join(process.env.HOME || process.env.USERPROFILE || "~", ".thesidebar");
}

/** Get the config file path (~/.thesidebar/config.json) */
export function getConfigPath(): string {
  return path.join(getConfigDir(), "config.json");
}

// ── CRUD ──

/**
 * Read the current configuration from disk.
 * Returns empty config if file doesn't exist.
 */
export function readConfig(configPath?: string): SidebarConfig {
  const p = configPath || getConfigPath();
  try {
    const raw = fs.readFileSync(p, "utf-8");
    return JSON.parse(raw) as SidebarConfig;
  } catch {
    return {};
  }
}

/**
 * Write configuration to disk.
 * Creates ~/.thesidebar/ directory if it doesn't exist.
 */
export function writeConfig(config: SidebarConfig, configPath?: string): void {
  const p = configPath || getConfigPath();
  const dir = path.dirname(p);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  }
  fs.writeFileSync(p, JSON.stringify(config, null, 2), { mode: 0o600 });
}

/**
 * Mask sensitive fields in a config object for safe display.
 * Replaces API keys with "sk-****xxxx" format showing only last 4 chars.
 */
export function maskConfig(config: SidebarConfig): SidebarConfig {
  const masked = { ...config };
  for (const field of SENSITIVE_FIELDS) {
    const val = masked[field] as string | undefined;
    if (val && typeof val === "string" && val.length > 4) {
      (masked as any)[field] = val.slice(0, 3) + "****" + val.slice(-4);
    } else if (val) {
      (masked as any)[field] = "****";
    }
  }
  return masked;
}

/**
 * Merge partial config updates into existing config.
 * Only updates fields that are explicitly provided.
 * Skips masked values (containing "****") to prevent overwriting real keys with masks.
 */
export function mergeConfig(existing: SidebarConfig, updates: Partial<SidebarConfig>): SidebarConfig {
  const merged = { ...existing };

  for (const [key, value] of Object.entries(updates)) {
    if (value === undefined) continue;

    // Skip masked values — don't overwrite real keys with masked versions
    if (typeof value === "string" && value.includes("****")) continue;

    (merged as any)[key] = value;
  }

  return merged;
}

// ── Express Route Handlers ──

/**
 * GET /api/settings handler.
 * Returns masked config.
 */
export function handleGetSettings(configPath?: string) {
  return (_req: any, res: any) => {
    const config = readConfig(configPath);
    res.json({ ok: true, data: maskConfig(config) });
  };
}

/**
 * POST /api/settings handler.
 * Merges updates into existing config and saves.
 */
export function handlePostSettings(configPath?: string) {
  return (req: any, res: any) => {
    const updates = req.body as Partial<SidebarConfig>;
    if (!updates || typeof updates !== "object") {
      return res.status(400).json({ ok: false, error: "Request body must be a JSON object" });
    }

    const existing = readConfig(configPath);
    const merged = mergeConfig(existing, updates);
    writeConfig(merged, configPath);

    res.json({ ok: true, data: maskConfig(merged) });
  };
}
