/**
 * @module sessions
 * Encrypted per-document session persistence.
 * 
 * Each session is encrypted with AES-256-GCM using a key derived via HKDF
 * from a machine-specific key + session ID. Files stored as binary:
 * [12-byte IV][16-byte authTag][ciphertext]
 */

import crypto from "crypto";
import fs from "fs";
import path from "path";

const CONFIG_DIR = path.join(process.env.HOME || "~", ".thesidebar");
const SESSIONS_DIR = path.join(CONFIG_DIR, "sessions");
const MACHINE_KEY_PATH = path.join(CONFIG_DIR, ".machine-key");

export interface SessionData {
  sessionId: string;
  conversationHistory: { role: string; content: string }[];
  model: string;
  createdAt: number;
  updatedAt: number;
  documentName?: string;
}

function ensureSessionsDir(): void {
  if (!fs.existsSync(SESSIONS_DIR)) {
    fs.mkdirSync(SESSIONS_DIR, { recursive: true, mode: 0o700 });
  }
}

function getMachineKey(): Buffer {
  if (!fs.existsSync(MACHINE_KEY_PATH)) {
    throw new Error("Machine key not found. Run first-run setup.");
  }
  return fs.readFileSync(MACHINE_KEY_PATH);
}

function deriveKey(machineKey: Buffer, sessionId: string): Buffer {
  return Buffer.from(crypto.hkdfSync("sha256", machineKey, sessionId, "thesidebar-session", 32));
}

function encrypt(data: SessionData, key: Buffer): Buffer {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const json = JSON.stringify(data);
  const encrypted = Buffer.concat([cipher.update(json, "utf8"), cipher.final()]);
  const authTag = cipher.getAuthTag();
  return Buffer.concat([iv, authTag, encrypted]);
}

function decrypt(buf: Buffer, key: Buffer): SessionData {
  if (buf.length < 28) throw new Error("Invalid session file");
  const iv = buf.subarray(0, 12);
  const authTag = buf.subarray(12, 28);
  const ciphertext = buf.subarray(28);
  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAuthTag(authTag);
  const decrypted = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
  return JSON.parse(decrypted.toString("utf8"));
}

function sessionPath(sessionId: string): string {
  const safe = sessionId.replace(/[^a-zA-Z0-9\-_]/g, "");
  return path.join(SESSIONS_DIR, `${safe}.enc`);
}

export function saveSession(data: SessionData): void {
  ensureSessionsDir();
  const key = deriveKey(getMachineKey(), data.sessionId);
  data.updatedAt = Date.now();
  const encrypted = encrypt(data, key);
  fs.writeFileSync(sessionPath(data.sessionId), encrypted, { mode: 0o600 });
}

export function loadSession(sessionId: string): SessionData | null {
  try {
    const fp = sessionPath(sessionId);
    if (!fs.existsSync(fp)) return null;
    const buf = fs.readFileSync(fp);
    const key = deriveKey(getMachineKey(), sessionId);
    return decrypt(buf, key);
  } catch (e: any) {
    console.error(`[sessions] Failed to load session ${sessionId}:`, e.message);
    return null;
  }
}

export function deleteSession(sessionId: string): boolean {
  try {
    const fp = sessionPath(sessionId);
    if (fs.existsSync(fp)) { fs.unlinkSync(fp); return true; }
    return false;
  } catch { return false; }
}

export function cleanExpiredSessions(ttlDays: number): number {
  if (ttlDays <= 0) return 0;
  ensureSessionsDir();
  const cutoff = Date.now() - ttlDays * 86400000;
  let cleaned = 0;
  for (const file of fs.readdirSync(SESSIONS_DIR)) {
    if (!file.endsWith(".enc")) continue;
    try {
      const stat = fs.statSync(path.join(SESSIONS_DIR, file));
      if (stat.mtimeMs < cutoff) { fs.unlinkSync(path.join(SESSIONS_DIR, file)); cleaned++; }
    } catch {}
  }
  if (cleaned > 0) console.log(`[sessions] Cleaned ${cleaned} expired sessions`);
  return cleaned;
}

export function ensureMachineKey(): void {
  if (!fs.existsSync(CONFIG_DIR)) fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  if (!fs.existsSync(MACHINE_KEY_PATH)) {
    fs.writeFileSync(MACHINE_KEY_PATH, crypto.randomBytes(32), { mode: 0o600 });
    console.log("[sessions] Generated machine key");
  }
  ensureSessionsDir();
}
