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
  conversationHistory: { role: string; content: string; timestamp?: number }[];
  model: string;
  createdAt: number;
  updatedAt: number;
  documentName?: string;
  changeSummaries?: { exchangeIndex: number; summary: string }[];
  topicsDiscussed?: string[];
  lastRecap?: string;
}

// ── Stopwords for topic extraction ──
const STOPWORDS = new Set([
  "the", "a", "an", "is", "are", "was", "were", "be", "been", "being",
  "have", "has", "had", "do", "does", "did", "will", "would", "could",
  "should", "may", "might", "shall", "can", "need", "dare", "ought",
  "and", "but", "or", "nor", "not", "so", "yet", "both", "either",
  "neither", "each", "every", "all", "any", "few", "more", "most",
  "other", "some", "such", "no", "only", "own", "same", "than",
  "too", "very", "just", "because", "as", "until", "while", "of",
  "at", "by", "for", "with", "about", "against", "between", "through",
  "during", "before", "after", "above", "below", "to", "from", "up",
  "down", "in", "out", "on", "off", "over", "under", "again", "further",
  "then", "once", "here", "there", "when", "where", "why", "how",
  "what", "which", "who", "whom", "this", "that", "these", "those",
  "i", "me", "my", "myself", "we", "our", "ours", "you", "your",
  "he", "him", "his", "she", "her", "it", "its", "they", "them",
  "their", "if", "also", "please", "thank", "thanks", "yes", "no",
  "okay", "ok", "sure", "right", "well", "like", "know", "think",
  "want", "get", "make", "go", "see", "look", "come", "take", "use",
]);

const ACTION_KEYWORDS: Record<string, string> = {
  "edited": "edits",
  "replaced": "replacements",
  "footnote": "footnotes",
  "searched": "searches",
  "formatted": "formatting",
  "inserted": "insertions",
  "comment": "comments",
  "found": "searches",
  "updated": "updates",
  "paragraph": "paragraph edits",
  "style": "style changes",
  "heading": "heading changes",
};

/**
 * Generate a session recap from stored history without calling any LLM.
 */
export function generateSessionRecap(session: SessionData): string {
  const history = session.conversationHistory;
  if (!history || history.length === 0) return "";

  const userMessages = history.filter(m => m.role === "user");
  const assistantMessages = history.filter(m => m.role === "assistant");
  const exchangeCount = Math.min(userMessages.length, assistantMessages.length);

  // Time range
  const timestamps = history.filter(m => m.timestamp).map(m => m.timestamp!);
  const firstTs = timestamps.length > 0 ? Math.min(...timestamps) : session.createdAt;
  const lastTs = timestamps.length > 0 ? Math.max(...timestamps) : session.updatedAt;
  const daySpan = Math.max(1, Math.ceil((lastTs - firstTs) / 86400000));

  // Extract action counts from assistant messages
  const actionCounts: Record<string, number> = {};
  for (const msg of assistantMessages) {
    const lower = msg.content.toLowerCase();
    for (const [keyword, category] of Object.entries(ACTION_KEYWORDS)) {
      const regex = new RegExp(keyword, "gi");
      const matches = lower.match(regex);
      if (matches) {
        actionCounts[category] = (actionCounts[category] || 0) + matches.length;
      }
    }
  }

  // Extract top terms from user messages
  const wordFreq: Record<string, number> = {};
  for (const msg of userMessages) {
    const words = msg.content.toLowerCase().replace(/[^a-z0-9\s-]/g, "").split(/\s+/);
    for (const w of words) {
      if (w.length > 3 && !STOPWORDS.has(w)) {
        wordFreq[w] = (wordFreq[w] || 0) + 1;
      }
    }
  }
  const topTerms = Object.entries(wordFreq)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6)
    .map(([word]) => word);

  // Format recap
  const lines: string[] = [];
  lines.push(`Session recap (${exchangeCount} exchanges over ${daySpan} day${daySpan !== 1 ? "s" : ""}):`);

  if (session.changeSummaries && session.changeSummaries.length > 0) {
    const totalEdits = session.changeSummaries.length;
    lines.push(`- Made ${totalEdits} tracked document edit${totalEdits !== 1 ? "s" : ""}`);
    const recent = session.changeSummaries.slice(-3);
    for (const cs of recent) {
      lines.push(`  • ${cs.summary}`);
    }
  }

  const actionEntries = Object.entries(actionCounts).filter(([, v]) => v > 0);
  if (actionEntries.length > 0) {
    const actionStr = actionEntries.map(([k, v]) => `${v} ${k}`).join(", ");
    lines.push(`- Actions detected: ${actionStr}`);
  }

  const lastActiveDate = new Date(lastTs);
  lines.push(`- Last active: ${lastActiveDate.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}`);

  if (topTerms.length > 0) {
    lines.push(`- Recent topics: ${topTerms.join(", ")}`);
  }

  return lines.join("\n");
}

/**
 * Search through conversation history for matching exchanges.
 */
export function searchConversationHistory(
  session: SessionData,
  query: string
): { role: string; content: string; timestamp?: number; exchangeIndex: number }[] {
  const results: { role: string; content: string; timestamp?: number; exchangeIndex: number }[] = [];
  const lowerQuery = query.toLowerCase();
  const history = session.conversationHistory;

  for (let i = 0; i < history.length; i++) {
    if (history[i].content.toLowerCase().includes(lowerQuery)) {
      const exchangeIndex = Math.floor(i / 2);
      results.push({ ...history[i], exchangeIndex });
      const pairIdx = history[i].role === "user" ? i + 1 : i - 1;
      if (pairIdx >= 0 && pairIdx < history.length && !history[pairIdx].content.toLowerCase().includes(lowerQuery)) {
        results.push({ ...history[pairIdx], exchangeIndex });
      }
    }
  }

  return results;
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
