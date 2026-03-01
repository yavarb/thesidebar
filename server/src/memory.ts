/**
 * @module memory
 * Learning memory system for The Sidebar.
 * 
 * Two tiers:
 * - Global memory (~/.thesidebar/memory.json) — preferences across all documents
 * - Document memory (per session) — corrections specific to a document
 */

import fs from "fs";
import path from "path";

const CONFIG_DIR = path.join(process.env.HOME || "~", ".thesidebar");
const GLOBAL_MEMORY_PATH = path.join(CONFIG_DIR, "memory.json");
const DOC_MEMORY_DIR = path.join(CONFIG_DIR, "doc-memory");

export interface MemoryEntry {
  id: string;
  text: string;
  scope: "global" | "document";
  category: "preference" | "correction" | "fact" | "style";
  created: number;
  source?: string;
}

export interface MemoryStore {
  version: number;
  entries: MemoryEntry[];
}

function ensureDir(dir: string): void {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function generateId(): string {
  return `mem_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

// ── Global Memory ──

export function readGlobalMemory(): MemoryStore {
  try {
    if (fs.existsSync(GLOBAL_MEMORY_PATH)) {
      return JSON.parse(fs.readFileSync(GLOBAL_MEMORY_PATH, "utf8"));
    }
  } catch {}
  return { version: 1, entries: [] };
}

export function writeGlobalMemory(store: MemoryStore): void {
  ensureDir(CONFIG_DIR);
  fs.writeFileSync(GLOBAL_MEMORY_PATH, JSON.stringify(store, null, 2));
}

export function addGlobalMemory(text: string, category: MemoryEntry["category"], source?: string): MemoryEntry {
  const store = readGlobalMemory();
  const isDuplicate = store.entries.some(e => e.text.toLowerCase().trim() === text.toLowerCase().trim());
  if (isDuplicate) return store.entries.find(e => e.text.toLowerCase().trim() === text.toLowerCase().trim())!;
  const entry: MemoryEntry = { id: generateId(), text, scope: "global", category, created: Date.now(), source };
  store.entries.push(entry);
  
  writeGlobalMemory(store);
  return entry;
}

// ── Document Memory ──

function docMemoryPath(sessionId: string): string {
  return path.join(DOC_MEMORY_DIR, `${sessionId}.json`);
}

export function readDocMemory(sessionId: string): MemoryStore {
  try {
    const p = docMemoryPath(sessionId);
    if (fs.existsSync(p)) return JSON.parse(fs.readFileSync(p, "utf8"));
  } catch {}
  return { version: 1, entries: [] };
}

export function writeDocMemory(sessionId: string, store: MemoryStore): void {
  ensureDir(DOC_MEMORY_DIR);
  fs.writeFileSync(docMemoryPath(sessionId), JSON.stringify(store, null, 2));
}

export function addDocMemory(sessionId: string, text: string, category: MemoryEntry["category"], source?: string): MemoryEntry {
  const store = readDocMemory(sessionId);
  const isDuplicate = store.entries.some(e => e.text.toLowerCase().trim() === text.toLowerCase().trim());
  if (isDuplicate) return store.entries.find(e => e.text.toLowerCase().trim() === text.toLowerCase().trim())!;
  const entry: MemoryEntry = { id: generateId(), text, scope: "document", category, created: Date.now(), source };
  store.entries.push(entry);
  
  writeDocMemory(sessionId, store);
  return entry;
}

// ── Memory Injection ──

export function buildMemoryContext(sessionId?: string): string {
  const global = readGlobalMemory();
  const doc = sessionId ? readDocMemory(sessionId) : { entries: [] };
  const allEntries = [...global.entries, ...doc.entries];
  if (allEntries.length === 0) return "";

  let context = "\n\n## Learned Preferences & Corrections\nApply these automatically — they are things learned from working with this user:\n\n";
  
  const groups: Record<string, { label: string; entries: MemoryEntry[] }> = {
    preference: { label: "Preferences", entries: [] },
    correction: { label: "Corrections (do not repeat these mistakes)", entries: [] },
    style: { label: "Style rules", entries: [] },
    fact: { label: "Facts about this document/case", entries: [] },
  };

  for (const e of allEntries) {
    if (groups[e.category]) groups[e.category].entries.push(e);
  }

  for (const g of Object.values(groups)) {
    if (g.entries.length > 0) {
      context += `**${g.label}:**\n`;
      for (const e of g.entries) context += `- ${e.text}\n`;
      context += "\n";
    }
  }
  return context;
}

// ── Memory Extraction ──

export const MEMORY_EXTRACTION_PROMPT = `Review the conversation above. Did the user:
1. Correct a mistake you made? (category: "correction")
2. Express a preference for how things should be done? (category: "preference") 
3. Specify a style rule or formatting convention? (category: "style")
4. Provide an important fact about the document, case, or project? (category: "fact")

Respond with a JSON array of learnings, or [] if nothing worth remembering.
Format: [{"text": "concise learning", "category": "preference|correction|style|fact", "scope": "global|document"}]
Be concise. Only extract genuine learnings. Respond ONLY with the JSON array.`;

// ── REST API handlers ──

export function getMemoryEntries(sessionId?: string): MemoryEntry[] {
  const global = readGlobalMemory().entries;
  const doc = sessionId ? readDocMemory(sessionId).entries : [];
  return [...global, ...doc];
}

export function deleteMemoryEntry(id: string, sessionId?: string): boolean {
  // Try global first
  const global = readGlobalMemory();
  const gIdx = global.entries.findIndex(e => e.id === id);
  if (gIdx >= 0) {
    global.entries.splice(gIdx, 1);
    writeGlobalMemory(global);
    return true;
  }
  // Try document memory
  if (sessionId) {
    const doc = readDocMemory(sessionId);
    const dIdx = doc.entries.findIndex(e => e.id === id);
    if (dIdx >= 0) {
      doc.entries.splice(dIdx, 1);
      writeDocMemory(sessionId, doc);
      return true;
    }
  }
  return false;
}
