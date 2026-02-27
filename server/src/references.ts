/**
 * @module references
 * RAG-based reference document manager for The Sidebar.
 * Scans configured folders for .docx/.pdf/.txt/.md files,
 * chunks them, and retrieves relevant chunks via embeddings or TF-IDF.
 */

import fs from "fs";
import path from "path";
import { v4 as uuidv4 } from "uuid";
import { readConfig } from "./settings";

// ── Types ──

interface ChunkData {
  text: string;
  index: number;
  embedding?: number[];
}

export interface ReferenceDoc {
  id: string;
  filepath: string;
  filename: string;
  addedAt: number;
  mtime: number;
  chunks: ChunkData[];
}

export interface QueryResult {
  docId: string;
  filename: string;
  chunkText: string;
  score: number;
}

// ── State ──
const documents: Map<string, ReferenceDoc> = new Map(); // keyed by filepath
let embeddingMode: "openai" | "local" | "tfidf" = "tfidf";
let idfMap: Map<string, number> = new Map();
let idfDirty = true;
let scanTimer: NodeJS.Timeout | null = null;
let scanning = false;
let lastScanAt = 0;
let lastScanError: string | null = null;

const SUPPORTED_EXTENSIONS = new Set([".docx", ".pdf", ".txt", ".md", ".doc"]);
const SCAN_INTERVAL = 5 * 60 * 1000; // 5 minutes

// ── Text Extraction ──

async function extractText(filepath: string): Promise<string> {
  const ext = path.extname(filepath).toLowerCase();

  if (ext === ".txt" || ext === ".md") {
    return fs.readFileSync(filepath, "utf-8");
  }

  if (ext === ".docx" || ext === ".doc") {
    const mammoth = require("mammoth");
    const result = await mammoth.extractRawText({ path: filepath });
    return result.value;
  }

  if (ext === ".pdf") {
    const pdfParse = require("pdf-parse");
    const buffer = fs.readFileSync(filepath);
    const data = await pdfParse(buffer);
    return data.text;
  }

  throw new Error(`Unsupported file type: ${ext}`);
}

// ── Chunking ──

function chunkText(text: string, targetChars: number = 2000): { text: string; index: number }[] {
  const paragraphs = text.split(/\n\s*\n/).filter(p => p.trim().length > 0);
  const chunks: { text: string; index: number }[] = [];
  let current = "";
  let chunkIndex = 0;

  for (const para of paragraphs) {
    const trimmed = para.trim();
    if (current.length + trimmed.length + 2 > targetChars && current.length > 0) {
      chunks.push({ text: current.trim(), index: chunkIndex++ });
      const lastSentence = extractLastSentence(current);
      current = lastSentence ? lastSentence + "\n\n" + trimmed : trimmed;
    } else {
      current += (current ? "\n\n" : "") + trimmed;
    }
  }

  if (current.trim()) {
    chunks.push({ text: current.trim(), index: chunkIndex });
  }

  return chunks;
}

function extractLastSentence(text: string): string {
  const sentences = text.match(/[^.!?]+[.!?]+/g);
  if (!sentences || sentences.length === 0) return "";
  return sentences[sentences.length - 1].trim();
}

// ── Cosine Similarity ──

function cosineSimilarity(a: number[], b: number[]): number {
  if (a.length !== b.length || a.length === 0) return 0;
  let dot = 0, normA = 0, normB = 0;
  for (let i = 0; i < a.length; i++) {
    dot += a[i] * b[i];
    normA += a[i] * a[i];
    normB += b[i] * b[i];
  }
  const denom = Math.sqrt(normA) * Math.sqrt(normB);
  return denom === 0 ? 0 : dot / denom;
}

// ── TF-IDF Fallback ──

function tokenize(text: string): string[] {
  return text.toLowerCase().replace(/[^\w\s]/g, " ").split(/\s+/).filter(w => w.length > 1);
}

function computeTF(tokens: string[]): Map<string, number> {
  const tf = new Map<string, number>();
  for (const t of tokens) tf.set(t, (tf.get(t) || 0) + 1);
  for (const [k, v] of tf) tf.set(k, v / tokens.length);
  return tf;
}

function rebuildIDF(): void {
  if (!idfDirty) return;
  const allChunks: ChunkData[] = [];
  for (const doc of documents.values()) allChunks.push(...doc.chunks);
  const docCount = allChunks.length;
  if (docCount === 0) { idfMap = new Map(); idfDirty = false; return; }

  const docFreq = new Map<string, number>();
  for (const chunk of allChunks) {
    const uniqueTokens = new Set(tokenize(chunk.text));
    for (const t of uniqueTokens) docFreq.set(t, (docFreq.get(t) || 0) + 1);
  }

  idfMap = new Map();
  for (const [term, freq] of docFreq) {
    idfMap.set(term, Math.log((docCount + 1) / (freq + 1)) + 1);
  }
  idfDirty = false;
}

function tfidfVector(text: string): Map<string, number> {
  rebuildIDF();
  const tokens = tokenize(text);
  const tf = computeTF(tokens);
  const vec = new Map<string, number>();
  for (const [term, tfVal] of tf) {
    const idf = idfMap.get(term) || Math.log((documents.size + 1) / 1) + 1;
    vec.set(term, tfVal * idf);
  }
  return vec;
}

function tfidfSimilarity(a: Map<string, number>, b: Map<string, number>): number {
  const allTerms = new Set([...a.keys(), ...b.keys()]);
  let dot = 0, normA = 0, normB = 0;
  for (const t of allTerms) {
    const av = a.get(t) || 0;
    const bv = b.get(t) || 0;
    dot += av * bv; normA += av * av; normB += bv * bv;
  }
  const denom = Math.sqrt(normA) * Math.sqrt(normB);
  return denom === 0 ? 0 : dot / denom;
}

// ── Embedding via API ──

async function fetchEmbeddings(texts: string[]): Promise<number[][] | null> {
  const config = readConfig();

  if (config.openaiApiKey) {
    try {
      const https = require("https");
      const body = JSON.stringify({ model: "text-embedding-3-small", input: texts });
      const data = await new Promise<string>((resolve, reject) => {
        const req = https.request("https://api.openai.com/v1/embeddings", {
          method: "POST",
          headers: { "Content-Type": "application/json", "Authorization": `Bearer ${config.openaiApiKey}` },
        }, (res: any) => {
          let body = ""; res.on("data", (c: Buffer) => body += c); res.on("end", () => resolve(body));
        });
        req.on("error", reject);
        req.setTimeout(30000, () => { req.destroy(); reject(new Error("Timeout")); });
        req.write(body); req.end();
      });
      const parsed = JSON.parse(data);
      if (parsed.data) { embeddingMode = "openai"; return parsed.data.map((d: any) => d.embedding); }
    } catch (e: any) {
      console.error("[references] OpenAI embeddings failed:", e.message);
    }
  }

  if (config.localEndpoints) {
    for (const ep of config.localEndpoints) {
      try {
        const url = new URL(ep.baseUrl.replace(/\/$/, "") + "/v1/embeddings");
        const body = JSON.stringify({ input: texts, model: "default" });
        const mod = url.protocol === "https:" ? require("https") : require("http");
        const data = await new Promise<string>((resolve, reject) => {
          const req = mod.request(url, { method: "POST", headers: { "Content-Type": "application/json" } },
            (res: any) => { let body = ""; res.on("data", (c: Buffer) => body += c); res.on("end", () => resolve(body)); });
          req.on("error", reject);
          req.setTimeout(30000, () => { req.destroy(); reject(new Error("Timeout")); });
          req.write(body); req.end();
        });
        const parsed = JSON.parse(data);
        if (parsed.data) { embeddingMode = "local"; return parsed.data.map((d: any) => d.embedding); }
      } catch (e: any) {
        console.error(`[references] Local embeddings (${ep.name}) failed:`, e.message);
      }
    }
  }

  return null;
}

async function embedChunks(chunks: ChunkData[]): Promise<void> {
  const texts = chunks.map(c => c.text);
  const batchSize = 2048;
  for (let i = 0; i < texts.length; i += batchSize) {
    const batch = texts.slice(i, i + batchSize);
    const embeddings = await fetchEmbeddings(batch);
    if (embeddings) {
      for (let j = 0; j < batch.length; j++) chunks[i + j].embedding = embeddings[j];
    } else {
      embeddingMode = "tfidf";
      break;
    }
  }
}

// ── Folder Scanning ──

function findSupportedFiles(folderPath: string): string[] {
  const results: string[] = [];
  function walk(dir: string) {
    let entries: fs.Dirent[];
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const entry of entries) {
      if (entry.name.startsWith(".") || entry.name.startsWith("~$")) continue; // skip hidden/temp
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) { walk(full); }
      else if (entry.isFile() && SUPPORTED_EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
        results.push(full);
      }
    }
  }
  walk(folderPath);
  return results;
}

export async function scanFolders(folders: string[]): Promise<{ added: number; updated: number; removed: number; errors: string[] }> {
  if (scanning) return { added: 0, updated: 0, removed: 0, errors: ["Scan already in progress"] };
  scanning = true;
  lastScanError = null;

  const result = { added: 0, updated: 0, removed: 0, errors: [] as string[] };

  try {
    // Collect all files from all folders
    const allFiles = new Set<string>();
    for (const folder of folders) {
      if (!folder || !fs.existsSync(folder)) {
        result.errors.push(`Folder not found: ${folder}`);
        continue;
      }
      const files = findSupportedFiles(folder);
      for (const f of files) allFiles.add(f);
    }

    console.log(`[references] Scanning ${folders.length} folder(s), found ${allFiles.size} files`);

    // Remove docs whose files no longer exist
    for (const [filepath, doc] of documents) {
      if (!allFiles.has(filepath)) {
        documents.delete(filepath);
        result.removed++;
        idfDirty = true;
      }
    }

    // Add or update files
    for (const filepath of allFiles) {
      try {
        const stat = fs.statSync(filepath);
        const mtime = stat.mtimeMs;
        const existing = documents.get(filepath);

        if (existing && existing.mtime >= mtime) continue; // unchanged

        const text = await extractText(filepath);
        if (!text.trim()) continue;

        const chunks: ChunkData[] = chunkText(text);
        await embedChunks(chunks);

        const doc: ReferenceDoc = {
          id: existing?.id || uuidv4(),
          filepath,
          filename: path.basename(filepath),
          addedAt: existing?.addedAt || Date.now(),
          mtime,
          chunks,
        };

        documents.set(filepath, doc);
        idfDirty = true;

        if (existing) { result.updated++; } else { result.added++; }
        console.log(`[references] ${existing ? "Updated" : "Added"} "${doc.filename}" (${chunks.length} chunks)`);
      } catch (e: any) {
        result.errors.push(`${path.basename(filepath)}: ${e.message}`);
      }
    }

    lastScanAt = Date.now();
    console.log(`[references] Scan complete: +${result.added} ~${result.updated} -${result.removed}, ${documents.size} total docs, mode=${embeddingMode}`);
  } catch (e: any) {
    lastScanError = e.message;
    result.errors.push(e.message);
  } finally {
    scanning = false;
  }

  return result;
}

/** Start periodic scanning based on configured folders */
export function startPeriodicScan(): void {
  stopPeriodicScan();

  // Initial scan
  const config = readConfig();
  const folders = config.referenceFolders || [];
  if (folders.length > 0) {
    scanFolders(folders).catch(e => console.error("[references] Initial scan failed:", e.message));
  }

  // Periodic re-scan
  scanTimer = setInterval(() => {
    const config = readConfig();
    const folders = config.referenceFolders || [];
    if (folders.length > 0) {
      scanFolders(folders).catch(e => console.error("[references] Periodic scan failed:", e.message));
    }
  }, SCAN_INTERVAL);
}

export function stopPeriodicScan(): void {
  if (scanTimer) { clearInterval(scanTimer); scanTimer = null; }
}

/** Trigger an immediate rescan (e.g., after settings change) */
export async function rescan(): Promise<{ added: number; updated: number; removed: number; errors: string[] }> {
  const config = readConfig();
  const folders = config.referenceFolders || [];
  if (folders.length === 0) {
    // Clear everything if no folders configured
    const count = documents.size;
    documents.clear();
    idfDirty = true;
    return { added: 0, updated: 0, removed: count, errors: [] };
  }
  return scanFolders(folders);
}

// ── Query ──

export async function queryDocuments(text: string, topK: number = 5): Promise<QueryResult[]> {
  if (documents.size === 0) return [];

  const firstDoc = documents.values().next().value!;
  const hasEmbeddings = firstDoc.chunks[0]?.embedding;

  if (hasEmbeddings) {
    const queryEmbeddings = await fetchEmbeddings([text]);
    if (queryEmbeddings && queryEmbeddings[0]) {
      const queryVec = queryEmbeddings[0];
      const results: QueryResult[] = [];
      for (const doc of documents.values()) {
        for (const chunk of doc.chunks) {
          if (chunk.embedding) {
            const score = cosineSimilarity(queryVec, chunk.embedding);
            results.push({ docId: doc.id, filename: doc.filename, chunkText: chunk.text, score });
          }
        }
      }
      results.sort((a, b) => b.score - a.score);
      return deduplicateResults(results.slice(0, topK * 2), topK);
    }
  }

  return queryWithTFIDF(text, topK);
}

function queryWithTFIDF(text: string, topK: number): QueryResult[] {
  const queryVec = tfidfVector(text);
  const results: QueryResult[] = [];
  for (const doc of documents.values()) {
    for (const chunk of doc.chunks) {
      const chunkVec = tfidfVector(chunk.text);
      const score = tfidfSimilarity(queryVec, chunkVec);
      if (score > 0) results.push({ docId: doc.id, filename: doc.filename, chunkText: chunk.text, score });
    }
  }
  results.sort((a, b) => b.score - a.score);
  return deduplicateResults(results.slice(0, topK * 2), topK);
}

function deduplicateResults(results: QueryResult[], topK: number): QueryResult[] {
  const final: QueryResult[] = [];
  for (const r of results) {
    if (final.length >= topK) break;
    const existing = final.find(f =>
      f.docId === r.docId &&
      (f.chunkText.endsWith(r.chunkText.substring(0, 100)) || r.chunkText.endsWith(f.chunkText.substring(0, 100)))
    );
    if (existing) {
      if (r.chunkText.length > existing.chunkText.length) existing.chunkText += "\n\n" + r.chunkText;
      existing.score = Math.max(existing.score, r.score);
    } else {
      final.push({ ...r });
    }
  }
  return final;
}

// ── Info ──

export function listDocuments(): { id: string; filepath: string; filename: string; chunkCount: number; addedAt: number }[] {
  return Array.from(documents.values()).map(d => ({
    id: d.id, filepath: d.filepath, filename: d.filename, chunkCount: d.chunks.length, addedAt: d.addedAt,
  }));
}

export function getStatus(): {
  documentCount: number; totalChunks: number; embeddingMode: string;
  scanning: boolean; lastScanAt: number; lastScanError: string | null;
} {
  let totalChunks = 0;
  for (const doc of documents.values()) totalChunks += doc.chunks.length;
  return { documentCount: documents.size, totalChunks, embeddingMode, scanning, lastScanAt, lastScanError };
}

export function getDocumentCount(): number { return documents.size; }
