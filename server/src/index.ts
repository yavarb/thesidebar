import { runAgentLoop } from "./agent-loop";
import { getContextSize, manageContext } from "./context";
import { resolveModel } from "./llm-router";
import { readConfig } from "./settings";
import { handleGetSettings, handlePostSettings } from "./settings";
import express from "express";
import compression from "compression";
import cors from "cors";

import http from "http";
import fs from "fs";
import path from "path";
import { WebSocketServer, WebSocket } from "ws";

// ── Config ──
const PORT = parseInt(process.env.SIDEBAR_PORT || "3001", 10);
const VERSION = "0.3.0";
const HEARTBEAT_INTERVAL = 10000;
const DEFAULT_TIMEOUT = 20000;

const app = express();
app.use(cors());
app.use(compression());
app.use(express.json({ limit: "10mb" }));

// ── Static Files (task pane UI) ──
const appDistCandidates = [
  path.join(__dirname, "../../app/dist"),        // packaged app & dev
  path.join(__dirname, "../../../app/dist"),      // fallback
];
for (const dir of appDistCandidates) {
  if (fs.existsSync(dir)) {
    app.use(express.static(dir));
    break;
  }
}

// ── HTTP Server (localhost only) ──
const server = http.createServer(app);
const wss = new WebSocketServer({ server });

// ── State ──
/** Session ID for OpenClaw context persistence */
let sessionId = "sidebar-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8);
/** Conversation history for non-OpenClaw models */
let conversationHistory: { role: string; content: string }[] = [];

/** Single task pane WebSocket connection (localhost only) */
let taskPaneWs: WebSocket | null = null;
let globalRequestId = 0;
const pending = new Map<number, { resolve: (v: any) => void; reject: (e: any) => void; timer: NodeJS.Timeout; command: string }>();

// Prompt queue
type PromptEntry = { id: number; text: string; model?: string; context?: string; timestamp: number; clientId?: string };
const promptQueue: PromptEntry[] = [];
const promptLog = new Map<number, PromptEntry>();
let promptId = 0;
const promptWaiters: { resolve: (v: any) => void; timer: NodeJS.Timeout }[] = [];

// ── Change Tracking & Revert ──
const MODIFYING_TOOLS = new Set([
  "updateParagraph", "replaceParagraph", "replaceSelection", "editSelection",
  "findReplace", "insert", "format", "setStyleFont", "addFootnote",
  "updateFootnote", "addComment", "batch",
]);
/** Maps exchangeId → number of modifying tool calls made during that exchange */
const exchangeUndoCounts: Map<number, number> = new Map();
let exchangeIdCounter = 0;


// ── Auto-process prompts via agent loop ──
async function processPrompt(entry: PromptEntry, ws: any) {
  try {
    // Get document context
    let documentContext = "";
    try {
      const doc = await sendCommand("getDocument");
      documentContext = doc?.text || "";
    } catch (e: any) {
      console.error("[agent] Failed to get document:", e.message);
    }

    // Read config for API keys
    const config = readConfig();
    const routerConfig = {
      openclawUrl: config.openclawUrl,
      openclawToken: config.openclawToken,
      openaiApiKey: config.openaiApiKey,
      anthropicApiKey: config.anthropicApiKey,
      localEndpoints: config.localEndpoints,
    };

    const model = entry.model || config.defaultModel || "openclaw";

    // Build prompt with context
    let fullPrompt = entry.text;
    if (entry.context) {
      fullPrompt = `[Selected text: ${entry.context}]\n\n${entry.text}`;
    }

    // Determine backend for conversation history strategy
    // For non-OpenClaw: manage context window
    let managedHistory: { role: string; content: string }[] | undefined;
    const isOpenClaw = !model.startsWith("openai:") && !model.startsWith("anthropic:") && !model.startsWith("local:");

    if (!isOpenClaw) {
      // No context management needed — OpenClaw handles its own sessions
    } else {
      // This branch is dead (isOpenClaw is true) — managedHistory stays undefined
    }

    if (!isOpenClaw && conversationHistory.length > 0) {
      const spec = resolveModel(model, routerConfig);
      const ctxConfig = {
        openaiApiKey: config.openaiApiKey,
        anthropicApiKey: config.anthropicApiKey,
        localBaseUrl: spec.baseUrl,
      };
      const contextSize = await getContextSize(spec.backend, spec.modelId, ctxConfig);
      const budgetPercent = config.contextBudgetPercent ?? 40;
      const managed = manageContext(conversationHistory, contextSize, documentContext, budgetPercent);
      managedHistory = managed.messages;
      if (managed.compactedCount > 0) {
        console.log(`[context] Compacted ${managed.compactedCount} messages, ~${managed.estimatedTokens} tokens used (budget: ${Math.floor(contextSize * budgetPercent / 100)})`);
      }
    }

    // Run agent loop and stream results — track modifying tool calls for revert
    const currentExchangeId = ++exchangeIdCounter;
    let modifyingCallCount = 0;
    let fullResponse = "";
    for await (const chunk of runAgentLoop({
      prompt: fullPrompt,
      model,
      config: routerConfig,
      documentContext,
      sessionUser: isOpenClaw ? sessionId : undefined,
      conversationHistory: !isOpenClaw ? managedHistory : undefined,
      onToolCall: (call) => {
        if (MODIFYING_TOOLS.has(call.name)) modifyingCallCount++;
        if (ws.readyState === 1) {
          ws.send(JSON.stringify({ type: "prompt_progress", promptId: entry.id, text: `Using tool: ${call.name}...` }));
        }
      },
    })) {
      fullResponse += chunk;
      if (ws.readyState === 1) {
        ws.send(JSON.stringify({ type: "prompt_progress", promptId: entry.id, text: fullResponse }));
      }
    }

    // Track conversation history for non-OpenClaw models
    conversationHistory.push({ role: "user", content: fullPrompt });
    conversationHistory.push({ role: "assistant", content: fullResponse || "(No response)" });
    // Keep history bounded
    if (conversationHistory.length > 40) conversationHistory = conversationHistory.slice(-20);

    // Track undo count for this exchange
    const hasChanges = modifyingCallCount > 0;
    if (hasChanges) {
      exchangeUndoCounts.set(currentExchangeId, modifyingCallCount);
      // Keep map bounded
      if (exchangeUndoCounts.size > 50) {
        const oldest = exchangeUndoCounts.keys().next().value;
        if (oldest !== undefined) exchangeUndoCounts.delete(oldest);
      }
    }

    // Send final response
    if (ws.readyState === 1) {
      ws.send(JSON.stringify({ type: "prompt_response", promptId: entry.id, text: fullResponse || "(No response)", timestamp: Date.now(), hasChanges, exchangeId: currentExchangeId }));
    }
  } catch (e: any) {
    console.error("[agent] Error processing prompt:", e.message);
    if (ws.readyState === 1) {
      ws.send(JSON.stringify({ type: "prompt_response", promptId: entry.id, text: `Error: ${e.message}`, timestamp: Date.now() }));
    }
  }
}

// ── WebSocket ──
wss.on("connection", (ws: any) => {
  console.log("[ws] Task pane connected");

  // Close any existing connection (only one allowed)
  if (taskPaneWs && taskPaneWs.readyState === WebSocket.OPEN) taskPaneWs.close();
  taskPaneWs = ws;
  ws._wrAlive = true;

  ws.on("pong", () => { ws._wrAlive = true; });

  ws.on("message", (raw: any) => {
    try {
      const msg = JSON.parse(raw.toString());
      if (msg.type === "prompt") {
        const entry: PromptEntry = { id: ++promptId, text: msg.text, model: msg.model, context: msg.context, timestamp: Date.now(), clientId: msg.clientId };
        promptQueue.push(entry);
        promptLog.set(entry.id, entry);
        while (promptWaiters.length > 0) {
          const w = promptWaiters.shift()!;
          clearTimeout(w.timer);
          w.resolve(entry);
        }
        ws.send(JSON.stringify({ type: "prompt_ack", id: entry.id, clientId: msg.clientId }));
        // Auto-process the prompt
        processPrompt(entry, ws).catch(e => console.error("[agent] processPrompt error:", e));
        return;
      }
      if (msg.id !== undefined && pending.has(msg.id)) {
        const p = pending.get(msg.id)!;
        clearTimeout(p.timer);
        pending.delete(msg.id);
        msg.error ? p.reject(new Error(msg.error)) : p.resolve(msg.data);
      }
    } catch (e) { console.error("[ws] Bad message:", e); }
  });

  ws.on("close", () => {
    console.log("[ws] Task pane disconnected");
    if (taskPaneWs === ws) taskPaneWs = null;
    for (const [id, p] of pending) {
      clearTimeout(p.timer);
      p.reject(new Error("Task pane disconnected"));
      pending.delete(id);
    }
  });
});

// ── Heartbeat ──
const heartbeatInterval = setInterval(() => {
  if (taskPaneWs) {
    if ((taskPaneWs as any)._wrAlive === false) {
      console.log("[ws] Task pane heartbeat timeout");
      taskPaneWs.terminate();
      taskPaneWs = null;
      return;
    }
    (taskPaneWs as any)._wrAlive = false;
    taskPaneWs.ping();
  }
}, HEARTBEAT_INTERVAL);

// ── Graceful Shutdown ──
function shutdown() {
  console.log("[thesidebar] Shutting down...");
  clearInterval(heartbeatInterval);
  if (taskPaneWs) taskPaneWs.close(1001, "Server shutting down");
  for (const [id, p] of pending) { clearTimeout(p.timer); p.reject(new Error("Shutting down")); pending.delete(id); }
  server.close(() => process.exit(0));
  setTimeout(() => process.exit(1), 5000);
}
process.on("SIGTERM", shutdown);
process.on("SIGINT", shutdown);

// ── Send Command ──
/**
 * Send a command to the connected task pane via WebSocket.
 * @param command - Command name to execute in the task pane
 * @param params - Optional parameters for the command
 * @param timeoutMs - Timeout in milliseconds (default: 20000)
 * @returns Promise resolving with the command result
 */
function sendCommand(command: string, params?: any, timeoutMs = DEFAULT_TIMEOUT): Promise<any> {
  return new Promise((resolve, reject) => {
    if (!taskPaneWs || taskPaneWs.readyState !== WebSocket.OPEN) {
      return reject(new Error("Task pane not connected"));
    }
    const id = ++globalRequestId;
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`Timeout (${timeoutMs}ms) on "${command}"`)); }, timeoutMs);
    pending.set(id, { resolve, reject, timer, command });
    taskPaneWs.send(JSON.stringify({ id, command, params }));
  });
}

// ── API Handler ──
/**
 * Create an Express route handler that sends a command to the task pane.
 * @param command - Command name to execute
 * @param extractParams - Optional function to extract params from the request
 */
function apiHandler(command: string, extractParams?: (req: express.Request) => any) {
  return async (req: express.Request, res: express.Response) => {
    const t0 = Date.now();
    try {
      const params = extractParams ? extractParams(req) : (Object.keys(req.body || {}).length ? req.body : undefined);
      const timeoutMs = req.query.timeout ? parseInt(req.query.timeout as string, 10) : DEFAULT_TIMEOUT;
      const data = await sendCommand(command, params, timeoutMs);
      res.json({ ok: true, data, _ms: Date.now() - t0 });
    } catch (e: any) {
      const status = e.message.includes("not connected") ? 503 : e.message.includes("Timeout") ? 504 : 500;
      res.status(status).json({ ok: false, error: e.message });
    }
  };
}

// ═══════════════════════════════════
// ROUTES
// ═══════════════════════════════════

// Health & Meta
app.get("/api/status", (_req, res) => {
  res.json({ ok: true, data: {
    version: VERSION, uptime: process.uptime(), 
    connected: taskPaneWs !== null && taskPaneWs.readyState === WebSocket.OPEN,
    pendingCommands: pending.size, promptQueueSize: promptQueue.length,
  }});
});

app.get("/api/help", (_req, res) => {
  res.json({ ok: true, data: { version: VERSION, endpoints: [
    { m: "GET", p: "/api/status", d: "Server status" },
    { m: "GET", p: "/api/help", d: "Endpoint listing" },
    { m: "GET", p: "/api/ping", d: "Ping task pane" },
    { m: "GET", p: "/api/prompts", d: "Pending prompts" },
    { m: "GET", p: "/api/prompts/wait", d: "Long-poll for prompt" },
    { m: "DELETE", p: "/api/prompts", d: "Clear prompts" },
    { m: "POST", p: "/api/prompts/:id/respond", d: "Push assistant response to task pane" },
    { m: "GET", p: "/api/document", d: "Full text" },
    { m: "GET", p: "/api/document/paragraphs", d: "Paragraphs (?from=&to=&compact=true)" },
    { m: "GET", p: "/api/document/stats", d: "Word/paragraph count" },
    { m: "GET", p: "/api/document/structure", d: "Outline tree" },
    { m: "GET", p: "/api/document/toc", d: "TOC entries" },
    { m: "GET", p: "/api/document/html", d: "HTML export" },
    { m: "GET", p: "/api/paragraph/:index", d: "Single paragraph" },
    { m: "GET", p: "/api/paragraph/:index/context", d: "With surrounding context" },
    { m: "PUT", p: "/api/paragraph/:index", d: "Update paragraph" },
    { m: "POST", p: "/api/paragraph/replace", d: "Safe replace by index/listString" },
    { m: "GET", p: "/api/selection", d: "Read selection" },
    { m: "POST", p: "/api/selection/replace", d: "Replace selection" },
    { m: "POST", p: "/api/select", d: "Select paragraph {index}" },
    { m: "POST", p: "/api/navigate", d: "Scroll to paragraph {index}" },
    { m: "POST", p: "/api/find", d: "Search text" },
    { m: "POST", p: "/api/find-replace", d: "Find & replace" },
    { m: "POST", p: "/api/insert", d: "Insert paragraph" },
    { m: "POST", p: "/api/format", d: "Format text" },
    { m: "GET", p: "/api/styles", d: "List styles" },
    { m: "POST", p: "/api/style/font", d: "Modify style font" },
    { m: "GET", p: "/api/footnotes", d: "List footnotes" },
    { m: "POST", p: "/api/footnote", d: "Add footnote" },
    { m: "PUT", p: "/api/footnote/:index", d: "Update footnote" },
    { m: "POST", p: "/api/footnote/search", d: "Search footnotes" },
    { m: "GET", p: "/api/comments", d: "List comments" },
    { m: "POST", p: "/api/comment", d: "Add comment" },
    { m: "POST", p: "/api/undo", d: "Undo last edit" },
    { m: "GET", p: "/api/undo/history", d: "Undo stack" },
    { m: "POST", p: "/api/batch", d: "Batch operations" },
  ]}});
});



// Prompts
app.get("/api/prompts", (_req, res) => res.json({ ok: true, data: { prompts: promptQueue, count: promptQueue.length }}));
app.get("/api/prompts/wait", (req, res) => {
  if (promptQueue.length > 0) return res.json({ ok: true, data: promptQueue.shift() });
  const timeoutMs = req.query.timeout ? parseInt(req.query.timeout as string, 10) : 30000;
  let done = false;
  const resolve = (prompt: any) => { if (!done) { done = true; res.json({ ok: true, data: prompt }); }};
  const timer = setTimeout(() => { if (!done) { done = true; const idx = promptWaiters.findIndex(w => w.resolve === resolve); if (idx >= 0) promptWaiters.splice(idx, 1); res.json({ ok: true, data: null }); }}, timeoutMs);
  promptWaiters.push({ resolve, timer });
});
app.delete("/api/prompts", (_req, res) => { const c = promptQueue.length; promptQueue.length = 0; res.json({ ok: true, data: { cleared: c }}); });

// Ping
app.get("/api/ping", apiHandler("ping"));

app.post("/api/prompts/:id/respond", (req, res) => {
  const id = parseInt(req.params.id, 10);
  const text = req.body?.text;
  if (!Number.isFinite(id)) return res.status(400).json({ ok: false, error: "valid prompt id required" });
  if (!text || typeof text !== "string") return res.status(400).json({ ok: false, error: "text required" });

  const prompt = promptLog.get(id);
  if (!prompt) return res.status(404).json({ ok: false, error: `prompt ${id} not found` });

  const delivered = taskPaneWs !== null && taskPaneWs.readyState === WebSocket.OPEN;
  if (delivered) {
    taskPaneWs!.send(JSON.stringify({ type: "prompt_response", promptId: id, text, timestamp: Date.now() }));
  }

  return res.json({ ok: true, data: { promptId: id, delivered } });
});


// Index
app.post("/api/index/build", apiHandler("buildIndex"));
app.get("/api/index", apiHandler("getIndex"));
app.get("/api/index/headings", apiHandler("getHeadings"));
app.get("/api/index/range", apiHandler("getIndexRange", (req) => ({ from: req.query.from ? parseInt(req.query.from as string, 10) : undefined, to: req.query.to ? parseInt(req.query.to as string, 10) : undefined })));
app.post("/api/index/delta", apiHandler("getDelta"));

// Document
app.get("/api/document", apiHandler("getDocument"));
app.get("/api/document/paragraphs", apiHandler("getParagraphs", (req) => ({ from: req.query.from ? parseInt(req.query.from as string, 10) : undefined, to: req.query.to ? parseInt(req.query.to as string, 10) : undefined, compact: req.query.compact === "true" })));
app.get("/api/document/stats", apiHandler("getDocumentStats"));
app.get("/api/document/structure", apiHandler("getDocumentStructure"));
app.get("/api/document/toc", apiHandler("getToc"));
app.get("/api/document/html", apiHandler("getDocumentHtml"));

// Paragraphs
app.get("/api/paragraph/:index", apiHandler("getParagraph", (req) => ({ index: parseInt(req.params.index, 10), compact: req.query.compact === "true" })));
app.get("/api/paragraph/:index/context", apiHandler("getParagraphContext", (req) => ({ index: parseInt(req.params.index, 10), radius: req.query.radius ? parseInt(req.query.radius as string, 10) : 2 })));
app.put("/api/paragraph/:index", apiHandler("updateParagraph", (req) => ({ index: parseInt(req.params.index, 10), ...req.body })));
app.post("/api/paragraph/replace", apiHandler("replaceParagraph"));

// Selection
app.get("/api/selection", apiHandler("getSelection"));
app.post("/api/selection/replace", apiHandler("replaceSelection"));
app.post("/api/selection/edit", apiHandler("editSelection"));
app.post("/api/select", apiHandler("selectParagraph"));
app.post("/api/navigate", apiHandler("navigateToParagraph"));

// Find
app.post("/api/find", apiHandler("find"));
app.post("/api/find-replace", apiHandler("findReplace"));

// Insert / Format
app.post("/api/insert", apiHandler("insert"));
app.post("/api/format", apiHandler("format"));

// Styles
app.get("/api/styles", apiHandler("getStyles"));
app.post("/api/style/font", apiHandler("setStyleFont"));

// Footnotes
app.get("/api/footnotes", apiHandler("getFootnotes"));
app.post("/api/footnote", apiHandler("addFootnote"));
app.put("/api/footnote/:index", apiHandler("updateFootnote", (req) => ({ index: parseInt(req.params.index, 10), ...req.body })));
app.post("/api/footnote/search", apiHandler("searchFootnotes"));

// Comments
app.get("/api/comments", apiHandler("getComments"));
app.post("/api/comment", apiHandler("addComment"));

// Undo
app.post("/api/undo", apiHandler("undo"));
app.get("/api/undo/history", apiHandler("undoHistory"));

// Advanced
app.post("/api/track-changes", apiHandler("trackChanges"));
app.post("/api/batch", apiHandler("batch"));

// ── Revert Endpoint ──
app.post("/api/revert/:exchangeId", async (req: express.Request, res: express.Response) => {
  const exchangeId = parseInt(req.params.exchangeId, 10);
  if (!Number.isFinite(exchangeId)) return res.status(400).json({ ok: false, error: "valid exchangeId required" });

  // Collect undo counts for this exchange and all subsequent ones
  let totalUndos = 0;
  const toRemove: number[] = [];
  for (const [eid, count] of exchangeUndoCounts) {
    if (eid >= exchangeId) {
      totalUndos += count;
      toRemove.push(eid);
    }
  }

  if (totalUndos === 0) return res.status(404).json({ ok: false, error: "No changes to revert for this exchange" });

  try {
    for (let i = 0; i < totalUndos; i++) {
      await sendCommand("undo");
    }
    for (const eid of toRemove) exchangeUndoCounts.delete(eid);
    res.json({ ok: true, data: { exchangeId, undoCount: totalUndos, revertedExchanges: toRemove } });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// Start
// ── Session Management ──
app.post("/api/session/new", (_req, res) => {
  sessionId = "sidebar-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8);
  conversationHistory = [];
  console.log("[session] New session:", sessionId);
  res.json({ ok: true, data: { sessionId } });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`\n  🎀 The Sidebar Server v${VERSION}`);
  console.log(`  ${"🌐 HTTP (localhost)"} on port ${PORT}`);
  console.log(`  📡 WebSocket waiting for connections...\n`);
});

// ── Settings ──
app.get("/api/settings", handleGetSettings());
app.post("/api/settings", handlePostSettings());

// ── OpenClaw Connection Test ──
app.post("/api/openclaw/test", async (req, res) => {
  const { readConfig } = require("./settings");
  const config = readConfig();
  const url = req.body?.url || config.openclawUrl;
  if (!url) return res.status(400).json({ ok: false, error: "No OpenClaw URL configured" });
  try {
    const { httpRequest } = require("./llm-router");
    const testUrl = url.replace(/\/$/, "") + "/v1/chat/completions";
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    if (config.openclawToken) headers["Authorization"] = `Bearer ${config.openclawToken}`;
    const response = await httpRequest(testUrl, { method: "POST", headers },
      { model: "openclaw:main", messages: [{ role: "user", content: "ping" }], max_tokens: 1 });
    let body = "";
    for await (const chunk of response) body += chunk.toString();
    if (response.statusCode && response.statusCode < 400) {
      res.json({ ok: true, data: { status: response.statusCode, url } });
    } else {
      res.json({ ok: false, error: `HTTP ${response.statusCode}: ${body.slice(0, 200)}` });
    }
  } catch (e: any) {
    res.json({ ok: false, error: e.message });
  }
});
// ── PDF Export (via AppleScript, Mac only) ──
app.post("/api/export/pdf", async (req, res) => {
  const outPath = req.body?.path || "/tmp/thesidebar-export.pdf";
  const { execSync } = require("child_process");
  try {
    // Use AppleScript to print to PDF
    execSync(`osascript -e '
      tell application "Microsoft Word"
        activate
        delay 0.5
      end tell
      tell application "System Events"
        tell process "Microsoft Word"
          keystroke "p" using command down
          delay 2
          try
            click menu button "PDF" of sheet 1 of window 1
            delay 0.5
            click menu item "Save as PDF…" of menu 1 of menu button "PDF" of sheet 1 of window 1
            delay 1
            keystroke "a" using command down
            keystroke "${outPath}"
            delay 0.5
            keystroke return
            delay 3
            -- Close print dialog if still open
          end try
        end tell
      end tell'`, { timeout: 30000 });
    // Verify file exists
    if (fs.existsSync(outPath)) {
      res.json({ ok: true, data: { path: outPath, size: fs.statSync(outPath).size } });
    } else {
      res.json({ ok: false, error: "PDF export may have failed — file not found. Check Word for dialogs." });
    }
  } catch (e: any) {
    res.status(500).json({ ok: false, error: `PDF export failed: ${e.message}` });
  }
});

// ── Section Read (by heading text or index range) ──
app.get("/api/section", apiHandler("getSection", (req) => ({
  heading: req.query.heading as string,
  headingIndex: req.query.headingIndex ? parseInt(req.query.headingIndex as string, 10) : undefined,
})));

// ── Bulk Paragraph Read (specific indices) ──
app.post("/api/paragraphs/bulk", apiHandler("getBulkParagraphs"));

// ── Document Properties ──
app.get("/api/document/properties", apiHandler("getDocumentProperties"));

// ── Diff (compare current paragraph to provided text) ──
app.post("/api/paragraph/diff", apiHandler("diffParagraph"));
