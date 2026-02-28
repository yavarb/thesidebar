import { parsePageContent, checkToaEntries } from "./toa-checker";
import { setTrackChanges, isTrackChangesEnabled } from "./track-changes";
import { runAgentLoop } from "./agent-loop";
import { v4 as uuidv4 } from "uuid";
import { saveSession, loadSession, deleteSession, cleanExpiredSessions, ensureMachineKey, SessionData, generateSessionRecap, searchConversationHistory } from "./sessions";
import { getContextSize, manageContext } from "./context";
import { resolveModel, cacheStats } from "./llm-router";
import { readConfig } from "./settings";
import { handleGetSettings, handlePostSettings } from "./settings";
import { queryDocuments, listDocuments as listRefDocs, getStatus as getRefStatus, rescan as rescanRefs, startPeriodicScan, stopPeriodicScan, getDocumentCount as getRefDocCount } from "./references";
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
let conversationHistory: { role: string; content: string; timestamp?: number }[] = [];
let currentSessionData: SessionData | null = null;
// ── In-Memory Document Index ──
let documentIndex: { paragraphs: {index: number, text: string, listString?: string}[], builtAt: number, hash: string } | null = null;





async function buildDocumentIndex(): Promise<typeof documentIndex> {
  try {
    const result = await sendCommand("getParagraphs", { compact: true });
    const paragraphs = (result?.paragraphs || []).map((p: any) => ({
      index: p.index,
      text: p.text,
      listString: p.listString,
    }));
    const raw = paragraphs.map((p: any) => p.text).join("");
    const hash = raw.substring(0, 100) + "|" + paragraphs.length;
    documentIndex = { paragraphs, builtAt: Date.now(), hash };
    console.log(`[index] Built document index: ${paragraphs.length} paragraphs`);
    return documentIndex;
  } catch (e: any) {
    console.error("[index] Failed to build:", e.message);
    return null;
  }
}

async function getDocumentContext(): Promise<string> {
  if (documentIndex && (Date.now() - documentIndex.builtAt) < 30000) {
    return documentIndex.paragraphs.map(p => p.text).join("\n");
  }
  await buildDocumentIndex();
  return documentIndex ? documentIndex.paragraphs.map(p => p.text).join("\n") : "";
}

function invalidateIndex(): void {
  if (documentIndex) {
    documentIndex.builtAt = 0; // Mark as stale
  }
}

/** Single task pane WebSocket connection (localhost only) */
let taskPaneWs: WebSocket | null = null;
let globalRequestId = 0;
const pending = new Map<number, { resolve: (v: any) => void; reject: (e: any) => void; timer: NodeJS.Timeout; command: string }>();

// Prompt queue
type PromptEntry = { id: number; text: string; model?: string; context?: string; timestamp: number; clientId?: string };
const promptQueue: PromptEntry[] = [];
let currentAbortController: AbortController | null = null;
const promptLog = new Map<number, PromptEntry>();
let promptId = 0;
const promptWaiters: { resolve: (v: any) => void; timer: NodeJS.Timeout }[] = [];

// ── Change Tracking & Revert ──
const MODIFYING_TOOLS = new Set([
  "updateParagraph", "replaceParagraph", "replaceSelection", "editSelection",
  "findReplace", "insert", "format", "setStyleFont", "addFootnote",
  "updateFootnote", "addComment", "batch", "deleteParagraph",
  "insertTable", "updateTableCell", "addTableRow", "addTableColumn",
  "setHeaderFooter", "insertBreak", "setListFormat", "highlightText",
  "setFontColor", "setParagraphFormat", "acceptTrackedChange", "rejectTrackedChange",
  "applyStyle", "createStyle", "modifyStyle", "deleteFootnote",
  "insertFootnoteWithFormat", "markCitation", "insertTableOfAuthorities",
  "insertCrossReference",
  "setPageSetup",
]);
/** Maps exchangeId → number of modifying tool calls made during that exchange */
const exchangeUndoCounts: Map<number, number> = new Map();
let exchangeIdCounter = 0;


// ── Auto-process prompts via agent loop ──
// ── Tool Progress Messages ──
const TOOL_PROGRESS: Record<string, (args: any) => string> = {
  readDocument: () => "📖 Reading document...",
  readParagraphs: (a: any) => `📖 Reading paragraphs ${a.from || ''}–${a.to || ''}...`,
  getParagraph: (a: any) => `📖 Reading paragraph ${a.index}...`,
  getParagraphs: () => "📖 Reading paragraphs...",
  updateParagraph: (a: any) => `✏️ Editing paragraph ${a.index}...`,
  replaceParagraph: (a: any) => `✏️ Editing paragraph ${a.index || a.listString || ''}...`,
  replaceSelection: () => "✏️ Replacing selection...",
  find: (a: any) => `🔍 Searching for "${(a.text || '').slice(0, 30)}..."`,
  findReplace: (a: any) => `🔍 Replacing "${(a.search || a.text || '').slice(0, 20)}" → "${(a.replace || a.replacement || '').slice(0, 20)}"...`,
  insert: (a: any) => `📝 Inserting at paragraph ${a.paragraphIndex || a.index || 'end'}...`,
  addFootnote: () => "📝 Adding footnote...",
  updateFootnote: (a: any) => `📝 Updating footnote ${a.index}...`,
  deleteFootnote: (a: any) => `🗑️ Deleting footnote ${a.index}...`,
  getFootnoteBody: (a: any) => `📖 Reading footnote ${a.index}...`,
  insertFootnoteWithFormat: () => "📝 Inserting formatted footnote...",
  reorderFootnotes: () => "📖 Listing footnotes with locations...",
  format: () => "🎨 Formatting...",
  setStyleFont: () => "🎨 Updating style font...",
  addComment: () => "💬 Adding comment...",
  getDocumentStructure: () => "📋 Analyzing structure...",
  getDocumentStats: () => "📋 Getting document stats...",
  batch: () => "⚡ Executing batch operations...",
  insertTable: () => "📊 Creating table...",
  readTable: (a: any) => `📊 Reading table ${a.index}...`,
  updateTableCell: (a: any) => `📊 Updating cell [${a.row},${a.column}]...`,
  addTableRow: () => "📊 Adding table row...",
  addTableColumn: () => "📊 Adding table column...",
  getTables: () => "📊 Listing tables...",
  getHeaderFooter: () => "📖 Reading header/footer...",
  setHeaderFooter: () => "📝 Setting header/footer...",
  deleteParagraph: (a: any) => `🗑️ Deleting paragraph ${a.index}...`,
  insertBreak: () => "📄 Inserting break...",
  setListFormat: () => "📝 Setting list format...",
  getBookmarks: () => "🔖 Listing bookmarks...",
  highlightText: (a: any) => `🖍️ Highlighting "${(a.text || '').slice(0, 20)}"...`,
  setFontColor: () => "🎨 Setting font color...",
  setParagraphFormat: (a: any) => `📐 Formatting paragraph ${a.index}...`,
  getTrackedChanges: () => "📋 Listing tracked changes...",
  acceptTrackedChange: () => "✅ Accepting tracked change...",
  rejectTrackedChange: () => "❌ Rejecting tracked change...",
  applyStyle: (a: any) => `🎨 Applying style "${a.styleName}"...`,
  createStyle: (a: any) => `🎨 Creating style "${a.name}"...`,
  modifyStyle: (a: any) => `🎨 Modifying style "${a.styleName}"...`,
  getStyleDetails: (a: any) => `📋 Getting style details for "${a.styleName}"...`,
  markCitation: (a: any) => `⚖️ Marking citation: ${(a.shortCite || '').slice(0, 30)}...`,
  insertTableOfAuthorities: () => "⚖️ Inserting Table of Authorities...",
  insertCrossReference: (a: any) => `🔗 Inserting cross-reference to ${a.type} "${(a.target || '').slice(0, 30)}"...`,
  validateCrossReferences: () => "🔍 Validating cross-references...",
  checkToaPages: () => "⚖️ Checking TOA pages (exporting PDF)...",
  getPageSetup: () => "📐 Reading page setup...",
  setPageSetup: () => "📐 Adjusting page margins...",
  getPageNumbers: () => "📋 Getting page info...",
};

async function processPrompt(entry: PromptEntry, ws: any) {
  const abortController = new AbortController();
  currentAbortController = abortController;
  try {
    // Determine model early (needed for context strategy)
    const config = readConfig();
    const model = entry.model || config.defaultModel || "openclaw";

    // Get document context (from cache if fresh)
    let documentContext = "";
    try {
      if (!model.startsWith("openai:") && !model.startsWith("anthropic:") && !model.startsWith("local:")) {
        // OpenClaw needs full doc since it uses curl (can't call tools natively)
        documentContext = await getDocumentContext();
      } else {
        // Direct API models have native tool access — send compact summary, let them read on demand
        const stats = await sendCommand("getDocumentStats", {});
        const structure = await sendCommand("getDocumentStructure", {});
        documentContext = `Document: ${stats?.wordCount || "?"} words, ${stats?.paragraphCount || "?"} paragraphs, ${stats?.footnoteCount || "?"} footnotes, ${stats?.sectionCount || "?"} sections.\n`;
        if (structure?.headings?.length) {
          documentContext += "Structure:\n" + structure.headings.map((h: any) => `${"  ".repeat((h.level || 1) - 1)}${h.text}`).join("\n");
        }
        documentContext += "\n\nUse readDocument, readParagraph, or readParagraphs tools to read specific content as needed. Do NOT ask the user what to read — just read it.";
      }
    } catch (e: any) {
      console.error("[agent] Failed to get document:", e.message);
    }

    const routerConfig = {
      openclawUrl: config.openclawUrl,
      openclawToken: config.openclawToken,
      openaiApiKey: config.openaiApiKey,
      anthropicApiKey: config.anthropicApiKey,
      localEndpoints: config.localEndpoints,
    };

    // ── Reference Documents (RAG) ──
    const isOpenClawModel = !model.startsWith("openai:") && !model.startsWith("anthropic:") && !model.startsWith("local:");
    const referenceFolders = config.referenceFolders || [];
    let referenceContext = "";

    if (referenceFolders.length > 0 && !isOpenClawModel) {
      // RAG retrieval for direct API models
      try {
        const refResults = await queryDocuments(entry.text, 5);
        if (refResults.length > 0) {
          referenceContext = "\n\n## Reference Documents\n\n";
          for (const r of refResults) {
            referenceContext += `From "${r.filename}" (relevance: ${r.score.toFixed(2)}):\n${r.chunkText}\n\n`;
          }
        }
      } catch (e: any) {
        console.error("[agent] Reference query failed:", e.message);
      }
    }

    if (referenceFolders.length > 0 && referenceContext) {
      documentContext += referenceContext;
    }

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
    let changeSummaries: string[] = [];
    // Build systemPrompt addendum for OpenClaw with reference folders
    let systemPromptOverride: string | undefined;
    if (isOpenClawModel && referenceFolders.length > 0) {
      let folderHint = "\n\nThe user has designated the following reference folders for this case:\n";
      for (const folder of referenceFolders) {
        folderHint += `- ${folder}\n`;
      }
      folderHint += "\nPrioritize these folders when searching for documents, exhibits, or supporting materials. You have full filesystem access — use it to read relevant files directly when the user\'s question relates to other documents.";
      systemPromptOverride = `You are The Sidebar, an AI assistant embedded inside Microsoft Word via a task pane add-in. You are connected to the CURRENTLY OPEN Word document.

CRITICAL RULES:
1. The OPEN Word document must be read and edited ONLY through The Sidebar's HTTP API at http://localhost:3001. Do NOT use filesystem tools (python-docx, read, write, edit) to read or modify the .docx file. It is live in Word — you control it through HTTP calls.
2. You DO have full filesystem access for everything else: reading reference documents, exhibits, research files, case folders, and any other supporting materials.
3. The document context provided with each prompt reflects the CURRENT state of the open document.

HOW TO USE THE SIDEBAR TOOLS:
You are connected via OpenClaw. Use your exec tool to run curl commands against the API. Do NOT try to call these as native tool functions — they are HTTP endpoints. Examples:

Read the full document:
  curl -s http://localhost:3001/api/document

Read a specific paragraph (index 5):
  curl -s http://localhost:3001/api/paragraph?index=5

Replace a paragraph:
  curl -s -X POST http://localhost:3001/api/paragraph/replace -H "Content-Type: application/json" -d '{"index": 5, "text": "New paragraph text here"}'

Find and replace:
  curl -s -X POST http://localhost:3001/api/find-replace -H "Content-Type: application/json" -d '{"find": "old text", "replace": "new text"}'

Insert text:
  curl -s -X POST http://localhost:3001/api/insert -H "Content-Type: application/json" -d '{"text": "New text", "location": "end"}'

Add a footnote:
  curl -s -X POST http://localhost:3001/api/footnote -H "Content-Type: application/json" -d '{"paragraphIndex": 12, "text": "Footnote text"}'

Read footnotes:
  curl -s http://localhost:3001/api/footnotes

Get document structure:
  curl -s http://localhost:3001/api/document/structure

IMPORTANT: Do NOT create new versions of the document. Do NOT use python-docx. Do NOT save files to disk. ALL edits go through the API above which modifies the document live in Word.

BEHAVIORAL RULE: When the user asks you to check, fix, or correct something in the document, MAKE THE CHANGES YOURSELF. Do not tell the user to do it manually. Do not suggest they "update fields", "press F9", "regenerate the TOA", or perform any manual steps. You have full editing capability — use it. Report what you found AND fix it.

Available tools (HTTP endpoints at http://localhost:3001/api/):
- READ: readDocument, readParagraph, readParagraphs, readSelection, getDocumentStats, getStructure, getToc, getDocumentProperties, getStyles, getStyleDetails, getBookmarks
- EDIT: replaceParagraph, editSelection, insertText, findReplace, find, deleteParagraph, batch, undo
- FORMAT: formatParagraph, setParagraphFormat, applyStyle, createStyle, modifyStyle, highlightText, setFontColor, setListFormat, insertBreak
- FOOTNOTES: addFootnote, readFootnotes, getFootnoteBody, updateFootnote, deleteFootnote, insertFootnoteWithFormat
- COMMENTS: addComment, getComments
- TABLES: insertTable, readTable, getTables, updateTableCell, addTableRow, addTableColumn
- HEADERS: getHeaderFooter, setHeaderFooter
- PAGE: getPageSetup, setPageSetup
- NAVIGATION: navigateTo, selectParagraph
- TRACKING: getTrackedChanges, acceptTrackedChange, rejectTrackedChange
- CITATIONS: markCitation, insertTableOfAuthorities, insertCrossReference, validateCrossReferences, checkToaPages

To call a tool, make an HTTP request (GET or POST) to http://localhost:3001/api/<endpoint> with JSON body parameters.` + folderHint;
    }


    // Inject session recap if available
    if (currentSessionData?.lastRecap) {
      const recapAddendum = "\n\n## Previous Session Context\n" + currentSessionData.lastRecap + "\n\nThe full conversation history is available. The user may reference prior discussions.";
      if (systemPromptOverride) {
        systemPromptOverride += recapAddendum;
      } else {
        systemPromptOverride = `You are The Sidebar, an AI assistant embedded inside Microsoft Word via a task pane add-in. You are connected to the CURRENTLY OPEN Word document.

CRITICAL RULES:
1. The OPEN Word document must be read and edited ONLY through The Sidebar's HTTP API at http://localhost:3001. Do NOT use filesystem tools (python-docx, read, write, edit) to read or modify the .docx file. It is live in Word — you control it through HTTP calls.
2. You DO have full filesystem access for everything else: reading reference documents, exhibits, research files, case folders, and any other supporting materials.
3. The document context provided with each prompt reflects the CURRENT state of the open document.

HOW TO USE THE SIDEBAR TOOLS:
You are connected via OpenClaw. Use your exec tool to run curl commands against the API. Do NOT try to call these as native tool functions — they are HTTP endpoints. Examples:

Read the full document:
  curl -s http://localhost:3001/api/document

Read a specific paragraph (index 5):
  curl -s http://localhost:3001/api/paragraph?index=5

Replace a paragraph:
  curl -s -X POST http://localhost:3001/api/paragraph/replace -H "Content-Type: application/json" -d '{"index": 5, "text": "New paragraph text here"}'

Find and replace:
  curl -s -X POST http://localhost:3001/api/find-replace -H "Content-Type: application/json" -d '{"find": "old text", "replace": "new text"}'

Insert text:
  curl -s -X POST http://localhost:3001/api/insert -H "Content-Type: application/json" -d '{"text": "New text", "location": "end"}'

Add a footnote:
  curl -s -X POST http://localhost:3001/api/footnote -H "Content-Type: application/json" -d '{"paragraphIndex": 12, "text": "Footnote text"}'

Read footnotes:
  curl -s http://localhost:3001/api/footnotes

Get document structure:
  curl -s http://localhost:3001/api/document/structure

IMPORTANT: Do NOT create new versions of the document. Do NOT use python-docx. Do NOT save files to disk. ALL edits go through the API above which modifies the document live in Word.

BEHAVIORAL RULE: When the user asks you to check, fix, or correct something in the document, MAKE THE CHANGES YOURSELF. Do not tell the user to do it manually. Do not suggest they "update fields", "press F9", "regenerate the TOA", or perform any manual steps. You have full editing capability — use it. Report what you found AND fix it.

Available tools (HTTP endpoints at http://localhost:3001/api/):
- READ: readDocument, readParagraph, readParagraphs, readSelection, getDocumentStats, getStructure, getToc, getDocumentProperties, getStyles, getStyleDetails, getBookmarks
- EDIT: replaceParagraph, editSelection, insertText, findReplace, find, deleteParagraph, batch, undo
- FORMAT: formatParagraph, setParagraphFormat, applyStyle, createStyle, modifyStyle, highlightText, setFontColor, setListFormat, insertBreak
- FOOTNOTES: addFootnote, readFootnotes, getFootnoteBody, updateFootnote, deleteFootnote, insertFootnoteWithFormat
- COMMENTS: addComment, getComments
- TABLES: insertTable, readTable, getTables, updateTableCell, addTableRow, addTableColumn
- HEADERS: getHeaderFooter, setHeaderFooter
- PAGE: getPageSetup, setPageSetup
- NAVIGATION: navigateTo, selectParagraph
- TRACKING: getTrackedChanges, acceptTrackedChange, rejectTrackedChange
- CITATIONS: markCitation, insertTableOfAuthorities, insertCrossReference, validateCrossReferences, checkToaPages

To call a tool, make an HTTP request (GET or POST) to http://localhost:3001/api/<endpoint> with JSON body parameters.` + recapAddendum;
      }
    }
    for await (const chunk of runAgentLoop({
      prompt: fullPrompt,
      model,
      config: routerConfig,
      documentContext,
      systemPrompt: systemPromptOverride,
      sessionUser: isOpenClaw ? "thesidebar:" + sessionId : undefined,
      conversationHistory: !isOpenClaw ? managedHistory : undefined,
      onToolCall: (call) => {
        if (MODIFYING_TOOLS.has(call.name)) {
          modifyingCallCount++;
          invalidateIndex();
        }
        (call as any)._startTime = Date.now();
        if (ws.readyState === 1) {
          const progressFn = TOOL_PROGRESS[call.name];
          const progressText = progressFn ? progressFn(call.arguments) : `Using tool: ${call.name}...`;
          ws.send(JSON.stringify({ type: "prompt_progress", promptId: entry.id, status: "tool", toolName: call.name, progressText }));
        }
      },
      onToolResult: (result) => {
        if (ws.readyState === 1) {
          ws.send(JSON.stringify({ type: "prompt_progress", promptId: entry.id, status: "tool_complete", toolName: result.name }));
        }
      },
    })) {
      // Check for change summaries marker
      if (chunk.startsWith("\n__CHANGE_SUMMARIES__")) {
        try {
          changeSummaries = JSON.parse(chunk.substring("\n__CHANGE_SUMMARIES__".length));
        } catch {}
        continue;
      }
      // Skip internal control markers
      if (chunk === "\n__TOOL_EXEC_START__") continue;
      if (chunk.startsWith("\n__TOOL_PHASE__")) {
        try {
          const info = JSON.parse(chunk.substring("\n__TOOL_PHASE__".length));
          ws.send(JSON.stringify({ type: "prompt_progress", promptId: entry.id, status: "tool_phase", toolCount: info.toolCount, tools: info.tools }));
        } catch {}
        continue;
      }
      fullResponse += chunk;
      if (ws.readyState === 1) {
        ws.send(JSON.stringify({ type: "prompt_progress", promptId: entry.id, text: fullResponse }));
      }
    }

    // Track conversation history for non-OpenClaw models
    conversationHistory.push({ role: "user", content: fullPrompt, timestamp: Date.now() });
    conversationHistory.push({ role: "assistant", content: fullResponse || "(No response)", timestamp: Date.now() });

    // Auto-save session
    if (currentSessionData) {
      currentSessionData.conversationHistory = [...conversationHistory];
      currentSessionData.model = model;
      // Store change summaries
      if (changeSummaries.length > 0) {
        if (!currentSessionData.changeSummaries) currentSessionData.changeSummaries = [];
        for (const s of changeSummaries) {
          currentSessionData.changeSummaries.push({ exchangeIndex: currentExchangeId, summary: s });
        }
      }
      saveSession(currentSessionData);
    }

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

    // Clear abort controller
    if (currentAbortController === abortController) currentAbortController = null;

    // Send final response
    if (ws.readyState === 1) {
      ws.send(JSON.stringify({ type: "prompt_response", promptId: entry.id, text: fullResponse || "(No response)", timestamp: Date.now(), hasChanges, exchangeId: currentExchangeId, changeSummaries: changeSummaries.length > 0 ? changeSummaries : undefined }));
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

  // Auto-build document index when task pane connects
  buildDocumentIndex().catch(e => console.error("[index] Auto-build failed:", e.message));

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
  stopPeriodicScan();
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

// Cache stats endpoint
app.get("/api/cache/stats", (_req, res) => {
  res.json({ ok: true, data: cacheStats });
});
app.post("/api/cache/stats/reset", (_req, res) => {
  cacheStats.hits = 0;
  cacheStats.misses = 0;
  cacheStats.anthropicCacheReadTokens = 0;
  cacheStats.anthropicCacheCreationTokens = 0;
  cacheStats.openaiCachedTokens = 0;
  res.json({ ok: true, data: { reset: true } });
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
    { m: "GET", p: "/api/references", d: "List reference docs + status" },
    { m: "GET", p: "/api/references/status", d: "Reference index status" },
    { m: "POST", p: "/api/references/rescan", d: "Trigger folder rescan" },
    { m: "POST", p: "/api/references/query", d: "Query reference chunks" },
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
app.post("/api/prompts/abort", (_req, res) => {
  if (currentAbortController) {
    currentAbortController.abort();
    currentAbortController = null;
    // Also clear queue
    const c = promptQueue.length;
    promptQueue.length = 0;
    res.json({ ok: true, data: { aborted: true, queueCleared: c } });
  } else {
    const c = promptQueue.length;
    promptQueue.length = 0;
    res.json({ ok: true, data: { aborted: false, queueCleared: c } });
  }
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

// Force refresh server-side document index
app.post("/api/index/refresh", async (_req, res) => {
  try {
    const idx = await buildDocumentIndex();
    res.json({ ok: true, data: { paragraphCount: idx?.paragraphs.length || 0, hash: idx?.hash || "" } });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

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

// ── Tables ──
app.get("/api/tables", apiHandler("getTables"));
app.get("/api/table/:index", apiHandler("readTable", (req) => ({ index: parseInt(req.params.index, 10) })));
app.post("/api/table/insert", apiHandler("insertTable"));
app.post("/api/table/cell", apiHandler("updateTableCell"));
app.post("/api/table/row", apiHandler("addTableRow"));
app.post("/api/table/column", apiHandler("addTableColumn"));

// ── Headers & Footers ──
app.post("/api/header-footer", apiHandler("getHeaderFooter"));
app.post("/api/header-footer/set", apiHandler("setHeaderFooter"));

// ── Paragraph Operations ──
app.post("/api/paragraph/delete", apiHandler("deleteParagraph"));
app.post("/api/paragraph/format", apiHandler("setParagraphFormat"));

// ── Breaks ──
app.post("/api/break", apiHandler("insertBreak"));

// ── Lists ──
app.post("/api/list-format", apiHandler("setListFormat"));

// ── Bookmarks ──
app.get("/api/bookmarks", apiHandler("getBookmarks"));

// ── Highlight & Font Color ──
app.post("/api/highlight", apiHandler("highlightText"));
app.post("/api/font-color", apiHandler("setFontColor"));

// ── Tracked Changes ──
app.get("/api/tracked-changes", apiHandler("getTrackedChanges"));
app.post("/api/tracked-changes/accept", apiHandler("acceptTrackedChange"));
app.post("/api/tracked-changes/reject", apiHandler("rejectTrackedChange"));

// ── Style Operations ──
app.post("/api/style/apply", apiHandler("applyStyle"));
app.post("/api/style/create", apiHandler("createStyle"));
app.post("/api/style/modify", apiHandler("modifyStyle"));
app.post("/api/style/details", apiHandler("getStyleDetails"));

// ── Footnotes (expanded) ──
app.post("/api/footnote/delete", apiHandler("deleteFootnote"));
app.get("/api/footnote/:index/body", apiHandler("getFootnoteBody", (req) => ({ index: parseInt(req.params.index, 10) })));
app.post("/api/footnote/insert", apiHandler("insertFootnoteWithFormat"));
app.get("/api/footnotes/detailed", apiHandler("reorderFootnotes"));

// ── Citations / Table of Authorities ──
app.post("/api/citation/mark", apiHandler("markCitation"));
app.post("/api/citation/toa", apiHandler("insertTableOfAuthorities"));

// ── Cross-References ──
app.post("/api/cross-reference", apiHandler("insertCrossReference"));
app.get("/api/cross-references/validate", apiHandler("validateCrossReferences"));

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
  const newId = uuidv4();
  conversationHistory = [];
  const config = readConfig();
  currentSessionData = {
    sessionId: newId,
    conversationHistory: [],
    model: config.defaultModel || "openclaw",
    createdAt: Date.now(),
    updatedAt: Date.now(),
    documentName: undefined,
  };
  saveSession(currentSessionData);
  sessionId = "sidebar-" + newId; // Keep OpenClaw session prefix
  console.log("[session] New session:", newId);
  res.json({ ok: true, data: { sessionId: newId } });
});

app.post("/api/session/resume", (req, res) => {
  const { sessionId: sid } = req.body || {};
  if (!sid) return res.status(400).json({ ok: false, error: "sessionId required" });
  const data = loadSession(sid);
  if (!data) return res.json({ ok: true, data: { found: false } });
  currentSessionData = data;
  conversationHistory = [...data.conversationHistory];
  sessionId = "sidebar-" + sid;
  const recap = generateSessionRecap(data);
  data.lastRecap = recap;
  saveSession(data);
  console.log("[session] Resumed session:", sid, "("+data.conversationHistory.length+" messages)");
  res.json({ ok: true, data: { found: true, conversationHistory: data.conversationHistory, model: data.model, recap } });
});




app.post("/api/session/search", (req, res) => {
  const { query } = req.body || {};
  if (!query) return res.status(400).json({ ok: false, error: "query required" });
  if (!currentSessionData) return res.json({ ok: true, data: { results: [], message: "No active session" } });
  const results = searchConversationHistory(currentSessionData, query);
  res.json({ ok: true, data: { results, count: results.length } });
});

app.post("/api/session/delete", (_req, res) => {
  if (currentSessionData) {
    deleteSession(currentSessionData.sessionId);
    console.log("[session] Deleted session:", currentSessionData.sessionId);
  }
  conversationHistory = [];
  currentSessionData = null;
  res.json({ ok: true });
});

app.post("/api/sessions/purge", (_req, res) => {
  const sessionsDir = require("path").join(process.env.HOME || "~", ".thesidebar", "sessions");
  let count = 0;
  try {
    const files = require("fs").readdirSync(sessionsDir);
    for (const f of files) {
      if (f.endsWith(".enc")) {
        require("fs").unlinkSync(require("path").join(sessionsDir, f));
        count++;
      }
    }
  } catch {}
  conversationHistory = [];
  currentSessionData = null;
  console.log("[session] Purged " + count + " sessions");
  res.json({ ok: true, data: { purged: count } });
});

// ── Session Init ──
ensureMachineKey();

// Periodic session cleanup
const config_init = readConfig();
const ttl = config_init.sessionTTLDays ?? 30;
cleanExpiredSessions(ttl);
setInterval(() => {
  const c = readConfig();
  cleanExpiredSessions(c.sessionTTLDays ?? 30);
}, 3600000); // every hour

server.listen(PORT, "127.0.0.1", () => {
  console.log(`\n  🎀 The Sidebar Server v${VERSION}`);
  console.log(`  ${"🌐 HTTP (localhost)"} on port ${PORT}`);
  console.log(`  📡 WebSocket waiting for connections...\n`);

  // Start reference document folder scanning
  startPeriodicScan();
});

// ── Settings ──
app.get("/api/settings", handleGetSettings());
app.post("/api/settings", (req, res, next) => {
  // Wrap handlePostSettings to trigger rescan when referenceFolders change
  const handler = handlePostSettings();
  handler(req, res);
  if (req.body?.referenceFolders !== undefined) {
    rescanRefs().catch(e => console.error("[references] Rescan after settings change failed:", e.message));
  }
});


app.post("/api/settings/mode", (req, res) => {
  setTrackChanges(req.body?.trackChanges ?? false);
  res.json({ ok: true, data: { trackChanges: isTrackChangesEnabled() } });
});
app.get("/api/settings/mode", (_req, res) => {
  res.json({ ok: true, data: { trackChanges: isTrackChangesEnabled() } });
});

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

// ── Reference Documents ──
app.get("/api/references", (_req, res) => {
  res.json({ ok: true, data: { documents: listRefDocs(), status: getRefStatus() } });
});

app.get("/api/references/status", (_req, res) => {
  res.json({ ok: true, data: getRefStatus() });
});

app.post("/api/references/rescan", async (_req, res) => {
  try {
    const result = await rescanRefs();
    res.json({ ok: true, data: result });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.post("/api/references/query", async (req, res) => {
  const { text, topK } = req.body || {};
  if (!text) return res.status(400).json({ ok: false, error: "text required" });
  try {
    const results = await queryDocuments(text, topK || 5);
    res.json({ ok: true, data: results });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── Custom Prompts Persistence ──
const CUSTOM_PROMPTS_PATH = path.join(process.env.HOME || "~", ".thesidebar", "custom-prompts.json");

app.get("/api/prompts/custom", (_req, res) => {
  try {
    if (fs.existsSync(CUSTOM_PROMPTS_PATH)) {
      const data = JSON.parse(fs.readFileSync(CUSTOM_PROMPTS_PATH, "utf-8"));
      res.json({ ok: true, data: Array.isArray(data) ? data : [] });
    } else {
      res.json({ ok: true, data: [] });
    }
  } catch (e: any) {
    res.json({ ok: true, data: [] });
  }
});

app.post("/api/prompts/custom", (req, res) => {
  try {
    const prompts = req.body;
    if (!Array.isArray(prompts)) return res.status(400).json({ ok: false, error: "array required" });
    const dir = path.dirname(CUSTOM_PROMPTS_PATH);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(CUSTOM_PROMPTS_PATH, JSON.stringify(prompts, null, 2));
    res.json({ ok: true });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e.message });
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

// ── TOA Page Check ──
app.post("/api/toa/check", async (_req, res) => {
  try {
    // Step 1: Get TOA entries from document
    const toaResult = await sendCommand("getToaEntries", {});
    if (!toaResult.entries?.length) return res.json({ ok: false, error: "No TOA entries found in document. Make sure there is a Table of Authorities section." });

    // Step 2: Export PDF and parse page-to-text mapping
    const pdfResult = await sendCommand("exportPdf", {});
    if (pdfResult.error || !pdfResult.pdf) return res.json({ ok: false, error: "PDF export failed: " + (pdfResult.error || "no data") });

    const pages = await parsePageContent(pdfResult.pdf);
    if (!pages.length) return res.json({ ok: false, error: "Could not extract any pages from PDF" });

    // Step 3: Build page map summary for LLM (text snippet per page)
    const pageMap = pages.map(p => ({
      page: p.pageNumber,
      // Truncate each page to ~2000 chars to keep context manageable
      text: p.text.substring(0, 2000) + (p.text.length > 2000 ? "..." : "")
    }));

    // Step 4: Return everything — let the caller send to LLM
    res.json({
      ok: true,
      data: {
        toaEntries: toaResult.entries,
        pageCount: pages.length,
        pageMap,
      }
    });
  } catch (e: any) { res.status(500).json({ ok: false, error: e.message }); }
});


// ── Model Discovery ──
app.get("/api/models/openai", async (_req, res) => {
  try {
    const cfg = readConfig();
    const key = cfg.openaiApiKey;
    if (!key) return res.json({ ok: false, error: "No OpenAI API key configured" });
    const resp = await fetch("https://api.openai.com/v1/models", {
      headers: { "Authorization": `Bearer ${key}` }
    });
    const data = await resp.json();
    // Filter to chat models only, sort by id
    const chatModels = ((data as any).data || [])
      .filter((m: any) => m.id && !m.id.includes("embedding") && !m.id.includes("whisper") && !m.id.includes("tts") && !m.id.includes("dall-e") && !m.id.includes("moderation"))
      .sort((a: any, b: any) => a.id.localeCompare(b.id));
    res.json({ ok: true, data: chatModels });
  } catch (e: any) { res.json({ ok: false, error: e.message }); }
});

app.get("/api/models/anthropic", async (_req, res) => {
  try {
    const cfg = readConfig();
    const key = cfg.anthropicApiKey;
    if (!key) return res.json({ ok: false, error: "No Anthropic API key configured" });
    const resp = await fetch("https://api.anthropic.com/v1/models", {
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01" }
    });
    const data = await resp.json();
    const models = ((data as any).data || [])
      .sort((a: any, b: any) => (a.id || "").localeCompare(b.id || ""));
    res.json({ ok: true, data: models });
  } catch (e: any) { res.json({ ok: false, error: e.message }); }
});

app.get("/api/models/local", async (req, res) => {
  try {
    const baseUrl = req.query.baseUrl as string;
    if (!baseUrl) return res.json({ ok: false, error: "baseUrl required" });
    const resp = await fetch(`${baseUrl}/v1/models`);
    const data = await resp.json();
    res.json({ ok: true, data: (data as any).data || [] });
  } catch (e: any) { res.json({ ok: false, error: e.message }); }
});

// ── Page Setup ──
app.get("/api/page/setup", apiHandler("getPageSetup", (req) => ({
  sectionIndex: req.query.sectionIndex ? parseInt(req.query.sectionIndex as string, 10) : 0,
})));

app.post("/api/page/setup", apiHandler("setPageSetup"));

app.get("/api/page/info", apiHandler("getPageNumbers"));
