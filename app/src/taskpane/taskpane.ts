import { marked } from "marked";

// Configure marked for safe rendering
marked.setOptions({ breaks: true, gfm: true });

function renderMarkdown(text: string): string {
  try {
    return marked.parse(text) as string;
  } catch {
    return text;
  }
}

/* global Office, Word */

// ── Config ──
const WS_URL = `ws://${window.location.hostname}:3001`;
const RECONNECT_BASE = 500;
const RECONNECT_MAX = 10000;
const HEARTBEAT_TIMEOUT = 35000; // 3.5x server heartbeat interval

// ── State ──
let socket: WebSocket | null = null;
let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
let reconnectDelay = RECONNECT_BASE;
let heartbeatTimer: ReturnType<typeof setTimeout> | null = null;
let missedHeartbeats = 0;

// ── Undo Stack ──
interface UndoEntry {
  command: string;
  original: string;
  replacement: string;
  paragraphIndex?: number;
  style?: string;
  timestamp: number;
}
const undoStack: UndoEntry[] = [];
const MAX_UNDO = 30;
function pushUndo(entry: UndoEntry) {
  undoStack.push(entry);
  if (undoStack.length > MAX_UNDO) undoStack.shift();
}

// ── Document Index ──
interface ParagraphMeta {
  index: number;
  preview: string;
  style: string;
  isListItem: boolean;
  listItemLevel?: number;
  listString?: string;
  charCount: number;
  hash: number;
}
interface DocumentIndex {
  version: number;
  paragraphCount: number;
  paragraphs: ParagraphMeta[];
  headings: { index: number; text: string; level: number; style: string }[];
  builtAt: number;
}
let docIndex: DocumentIndex | null = null;

function simpleHash(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  return h;
}

// ── Timeout Wrapper ──
function withTimeout<T>(promise: Promise<T>, ms: number, label: string): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`Timeout (${ms}ms) on ${label}`)), ms);
    promise.then(v => { clearTimeout(timer); resolve(v); }, e => { clearTimeout(timer); reject(e); });
  });
}
const DEFAULT_TIMEOUT = 15000;

// ── Prompt Conversation UI State ──
const pendingPromptEls = new Map<string, HTMLElement>();
let selectedModel = "";
let currentDocSessionId: string | null = null;

function addThinkingIndicator(): HTMLElement {
  const history = document.getElementById("prompt-history")!;
  const el = document.createElement("div");
  el.className = "chat-thinking";
  el.id = "thinking-indicator";
  el.innerHTML = '<div class="dot"></div><div class="dot"></div><div class="dot"></div>';
  history.appendChild(el);
  history.scrollTop = history.scrollHeight;
  return el;
}

function removeThinkingIndicator(): void {
  const el = document.getElementById("thinking-indicator");
  if (el) el.remove();
}



function appendChatEntry(role: "user" | "assistant", text: string, status?: string): HTMLElement {
  const history = document.getElementById("prompt-history")!;
  const el = document.createElement("div");
  el.className = `chat-entry chat-${role}`;
  const roleEl = document.createElement("div");
  roleEl.className = "chat-role";
  roleEl.textContent = role === "user" ? "You" : "OpenClaw";
  const textEl = document.createElement("div");
  textEl.className = "chat-text";
  if (role === "assistant") {
    textEl.innerHTML = renderMarkdown(text);
  } else {
    textEl.textContent = text;
  }
  el.appendChild(roleEl);
  el.appendChild(textEl);
  if (status) {
    const statusEl = document.createElement("div");
    statusEl.className = "chat-status";
    statusEl.textContent = status;
    el.appendChild(statusEl);
  }
  history.appendChild(el);
  history.scrollTop = history.scrollHeight;
  return el;
}


/** Load model options into main dropdown on init */
async function loadModels() {
  const select = document.getElementById("model-select") as HTMLSelectElement | null;
  if (!select) return;
  const saved = localStorage.getItem("wr:model");
  await populateModelDropdown(select, saved || undefined);
  if (saved) {
    selectedModel = saved;
    select.value = saved;
  }
}

// ── Office Init ──
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";

    setupPromptUI();
    void loadModels();
    setupSettingsUI();
    setupTrackChangesToggle();
    setupRefStatus();
    setupQuickActions();
    connectWebSocket();
    // Initialize document session after WebSocket connects
    setTimeout(() => initSession(), 1000);
  }
});

// ── Session Persistence ──

/** Read SidebarSessionId from Word custom properties */
async function getDocSessionId(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const props = context.document.properties.customProperties;
      props.load("items");
      await context.sync();
      for (const prop of props.items) {
        if (prop.key === "SidebarSessionId") {
          prop.load("value");
          await context.sync();
          return prop.value as string;
        }
      }
      return null;
    });
  } catch (e) {
    console.error("[session] Failed to read doc property:", e);
    return null;
  }
}

/** Write SidebarSessionId to Word custom properties */
async function setDocSessionId(sessionId: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      context.document.properties.customProperties.add("SidebarSessionId", sessionId);
      await context.sync();
    });
  } catch (e) {
    console.error("[session] Failed to write doc property:", e);
  }
}

/** Remove all Sidebar-related custom properties from the document */
async function removeAITraces(): Promise<number> {
  let removed = 0;
  try {
    await Word.run(async (context) => {
      const props = context.document.properties.customProperties;
      props.load("items");
      await context.sync();
      for (const prop of props.items) {
        prop.load("key");
      }
      await context.sync();
      for (const prop of props.items) {
        if (prop.key.toLowerCase().includes("sidebar")) {
          prop.delete();
          removed++;
        }
      }
      await context.sync();
    });
  } catch (e) {
    console.error("[session] Failed to remove AI traces:", e);
  }
  return removed;
}

/** Initialize or resume a document session */
async function initSession(): Promise<void> {
  const existingId = await getDocSessionId();
  if (existingId) {
    // Try to resume
    try {
      const r = await fetch("http://localhost:3001/api/session/resume", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sessionId: existingId }),
      });
      const j = await r.json();
      if (j?.ok && j.data?.found) {
        currentDocSessionId = existingId;
        if (j.data.conversationHistory?.length) {
          repopulateChat(j.data.conversationHistory, j.data.recap);
        }
        console.log("[session] Resumed:", existingId);
        return;
      }
    } catch (e) {
      console.error("[session] Resume failed:", e);
    }
  }
  // Start fresh session
  await createNewSession();
}

/** Create a new session and write ID to document */
async function createNewSession(): Promise<string> {
  try {
    const r = await fetch("http://localhost:3001/api/session/new", { method: "POST" });
    const j = await r.json();
    if (j?.ok && j.data?.sessionId) {
      currentDocSessionId = j.data.sessionId;
      await setDocSessionId(j.data.sessionId);
      console.log("[session] Created new:", j.data.sessionId);
      return j.data.sessionId;
    }
  } catch (e) {
    console.error("[session] Failed to create:", e);
  }
  return "";
}

/** Repopulate chat UI from conversation history */
function repopulateChat(history: { role: string; content: string; timestamp?: number }[], recap?: string): void {
  const historyEl = document.getElementById("prompt-history");
  if (!historyEl) return;
  historyEl.innerHTML = "";

  // Show restoration banner
  if (history.length > 0) {
    const timestamps = history.filter(m => m.timestamp).map(m => m.timestamp!);
    const earliest = timestamps.length > 0 ? new Date(Math.min(...timestamps)) : null;
    const bannerText = earliest
      ? `📋 Session restored — ${history.length} messages from ${earliest.toLocaleDateString("en-US", { month: "short", day: "numeric" })}`
      : `📋 Session restored — ${history.length} messages`;
    const banner = document.createElement("div");
    banner.className = "session-restored-banner";
    banner.textContent = bannerText;
    historyEl.appendChild(banner);
  }

  // Show recap as system message
  if (recap) {
    const recapEl = document.createElement("div");
    recapEl.className = "chat-entry session-recap";
    const textEl = document.createElement("div");
    textEl.className = "chat-text";
    textEl.textContent = recap;
    recapEl.appendChild(textEl);
    historyEl.appendChild(recapEl);
  }

  for (const msg of history) {
    if (msg.role === "user" || msg.role === "assistant") {
      appendChatEntry(msg.role as "user" | "assistant", msg.content);
    }
  }
}

// ── Prompt UI ──
function setupPromptUI() {
  const input = document.getElementById("prompt-input") as HTMLInputElement;
  const btn = document.getElementById("prompt-send") as HTMLButtonElement;
  const modelSelect = document.getElementById("model-select") as HTMLSelectElement | null;
  if (modelSelect) {
    modelSelect.addEventListener("change", () => {
      if (modelSelect.value === "__configure__") {
        const toggle = document.getElementById("settings-toggle");
        if (toggle) toggle.click();
        modelSelect.value = "";
        return;
      }
      selectedModel = modelSelect.value;
      localStorage.setItem("wr:model", selectedModel);
    });
  }

  const sendPrompt = async () => {
    const text = input.value.trim();
    if (!text || !socket || socket.readyState !== WebSocket.OPEN) return;
    
    // Get current selection for context
    let context = "";
    try {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        if (sel.text) context = sel.text.substring(0, 500);
      });
    } catch {}
    
    const clientId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const model = modelSelect?.value || selectedModel || undefined;
    socket.send(JSON.stringify({ type: "prompt", clientId, text, model, context: context || undefined }));

    const el = appendChatEntry("user", text, "Sending…");
    el.setAttribute("data-client-id", clientId);
    pendingPromptEls.set(clientId, el);
    input.value = "";
  };

  btn.addEventListener("click", sendPrompt);
  input.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendPrompt(); } });

  // New Chat button
  const newChatBtn = document.getElementById("new-chat-btn");
  if (newChatBtn) {
    newChatBtn.addEventListener("click", async () => {
      // Delete old session on server
      try { await fetch("http://localhost:3001/api/session/delete", { method: "POST" }); } catch {}
      // Create new session and write to doc
      await createNewSession();
      // Clear chat UI
      const history = document.getElementById("prompt-history");
      if (history) history.innerHTML = "";
      pendingPromptEls.clear();
    });
  }
}

// ── WebSocket ──
function setStatus(text: string, ok: boolean) {
  const el = document.getElementById("status")!;
  el.textContent = text;
  el.className = ok ? "status-ok" : "status-err";
}

function resetHeartbeat() {
  if (heartbeatTimer) clearTimeout(heartbeatTimer);
  missedHeartbeats = 0;
  heartbeatTimer = setTimeout(onHeartbeatTimeout, HEARTBEAT_TIMEOUT);
}

function onHeartbeatTimeout() {
  missedHeartbeats++;
  if (missedHeartbeats >= 3) {
    setStatus("Reconnecting...", false);
    socket?.close();
  } else {
    setStatus(`Connected (${missedHeartbeats} missed)`, true);
    heartbeatTimer = setTimeout(onHeartbeatTimeout, HEARTBEAT_TIMEOUT);
  }
}

function connectWebSocket() {
  if (socket && (socket.readyState === WebSocket.OPEN || socket.readyState === WebSocket.CONNECTING)) return;
  try { socket = new WebSocket(WS_URL); } catch { scheduleReconnect(); return; }

  socket.onopen = () => {
    setStatus("Connected", true);
    reconnectDelay = RECONNECT_BASE;
    resetHeartbeat();
    if (reconnectTimer) { clearTimeout(reconnectTimer); reconnectTimer = null; }
  };

  socket.onclose = () => {
    setStatus("Disconnected", false);
    socket = null;
    if (heartbeatTimer) { clearTimeout(heartbeatTimer); heartbeatTimer = null; }
    scheduleReconnect();
  };

  socket.onerror = () => socket?.close();

  socket.onmessage = async (event) => {
    resetHeartbeat();
    try {
      const msg = JSON.parse(event.data);
      if (msg.type === "prompt_ack") {
        const clientId = msg.clientId as string | undefined;
        const promptId = msg.id as number | undefined;
        if (clientId && pendingPromptEls.has(clientId)) {
          const el = pendingPromptEls.get(clientId)!;
          const st = el.querySelector(".chat-status") as HTMLElement | null;
          if (st) st.textContent = "Sent";
          if (promptId !== undefined) el.setAttribute("data-prompt-id", String(promptId));
          pendingPromptEls.delete(clientId);
          // Show thinking indicator
          addThinkingIndicator();
        }
        return;
      }
      if (msg.type === "prompt_progress") {
        const promptId = msg.promptId as number | undefined;
        const progressText = msg.text as string | undefined;
        const progressLabel = msg.progressText as string | undefined;
        if (promptId === undefined) return;

        // If we have a descriptive progress label (tool activity), show it in thinking indicator
        if (progressLabel && !progressText) {
          const thinkingEl = document.getElementById("thinking-indicator");
          if (thinkingEl) {
            // Replace dots with progress text
            let labelEl = thinkingEl.querySelector(".progress-text") as HTMLElement;
            if (!labelEl) {
              thinkingEl.innerHTML = "";
              labelEl = document.createElement("span");
              labelEl.className = "progress-text";
              thinkingEl.appendChild(labelEl);
            }
            // Trigger re-animation by cloning
            const newLabel = labelEl.cloneNode(false) as HTMLElement;
            newLabel.textContent = progressLabel;
            labelEl.replaceWith(newLabel);
          }
        }

        if (!progressText) return;
        const history = document.getElementById("prompt-history")!;
        // Update or create streaming entry
        let streamEl = history.querySelector(`[data-streaming-for="${promptId}"]`) as HTMLElement | null;
        if (!streamEl) {
          streamEl = document.createElement("div");
          streamEl.className = "chat-entry chat-assistant";
          streamEl.setAttribute("data-streaming-for", String(promptId));
          const roleEl = document.createElement("div");
          roleEl.className = "chat-role";
          roleEl.textContent = "Assistant";
          const textEl = document.createElement("div");
          textEl.className = "chat-text";
          textEl.innerHTML = renderMarkdown(progressText);
          streamEl.appendChild(roleEl);
          streamEl.appendChild(textEl);
          const userEl = history.querySelector(`[data-prompt-id="${promptId}"]`) as HTMLElement | null;
          if (userEl) userEl.insertAdjacentElement("afterend", streamEl);
          else history.appendChild(streamEl);
        } else {
          const textEl = streamEl.querySelector(".chat-text") as HTMLElement;
          if (textEl) textEl.innerHTML = renderMarkdown(progressText);
        }
        history.scrollTop = history.scrollHeight;
        return;
      }
      if (msg.type === "prompt_response") {
        removeThinkingIndicator();
        const promptId = msg.promptId as number | undefined;
        const responseText = msg.text as string | undefined;
        if (!responseText) return;
        // Remove streaming preview if it exists
        if (promptId !== undefined) {
          const streamEl = document.getElementById("prompt-history")?.querySelector(`[data-streaming-for="${promptId}"]`);
          if (streamEl) streamEl.remove();
        }
        const history = document.getElementById("prompt-history")!;
        let inserted = false;
        if (promptId !== undefined) {
          const userEl = history.querySelector(`[data-prompt-id="${promptId}"]`) as HTMLElement | null;
          if (userEl) {
            const el = document.createElement("div");
            el.className = "chat-entry chat-assistant";
            el.setAttribute("data-response-for", String(promptId));
            const roleEl = document.createElement("div");
            roleEl.className = "chat-role";
            roleEl.textContent = "OpenClaw";
            const textEl = document.createElement("div");
            textEl.className = "chat-text";
            textEl.innerHTML = renderMarkdown(responseText);
            el.appendChild(roleEl);
            el.appendChild(textEl);
            // Add revert button if this exchange made document changes
            if (msg.hasChanges && msg.exchangeId) {
              const revertBtn = document.createElement("button");
              revertBtn.className = "revert-btn";
              revertBtn.textContent = "↩ Revert";
              revertBtn.setAttribute("data-exchange-id", String(msg.exchangeId));
              revertBtn.addEventListener("click", () => handleRevert(revertBtn, msg.exchangeId));
              el.appendChild(revertBtn);
            }
            // Add change summaries if present
            if (msg.changeSummaries && msg.changeSummaries.length > 0) {
              const summaryDiv = document.createElement("div");
              summaryDiv.className = "change-summary";
              const collapsed = msg.changeSummaries.length > 5;
              const header = document.createElement("div");
              header.className = "change-summary-header";
              header.textContent = `\ud83d\udcdd Changes (${msg.changeSummaries.length})`;
              header.addEventListener("click", () => {
                const lines = summaryDiv.querySelector(".change-summary-lines") as HTMLElement;
                if (lines) {
                  const isHidden = lines.style.display === "none";
                  lines.style.display = isHidden ? "block" : "none";
                  header.classList.toggle("expanded", isHidden);
                }
              });
              summaryDiv.appendChild(header);
              const linesDiv = document.createElement("div");
              linesDiv.className = "change-summary-lines";
              if (collapsed) linesDiv.style.display = "none";
              for (const line of msg.changeSummaries) {
                const lineEl = document.createElement("div");
                lineEl.className = "change-line";
                // Highlight old→new text with colors
                const arrowMatch = line.match(/^(.+?)"(.+?)"\s*\u2192\s*"(.+?)"(.*)$/);
                if (arrowMatch) {
                  lineEl.innerHTML = escapeHtml(arrowMatch[1]) +
                    '"<span class="old-text">' + escapeHtml(arrowMatch[2]) + '</span>"' +
                    ' → ' +
                    '"<span class="new-text">' + escapeHtml(arrowMatch[3]) + '</span>"' +
                    escapeHtml(arrowMatch[4]);
                } else {
                  lineEl.textContent = line;
                }
                linesDiv.appendChild(lineEl);
              }
              summaryDiv.appendChild(linesDiv);
              el.appendChild(summaryDiv);
            }
            userEl.insertAdjacentElement("afterend", el);
            inserted = true;
          }
        }
        if (!inserted) appendChatEntry("assistant", responseText);
        history.scrollTop = history.scrollHeight;
        return;
      }
      if (msg.id !== undefined && msg.command) {
        let result: any;
        let error: string | undefined;
        try {
          const timeout = msg.params?._timeout ?? DEFAULT_TIMEOUT;
          result = await withTimeout(handleCommand(msg.command, msg.params), timeout, msg.command);
        } catch (e: any) {
          error = e.message || String(e);
        }
        socket?.send(JSON.stringify({ id: msg.id, data: result, error }));
      }
    } catch (e) { console.error("[wr] Bad message:", e); }
  };
}

function scheduleReconnect() {
  if (reconnectTimer) return;
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connectWebSocket();
  }, reconnectDelay);
  reconnectDelay = Math.min(reconnectDelay * 1.5, RECONNECT_MAX);
}

// ══════════════════════════════════════
// COMMAND HANDLERS
// ══════════════════════════════════════

async function buildIndex(): Promise<DocumentIndex> {
  return Word.run(async (ctx) => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load("items");
    await ctx.sync();
    const batchSize = 200;
    const allMeta: ParagraphMeta[] = [];
    const headings: DocumentIndex["headings"] = [];
    for (let i = 0; i < paragraphs.items.length; i += batchSize) {
      const batch = paragraphs.items.slice(i, i + batchSize);
      const listItems: Word.ListItem[] = [];
      for (const p of batch) {
        p.load("text,style,isListItem,listItemLevel");
        listItems.push(p.listItemOrNullObject);
      }
      for (const li of listItems) li.load("listString");
      await ctx.sync();
      for (let j = 0; j < batch.length; j++) {
        const p = batch[j]; const li = listItems[j]; const idx = i + j;
        const text = p.text;
        allMeta.push({
          index: idx, preview: text.substring(0, 100), style: p.style,
          isListItem: p.isListItem, listItemLevel: p.isListItem ? p.listItemLevel : undefined,
          listString: li.isNullObject ? undefined : li.listString,
          charCount: text.length, hash: simpleHash(text),
        });
        const sLow = p.style.toLowerCase();
        if (sLow.startsWith("heading") || sLow.includes("_heading") || sLow.includes("centered_heading")) {
          const m = p.style.match(/\d+/);
          headings.push({ index: idx, text: text.substring(0, 120), level: m ? parseInt(m[0], 10) : 1, style: p.style });
        }
      }
    }
    docIndex = { version: (docIndex?.version ?? 0) + 1, paragraphCount: paragraphs.items.length, paragraphs: allMeta, headings, builtAt: Date.now() };
    return docIndex;
  });
}

async function handleCommand(command: string, params: any): Promise<any> {
  switch (command) {
    case "ping":
      return { pong: true, timestamp: Date.now() };

    case "buildIndex": return buildIndex();
    case "getIndex":
      if (!docIndex) await buildIndex();
      return docIndex;
    case "getHeadings":
      if (!docIndex) await buildIndex();
      return { headings: docIndex!.headings, paragraphCount: docIndex!.paragraphCount };
    case "getDelta": {
      const oldIdx = docIndex;
      const newIdx = await buildIndex();
      if (!oldIdx || (params?.sinceVersion ?? 0) < oldIdx.version) return { full: true, index: newIdx };
      const changed: number[] = [];
      const max = Math.max(oldIdx.paragraphs.length, newIdx.paragraphs.length);
      for (let i = 0; i < max; i++) {
        if (!oldIdx.paragraphs[i] || !newIdx.paragraphs[i] || oldIdx.paragraphs[i].hash !== newIdx.paragraphs[i].hash) changed.push(i);
      }
      return { full: false, version: newIdx.version, changed, paragraphCount: newIdx.paragraphCount };
    }
    case "getIndexRange": {
      if (!docIndex) await buildIndex();
      const from = params?.from ?? 0;
      const to = params?.to ?? docIndex!.paragraphs.length;
      return { paragraphs: docIndex!.paragraphs.slice(from, to), paragraphCount: docIndex!.paragraphCount, version: docIndex!.version };
    }

    case "getDocument":
      return Word.run(async (ctx) => { const b = ctx.document.body; b.load("text"); await ctx.sync(); return { text: b.text }; });

    case "getParagraphs":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const from = params?.from ?? 0;
        const to = params?.to ?? paragraphs.items.length;
        const slice = paragraphs.items.slice(from, to);
        const listItems: Word.ListItem[] = [];
        for (const p of slice) { p.load("text,style,isListItem,listItemLevel"); listItems.push(p.listItemOrNullObject); }
        for (const li of listItems) li.load("listString");
        await ctx.sync();
        const items = slice.map((p, i) => {
          const li = listItems[i];
          const base: any = { index: from + i, text: params?.compact ? p.text.substring(0, 100) : p.text, style: p.style, isListItem: p.isListItem };
          if (p.isListItem) base.listItemLevel = p.listItemLevel;
          if (!li.isNullObject) base.listString = li.listString;
          return base;
        });
        return { paragraphs: items, count: paragraphs.items.length };
      });

    case "getParagraph":
      return Word.run(async (ctx) => {
        const idx = params?.index;
        if (idx === undefined) throw new Error("params.index required");
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        if (idx < 0 || idx >= paragraphs.items.length) throw new Error(`index ${idx} out of range (0-${paragraphs.items.length - 1})`);
        const p = paragraphs.items[idx];
        p.load("text,style,isListItem,listItemLevel,font");
        const li = p.listItemOrNullObject;
        li.load("listString,siblingIndex,level");
        await ctx.sync();
        const result: any = { index: idx, text: p.text, style: p.style, isListItem: p.isListItem };
        if (p.isListItem) result.listItemLevel = p.listItemLevel;
        if (!li.isNullObject) result.listString = li.listString;
        if (!params?.compact) result.font = { name: p.font.name, size: p.font.size, bold: p.font.bold, italic: p.font.italic, color: p.font.color };
        return result;
      });

    case "getParagraphContext":
      return Word.run(async (ctx) => {
        const idx = params?.index;
        const radius = params?.radius ?? 2;
        if (idx === undefined) throw new Error("params.index required");
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const from = Math.max(0, idx - radius);
        const to = Math.min(paragraphs.items.length, idx + radius + 1);
        const slice = paragraphs.items.slice(from, to);
        const listItems: Word.ListItem[] = [];
        for (const p of slice) { p.load("text,style,isListItem"); listItems.push(p.listItemOrNullObject); }
        for (const li of listItems) li.load("listString");
        await ctx.sync();
        return {
          paragraphs: slice.map((p, i) => ({
            index: from + i, text: p.text, style: p.style,
            listString: listItems[i].isNullObject ? undefined : listItems[i].listString,
            isFocus: from + i === idx,
          })),
          focusIndex: idx,
        };
      });

    case "getDocumentStats":
      return Word.run(async (ctx) => {
        const body = ctx.document.body;
        body.load("text");
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        const sections = ctx.document.sections;
        sections.load("items");
        const footnotes = ctx.document.body.footnotes;
        footnotes.load("items");
        await ctx.sync();
        const text = body.text;
        const words = text.split(/\s+/).filter(w => w.length > 0).length;
        return { paragraphCount: paragraphs.items.length, wordCount: words, charCount: text.length, sectionCount: sections.items.length, footnoteCount: footnotes.items.length };
      });

    case "getDocumentStructure":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        // Load all at once in batches
        const tree: any[] = [];
        const batchSize = 200;
        for (let i = 0; i < paragraphs.items.length; i += batchSize) {
          const batch = paragraphs.items.slice(i, i + batchSize);
          for (const p of batch) p.load("text,style,isListItem");
          await ctx.sync();
          for (let j = 0; j < batch.length; j++) {
            const p = batch[j]; const sLow = p.style.toLowerCase();
            if (sLow.startsWith("heading") || sLow.includes("_heading") || sLow.includes("centered_heading") || sLow.includes("centered heading")) {
              const m = p.style.match(/\d+/);
              tree.push({ index: i + j, text: p.text.substring(0, 150), level: m ? parseInt(m[0], 10) : 1, style: p.style });
            }
          }
        }
        return { outline: tree, paragraphCount: paragraphs.items.length };
      });

    case "getToc":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const tocEntries: any[] = [];
        const batchSize = 200;
        for (let i = 0; i < paragraphs.items.length; i += batchSize) {
          const batch = paragraphs.items.slice(i, i + batchSize);
          for (const p of batch) p.load("text,style");
          await ctx.sync();
          for (let j = 0; j < batch.length; j++) {
            const p = batch[j];
            if (p.style.toLowerCase().startsWith("toc")) {
              tocEntries.push({ index: i + j, text: p.text, style: p.style });
            }
          }
        }
        return { entries: tocEntries, count: tocEntries.length };
      });

    case "getSelection":
      return Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text,style,font,isEmpty");
        const para = sel.paragraphs.getFirst();
        para.load("style,isListItem");
        const li = para.listItemOrNullObject;
        li.load("listString");
        await ctx.sync();
        return { text: sel.text, style: sel.style, isEmpty: sel.isEmpty, paragraphStyle: para.style, listString: li.isNullObject ? undefined : li.listString };
      });

    case "replaceSelection":
      return Word.run(async (ctx) => {
        const { text } = params || {};
        if (text === undefined) throw new Error("params.text required");
        const sel = ctx.document.getSelection();
        sel.load("text");
        const para = sel.paragraphs.getFirst();
        para.load("style");
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const original = sel.text;
        let paraIndex = -1;
        for (let i = 0; i < paragraphs.items.length; i++) { if (paragraphs.items[i] === para) { paraIndex = i; break; } }
        sel.insertText(text, Word.InsertLocation.replace);
        await ctx.sync();
        pushUndo({ command: "replaceSelection", original, replacement: text, paragraphIndex: paraIndex, style: para.style, timestamp: Date.now() });
        return { original, replacement: text, paragraphIndex: paraIndex, undoAvailable: true };
      });

    case "replaceParagraph":
      return Word.run(async (ctx) => {
        const { index, text, listString: targetLS } = params || {};
        let paraIndex = index;
        if (targetLS && paraIndex === undefined) {
          const paragraphs = ctx.document.body.paragraphs; paragraphs.load("items"); await ctx.sync();
          const lis: Word.ListItem[] = [];
          for (const p of paragraphs.items) lis.push(p.listItemOrNullObject);
          for (const l of lis) l.load("listString");
          await ctx.sync();
          for (let i = 0; i < paragraphs.items.length; i++) { if (!lis[i].isNullObject && lis[i].listString === targetLS) { paraIndex = i; break; } }
          if (paraIndex === undefined) throw new Error(`listString "${targetLS}" not found`);
        }
        if (paraIndex === undefined || text === undefined) throw new Error("index/listString and text required");
        const paragraphs = ctx.document.body.paragraphs; paragraphs.load("items"); await ctx.sync();
        if (paraIndex < 0 || paraIndex >= paragraphs.items.length) throw new Error(`index ${paraIndex} out of range`);
        const p = paragraphs.items[paraIndex]; p.load("text,style"); await ctx.sync();
        const original = p.text;
        const range = p.getRange(Word.RangeLocation.content);
        range.insertText(text, Word.InsertLocation.replace);
        await ctx.sync();
        pushUndo({ command: "replaceParagraph", original, replacement: text, paragraphIndex: paraIndex, style: p.style, timestamp: Date.now() });
        return { original, replacement: text, paragraphIndex: paraIndex, undoAvailable: true };
      });

    case "editSelection":
      return Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text,style,isEmpty");
        const para = sel.paragraphs.getFirst();
        para.load("style,isListItem");
        const li = para.listItemOrNullObject;
        li.load("listString");
        await ctx.sync();
        if (params?.replacement !== undefined) {
          const original = sel.text;
          sel.insertText(params.replacement, Word.InsertLocation.replace);
          await ctx.sync();
          pushUndo({ command: "editSelection", original, replacement: params.replacement, style: sel.style, timestamp: Date.now() });
        }
        return { text: sel.text, style: sel.style, isEmpty: sel.isEmpty, listString: li.isNullObject ? undefined : li.listString, replaced: params?.replacement !== undefined };
      });

    case "selectParagraph":
      return Word.run(async (ctx) => {
        const idx = params?.index;
        if (idx === undefined) throw new Error("params.index required");
        const paragraphs = ctx.document.body.paragraphs; paragraphs.load("items"); await ctx.sync();
        if (idx < 0 || idx >= paragraphs.items.length) throw new Error(`index ${idx} out of range`);
        const p = paragraphs.items[idx];
        p.load("text");
        p.select();
        await ctx.sync();
        return { index: idx, text: p.text };
      });

    case "navigateToParagraph":
      return Word.run(async (ctx) => {
        const idx = params?.index;
        if (idx === undefined) throw new Error("params.index required");
        const paragraphs = ctx.document.body.paragraphs; paragraphs.load("items"); await ctx.sync();
        if (idx < 0 || idx >= paragraphs.items.length) throw new Error(`index ${idx} out of range`);
        paragraphs.items[idx].select();
        await ctx.sync();
        return { index: idx, scrolled: true };
      });

    case "getStyles":
      return Word.run(async (ctx) => {
        const styles = ctx.document.getStyles(); styles.load("items"); await ctx.sync();
        styles.load("nameLocal,type,builtIn"); await ctx.sync();
        return { styles: styles.items.map(s => ({ name: s.nameLocal, type: s.type, builtIn: s.builtIn })) };
      });

    case "setStyleFont":
      return Word.run(async (ctx) => {
        const { styleName, fontName, fontSize, bold, italic, color } = params || {};
        if (!styleName) throw new Error("params.styleName required");
        const style = ctx.document.getStyles().getByNameOrNullObject(styleName);
        style.load("nameLocal"); await ctx.sync();
        if (style.isNullObject) throw new Error(`Style "${styleName}" not found`);
        if (fontName) style.font.name = fontName;
        if (fontSize) style.font.size = fontSize;
        if (bold !== undefined) style.font.bold = bold;
        if (italic !== undefined) style.font.italic = italic;
        if (color) style.font.color = color;
        await ctx.sync();
        return { styleName, fontName, fontSize };
      });

    case "find":
      return Word.run(async (ctx) => {
        const searchText: string = params?.text;
        if (!searchText) throw new Error("params.text required");
        const results = ctx.document.body.search(searchText, { matchCase: params?.matchCase ?? false, matchWholeWord: params?.matchWholeWord ?? false });
        results.load("items"); await ctx.sync();
        results.load("text,style"); await ctx.sync();
        return { matches: results.items.map((r, i) => ({ index: i, text: r.text, style: r.style })), count: results.items.length };
      });

    case "findReplace":
      return Word.run(async (ctx) => {
        const { text, replacement, matchCase, matchWholeWord } = params || {};
        if (!text || replacement === undefined) throw new Error("text and replacement required");
        const results = ctx.document.body.search(text, { matchCase: matchCase ?? false, matchWholeWord: matchWholeWord ?? false });
        results.load("items"); await ctx.sync();
        const count = results.items.length;
        for (const item of results.items) item.insertText(replacement, Word.InsertLocation.replace);
        await ctx.sync();
        return { replacedCount: count };
      });

    case "insert":
      return Word.run(async (ctx) => {
        const { text, location, paragraphIndex, style } = params || {};
        if (!text) throw new Error("params.text required");
        let paragraph: Word.Paragraph;
        if (paragraphIndex !== undefined) {
          const paragraphs = ctx.document.body.paragraphs; paragraphs.load("items"); await ctx.sync();
          if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) throw new Error(`paragraphIndex out of range`);
          paragraph = paragraphs.items[paragraphIndex].insertParagraph(text, location === "before" ? Word.InsertLocation.before : Word.InsertLocation.after);
        } else if (location === "start") {
          paragraph = ctx.document.body.insertParagraph(text, Word.InsertLocation.start);
        } else {
          paragraph = ctx.document.body.insertParagraph(text, Word.InsertLocation.end);
        }
        if (style) paragraph.style = style;
        paragraph.load("text,style"); await ctx.sync();
        return { text: paragraph.text, style: paragraph.style };
      });

    case "format":
      return Word.run(async (ctx) => {
        const { text, bold, italic, underline, color, highlightColor, style } = params || {};
        if (!text) throw new Error("params.text required");
        const results = ctx.document.body.search(text, { matchCase: true });
        results.load("items"); await ctx.sync();
        if (results.items.length === 0) throw new Error("Text not found");
        for (const item of results.items) {
          if (bold !== undefined) item.font.bold = bold;
          if (italic !== undefined) item.font.italic = italic;
          if (underline !== undefined) item.font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
          if (color) item.font.color = color;
          if (highlightColor) item.font.highlightColor = highlightColor;
          if (style) item.style = style;
        }
        await ctx.sync();
        return { formattedCount: results.items.length };
      });

    case "updateParagraph":
      return Word.run(async (ctx) => {
        const { index, text, style } = params || {};
        if (index === undefined) throw new Error("params.index required");
        const paragraphs = ctx.document.body.paragraphs; paragraphs.load("items"); await ctx.sync();
        if (index < 0 || index >= paragraphs.items.length) throw new Error(`index out of range`);
        const p = paragraphs.items[index];
        if (text !== undefined) p.insertText(text, Word.InsertLocation.replace);
        if (style) p.style = style;
        p.load("text,style"); await ctx.sync();
        return { text: p.text, style: p.style };
      });

    case "getFootnotes":
      return Word.run(async (ctx) => {
        const footnotes = ctx.document.body.footnotes; footnotes.load("items"); await ctx.sync();
        for (const fn of footnotes.items) fn.body.load("text");
        await ctx.sync();
        return { footnotes: footnotes.items.map((fn, i) => ({ index: i, body: fn.body.text })), count: footnotes.items.length };
      });

    case "addFootnote":
      return Word.run(async (ctx) => {
        const { searchText, footnoteText } = params || {};
        if (!searchText || !footnoteText) throw new Error("searchText and footnoteText required");
        const results = ctx.document.body.search(searchText, { matchCase: true });
        results.load("items"); await ctx.sync();
        if (results.items.length === 0) throw new Error("searchText not found");
        const range = results.items[0].getRange(Word.RangeLocation.end);
        const fn = range.insertFootnote(footnoteText);
        fn.body.load("text"); await ctx.sync();
        return { body: fn.body.text };
      });

    case "updateFootnote":
      return Word.run(async (ctx) => {
        const { index, text } = params || {};
        if (index === undefined || !text) throw new Error("index and text required");
        const footnotes = ctx.document.body.footnotes; footnotes.load("items"); await ctx.sync();
        if (index < 0 || index >= footnotes.items.length) throw new Error(`index out of range`);
        footnotes.items[index].body.insertText(text, Word.InsertLocation.replace);
        footnotes.items[index].body.load("text"); await ctx.sync();
        return { body: footnotes.items[index].body.text };
      });

    case "searchFootnotes":
      return Word.run(async (ctx) => {
        const searchText: string = params?.text;
        if (!searchText) throw new Error("params.text required");
        const footnotes = ctx.document.body.footnotes; footnotes.load("items"); await ctx.sync();
        for (const fn of footnotes.items) fn.body.load("text");
        await ctx.sync();
        const matches = footnotes.items
          .map((fn, i) => ({ index: i, body: fn.body.text }))
          .filter(fn => fn.body.toLowerCase().includes(searchText.toLowerCase()));
        return { matches, count: matches.length };
      });

    case "getComments":
      return Word.run(async (ctx) => {
        // Word JS API comment support — requires WordApi 1.4+
        try {
          const comments = ctx.document.body.getComments();
          comments.load("items"); await ctx.sync();
          for (const c of comments.items) { c.load("content,authorName,createdDate"); }
          await ctx.sync();
          return { comments: comments.items.map((c, i) => ({ index: i, content: c.content, author: c.authorName, created: c.createdDate })), count: comments.items.length };
        } catch {
          return { comments: [], count: 0, note: "Comments API not available in this Word version" };
        }
      });

    case "addComment":
      return Word.run(async (ctx) => {
        const { searchText, commentText } = params || {};
        if (!searchText || !commentText) throw new Error("searchText and commentText required");
        const results = ctx.document.body.search(searchText, { matchCase: true });
        results.load("items"); await ctx.sync();
        if (results.items.length === 0) throw new Error("searchText not found");
        try {
          const comment = results.items[0].insertComment(commentText);
          comment.load("content"); await ctx.sync();
          return { content: comment.content };
        } catch {
          return { error: "Comments API not available in this Word version" };
        }
      });

    case "undo":
      if (undoStack.length === 0) return { error: "Nothing to undo" };
      const entry = undoStack.pop()!;
      return Word.run(async (ctx) => {
        if (entry.paragraphIndex !== undefined && entry.paragraphIndex >= 0) {
          const paragraphs = ctx.document.body.paragraphs; paragraphs.load("items"); await ctx.sync();
          if (entry.paragraphIndex < paragraphs.items.length) {
            const range = paragraphs.items[entry.paragraphIndex].getRange(Word.RangeLocation.content);
            range.insertText(entry.original, Word.InsertLocation.replace);
            await ctx.sync();
            return { reverted: true, paragraphIndex: entry.paragraphIndex, text: entry.original };
          }
        }
        return { reverted: false, reason: "Could not find original paragraph" };
      });

    case "undoHistory":
      return {
        entries: undoStack.map((e, i) => ({ index: i, command: e.command, paragraphIndex: e.paragraphIndex, originalPreview: e.original.substring(0, 80), replacementPreview: e.replacement.substring(0, 80), timestamp: e.timestamp })),
        count: undoStack.length,
      };

    case "trackChanges":
      return { note: "Track changes toggle not available via JS API on Mac. Use Word UI." };

    case "getDocumentHtml":
      return Word.run(async (ctx) => { const html = ctx.document.body.getHtml(); await ctx.sync(); return { html: html.value }; });

    case "getPageCount":
      return Word.run(async (ctx) => {
        const sections = ctx.document.sections; sections.load("items"); await ctx.sync();
        return { sectionCount: sections.items.length, note: "Word JS API does not expose page numbers. Use PDF export." };
      });

    case "batch":
      return Word.run(async (ctx) => {
        const ops: any[] = params?.operations;
        if (!Array.isArray(ops)) throw new Error("operations must be an array");
        const results: any[] = [];
        for (const op of ops) results.push(await handleCommand(op.command, op.params));
        return { results };
      });


    case "getSection":
      return Word.run(async (ctx) => {
        const { heading, headingIndex } = params || {};
        if (!heading && headingIndex === undefined) throw new Error("heading or headingIndex required");
        // Build index if needed
        if (!docIndex) await buildIndex();
        // Find the heading
        let startIdx = -1;
        let endIdx = docIndex!.paragraphCount;
        let headingLevel = 0;
        for (let i = 0; i < docIndex!.headings.length; i++) {
          const h = docIndex!.headings[i];
          if (heading && h.text.includes(heading)) {
            startIdx = h.index;
            headingLevel = h.level;
            // Find next heading at same or higher level
            for (let j = i + 1; j < docIndex!.headings.length; j++) {
              if (docIndex!.headings[j].level <= headingLevel) {
                endIdx = docIndex!.headings[j].index;
                break;
              }
            }
            break;
          }
          if (headingIndex !== undefined && headingIndex === i) {
            startIdx = h.index;
            headingLevel = h.level;
            for (let j = i + 1; j < docIndex!.headings.length; j++) {
              if (docIndex!.headings[j].level <= headingLevel) {
                endIdx = docIndex!.headings[j].index;
                break;
              }
            }
            break;
          }
        }
        if (startIdx === -1) throw new Error("Heading not found");
        // Now read those paragraphs
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const slice = paragraphs.items.slice(startIdx, endIdx);
        const listItems: Word.ListItem[] = [];
        for (const p of slice) { p.load("text,style,isListItem"); listItems.push(p.listItemOrNullObject); }
        for (const li of listItems) li.load("listString");
        await ctx.sync();
        return {
          heading: docIndex!.headings.find(h => h.index === startIdx),
          paragraphs: slice.map((p, i) => ({
            index: startIdx + i, text: p.text, style: p.style,
            listString: listItems[i].isNullObject ? undefined : listItems[i].listString,
          })),
          range: { from: startIdx, to: endIdx },
        };
      });

    case "getBulkParagraphs":
      return Word.run(async (ctx) => {
        const indices: number[] = params?.indices;
        if (!Array.isArray(indices)) throw new Error("params.indices must be an array");
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const results: any[] = [];
        // Batch load
        const items = indices.map(idx => {
          if (idx < 0 || idx >= paragraphs.items.length) return null;
          return paragraphs.items[idx];
        });
        const listItems: (Word.ListItem | null)[] = [];
        for (const p of items) {
          if (p) { p.load("text,style,isListItem"); listItems.push(p.listItemOrNullObject); }
          else listItems.push(null);
        }
        for (const li of listItems) { if (li) li.load("listString"); }
        await ctx.sync();
        for (let i = 0; i < indices.length; i++) {
          const p = items[i];
          const li = listItems[i];
          if (!p) { results.push({ index: indices[i], error: "out of range" }); continue; }
          results.push({
            index: indices[i], text: p.text, style: p.style,
            listString: li && !li.isNullObject ? li.listString : undefined,
          });
        }
        return { paragraphs: results };
      });

    case "getDocumentProperties":
      return Word.run(async (ctx) => {
        const props = ctx.document.properties;
        props.load("title,subject,author,keywords,comments,category,lastAuthor,revisionNumber,creationDate,lastSaveTime");
        await ctx.sync();
        return {
          title: props.title, subject: props.subject, author: props.author,
          keywords: props.keywords, comments: props.comments, category: props.category,
          lastAuthor: props.lastAuthor, revisionNumber: props.revisionNumber,
          created: props.creationDate, lastSaved: props.lastSaveTime,
        };
      });

    case "diffParagraph":
      return Word.run(async (ctx) => {
        const { index, compareText } = params || {};
        if (index === undefined || !compareText) throw new Error("index and compareText required");
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        if (index < 0 || index >= paragraphs.items.length) throw new Error(`index ${index} out of range`);
        const p = paragraphs.items[index];
        p.load("text");
        await ctx.sync();
        const current = p.text;
        // Simple word-level diff
        const currentWords = current.split(/\s+/);
        const compareWords = compareText.split(/\s+/);
        const added: string[] = [];
        const removed: string[] = [];
        const cSet = new Set(currentWords);
        const nSet = new Set(compareWords);
        for (const w of currentWords) { if (!nSet.has(w)) removed.push(w); }
        for (const w of compareWords) { if (!cSet.has(w)) added.push(w); }
        return {
          index, current, compareText,
          same: current === compareText,
          added, removed,
          currentWordCount: currentWords.length,
          compareWordCount: compareWords.length,
        };
      });
    default:
      throw new Error(`Unknown command: ${command}`);
  }
}

// ── Revert Handler ──
async function handleRevert(btn: HTMLButtonElement, exchangeId: number): Promise<void> {
  if (btn.classList.contains("reverted")) return;
  btn.textContent = "Reverting...";
  btn.disabled = true;
  try {
    const r = await fetch(`http://localhost:3001/api/revert/${exchangeId}`, { method: "POST" });
    const j = await r.json();
    if (j.ok) {
      btn.textContent = "↩ Reverted";
      btn.classList.add("reverted");
      // Grey out all subsequent exchanges that were also reverted
      const revertedIds: number[] = j.data?.revertedExchanges || [exchangeId];
      const history = document.getElementById("prompt-history");
      if (history) {
        const allEntries = history.querySelectorAll(".chat-entry.chat-assistant");
        let foundCurrent = false;
        allEntries.forEach((entry) => {
          const entryBtn = entry.querySelector(".revert-btn") as HTMLButtonElement | null;
          const eid = entryBtn?.getAttribute("data-exchange-id");
          if (eid && revertedIds.includes(parseInt(eid, 10))) {
            entry.classList.add("reverted");
            if (entryBtn && entryBtn !== btn) {
              entryBtn.textContent = "↩ Reverted";
              entryBtn.classList.add("reverted");
              entryBtn.disabled = true;
            }
          }
        });
      }
    } else {
      btn.textContent = "↩ Failed";
      btn.disabled = false;
    }
  } catch (e) {
    console.error("Revert failed:", e);
    btn.textContent = "↩ Failed";
    btn.disabled = false;
  }
}

// ── Settings Panel ──

/** Available models for the settings dropdown */
interface SettingsModel {
  id: string;
  label: string;
  backend: string;
}

/** Load settings from the server and populate the form */
async function loadSettings(): Promise<void> {
  try {
    const r = await fetch("http://localhost:3001/api/settings");
    const j = await r.json();
    if (!j?.ok) return;
    const data = j.data || {};

    const openclawUrl = document.getElementById("settings-openclaw-url") as HTMLInputElement;
    const openaiKey = document.getElementById("settings-openai-key") as HTMLInputElement;
    const anthropicKey = document.getElementById("settings-anthropic-key") as HTMLInputElement;
    const defaultModel = document.getElementById("settings-default-model") as HTMLSelectElement;

    if (openclawUrl) openclawUrl.value = data.openclawUrl || "";
    const openclawToken = document.getElementById("settings-openclaw-token") as HTMLInputElement;
    if (openclawToken) openclawToken.value = data.openclawToken || "";
    if (openaiKey) openaiKey.value = data.openaiApiKey || "";
    if (anthropicKey) anthropicKey.value = data.anthropicApiKey || "";

    // Populate reference folders
    const refContainer = document.getElementById("settings-reference-folders");
    if (refContainer) {
      refContainer.innerHTML = "";
      const folders = data.referenceFolders || [];
      for (const f of folders) {
        addReferenceFolderRow(f);
      }
    }

    // Populate local endpoints
    const container = document.getElementById("settings-local-endpoints");
    if (container) {
      container.innerHTML = "";
      const endpoints = data.localEndpoints || [];
      for (const ep of endpoints) {
        addLocalEndpointRow(ep.name, ep.baseUrl);
      }
    }

    // Populate model dropdown with known models
    if (defaultModel) {
      await populateModelDropdown(defaultModel, data.defaultModel);
    }
  } catch (e) {
    console.error("Failed to load settings:", e);
  }
}

/** Add a local endpoint row to the settings form */
function addLocalEndpointRow(name: string = "", baseUrl: string = ""): void {
  const container = document.getElementById("settings-local-endpoints");
  if (!container) return;

  const row = document.createElement("div");
  row.className = "local-endpoint";
  row.innerHTML = `
    <input type="text" placeholder="Name" value="${escapeHtml(name)}" class="ep-name" />
    <input type="text" placeholder="http://host:port/v1" value="${escapeHtml(baseUrl)}" class="ep-url" />
    <button class="remove-endpoint" title="Remove">✕</button>
  `;

  const removeBtn = row.querySelector(".remove-endpoint") as HTMLButtonElement;
  removeBtn.addEventListener("click", () => row.remove());

  container.appendChild(row);
}

/** Add a reference folder row to the settings form */
function addReferenceFolderRow(folderPath: string = ""): void {
  const container = document.getElementById("settings-reference-folders");
  if (!container) return;
  const row = document.createElement("div");
  row.className = "ref-folder-row";
  row.innerHTML = `
    <input type="text" placeholder="/path/to/case/folder" value="${escapeHtml(folderPath)}" class="ref-path" />
    <button class="remove-endpoint" title="Remove">✕</button>
  `;
  const removeBtn = row.querySelector(".remove-endpoint") as HTMLButtonElement;
  removeBtn.addEventListener("click", () => row.remove());
  container.appendChild(row);
}

/** Escape HTML special characters */
function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Populate the model dropdown with available models from all backends */
async function populateModelDropdown(select: HTMLSelectElement, currentDefault?: string): Promise<void> {
  const models: SettingsModel[] = [];

  // Fetch settings to determine which backends are configured
  let settings: any = {};
  try {
    const r = await fetch("http://localhost:3001/api/settings");
    const j = await r.json();
    if (j?.ok) settings = j.data || {};
  } catch (e) {
    console.error("Failed to fetch settings for model list:", e);
  }

  // OpenClaw
  if (settings.openclawUrl) {
    models.push({ id: "openclaw", label: "OpenClaw (default)", backend: "openclaw" });
  }

  // OpenAI — only if key is configured
  if (settings.openaiApiKey) {
    models.push(
      { id: "openai:gpt-4o", label: "GPT-4o (OpenAI)", backend: "openai" },
      { id: "openai:gpt-4o-mini", label: "GPT-4o Mini (OpenAI)", backend: "openai" },
      { id: "openai:o1", label: "o1 (OpenAI)", backend: "openai" },
      { id: "openai:o3-mini", label: "o3-mini (OpenAI)", backend: "openai" },
    );
  }

  // Anthropic — only if key is configured
  if (settings.anthropicApiKey) {
    models.push(
      { id: "anthropic:claude-sonnet-4-20250514", label: "Claude Sonnet 4 (Anthropic)", backend: "anthropic" },
      { id: "anthropic:claude-opus-4-20250514", label: "Claude Opus 4 (Anthropic)", backend: "anthropic" },
      { id: "anthropic:claude-3-haiku-20240307", label: "Claude 3 Haiku (Anthropic)", backend: "anthropic" },
    );
  }

  // Local endpoints
  const endpoints = settings.localEndpoints || [];
  for (const ep of endpoints) {
    if (ep.name && ep.baseUrl) {
      models.push({
        id: `local:${ep.baseUrl}:default`,
        label: `${ep.name} (Local)`,
        backend: "local",
      });
    }
  }

  // Build the select options
  select.innerHTML = "";

  if (models.length === 0) {
    // No backends configured — prompt user to configure
    const opt = document.createElement("option");
    opt.value = "__configure__";
    opt.textContent = "⚙️ Configure API keys in settings...";
    select.appendChild(opt);

    return;
  }

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select model...";
  select.appendChild(placeholder);

  for (const m of models) {
    const opt = document.createElement("option");
    opt.value = m.id;
    opt.textContent = m.label;
    if (currentDefault && m.id === currentDefault) opt.selected = true;
    select.appendChild(opt);
  }
}

/** Save settings to the server */
async function saveSettings(): Promise<boolean> {
  const openclawUrl = (document.getElementById("settings-openclaw-url") as HTMLInputElement)?.value;
  const openclawToken = (document.getElementById("settings-openclaw-token") as HTMLInputElement)?.value;
  const openaiKey = (document.getElementById("settings-openai-key") as HTMLInputElement)?.value;
  const anthropicKey = (document.getElementById("settings-anthropic-key") as HTMLInputElement)?.value;
  const defaultModel = (document.getElementById("settings-default-model") as HTMLSelectElement)?.value;

  // Collect local endpoints
  const localEndpoints: { name: string; baseUrl: string }[] = [];
  const container = document.getElementById("settings-local-endpoints");
  if (container) {
    const rows = container.querySelectorAll(".local-endpoint");
    rows.forEach((row) => {
      const name = (row.querySelector(".ep-name") as HTMLInputElement)?.value?.trim();
      const baseUrl = (row.querySelector(".ep-url") as HTMLInputElement)?.value?.trim();
      if (name && baseUrl) localEndpoints.push({ name, baseUrl });
    });
  }

  // Collect reference folders
  const referenceFolders: string[] = [];
  const refContainer = document.getElementById("settings-reference-folders");
  if (refContainer) {
    const rows = refContainer.querySelectorAll(".ref-folder-row");
    rows.forEach((row) => {
      const p = (row.querySelector(".ref-path") as HTMLInputElement)?.value?.trim();
      if (p) referenceFolders.push(p);
    });
  }

  const body: Record<string, any> = {};
  if (referenceFolders.length > 0) body.referenceFolders = referenceFolders;
  else body.referenceFolders = [];
  if (openclawUrl) body.openclawUrl = openclawUrl;
  if (openclawToken) body.openclawToken = openclawToken;
  if (openaiKey) body.openaiApiKey = openaiKey;
  if (anthropicKey) body.anthropicApiKey = anthropicKey;
  if (localEndpoints.length) body.localEndpoints = localEndpoints;
  if (defaultModel) body.defaultModel = defaultModel;

  try {
    const r = await fetch("http://localhost:3001/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const j = await r.json();
    return j?.ok === true;
  } catch (e) {
    console.error("Failed to save settings:", e);
    return false;
  }
}

/** Initialize the settings panel UI */
function setupSettingsUI(): void {
  const toggle = document.getElementById("settings-toggle");
  const panel = document.getElementById("settings-panel");
  const saveBtn = document.getElementById("settings-save");
  const cancelBtn = document.getElementById("settings-cancel");
  const addEndpointBtn = document.getElementById("settings-add-endpoint");
  const statusEl = document.getElementById("settings-status");

  if (!toggle || !panel) return;

  const overlay = document.getElementById("settings-overlay");

  function openSettings() {
    overlay?.classList.add("visible");
    panel.classList.add("visible");
    loadSettings();
  }
  function closeSettings() {
    overlay?.classList.remove("visible");
    panel.classList.remove("visible");
  }

  toggle.addEventListener("click", () => {
    const isVisible = panel.classList.contains("visible");
    if (isVisible) closeSettings();
    else openSettings();
  });

  cancelBtn?.addEventListener("click", closeSettings);

  overlay?.addEventListener("click", (e) => {
    if (e.target === overlay) closeSettings();
  });

  addEndpointBtn?.addEventListener("click", () => {
    addLocalEndpointRow();
  });

  const addFolderBtn = document.getElementById("settings-add-folder");
  addFolderBtn?.addEventListener("click", () => {
    addReferenceFolderRow();
  });

  // Test OpenClaw connection button
  const testBtn = document.getElementById("settings-test-openclaw");
  testBtn?.addEventListener("click", async () => {
    const urlInput = document.getElementById("settings-openclaw-url") as HTMLInputElement;
    const url = urlInput?.value?.trim();
    if (!url) {
      if (statusEl) { statusEl.style.display = "block"; statusEl.textContent = "Enter an OpenClaw URL first"; statusEl.classList.add("error"); }
      return;
    }
    if (statusEl) { statusEl.style.display = "block"; statusEl.textContent = "Testing connection..."; statusEl.classList.remove("error"); }
    try {
      const r = await fetch(`${url.replace(/\/$/, "")}/v1/chat/completions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "openclaw:main", messages: [{ role: "user", content: "ping" }], max_tokens: 1 }),
      });
      if (r.ok) {
        if (statusEl) { statusEl.textContent = "✓ OpenClaw reachable"; statusEl.classList.remove("error"); }
      } else {
        const text = await r.text();
        if (statusEl) { statusEl.textContent = `✕ HTTP ${r.status}: ${text.slice(0, 100)}`; statusEl.classList.add("error"); }
      }
    } catch (e: any) {
      if (statusEl) { statusEl.textContent = `✕ ${e.message}`; statusEl.classList.add("error"); }
    }
  });

  saveBtn?.addEventListener("click", async () => {
    if (statusEl) {
      statusEl.style.display = "block";
      statusEl.textContent = "Saving...";
      statusEl.classList.remove("error");
    }

    const ok = await saveSettings();

    if (statusEl) {
      if (ok) {
        statusEl.textContent = "✓ Settings saved";
        statusEl.classList.remove("error");
        // Also refresh the model dropdown in the prompt UI
        await loadModels();
        setTimeout(() => {
          statusEl.style.display = "none";
          overlay?.classList.remove("visible"); panel.classList.remove("visible");
        }, 1500);
      } else {
        statusEl.textContent = "✕ Failed to save settings";
        statusEl.classList.add("error");
      }
    }
  });
}

// ── Reference Status ──

async function updateRefStatus(): Promise<void> {
  const el = document.getElementById("ref-status");
  if (!el) return;
  try {
    const r = await fetch("http://localhost:3001/api/references/status");
    const j = await r.json();
    if (j?.ok && j.data) {
      const { documentCount, totalChunks, scanning: isScanning } = j.data;
      if (documentCount > 0) {
        el.style.display = "inline";
        el.childNodes[0].textContent = `📁 ${documentCount} doc${documentCount !== 1 ? "s" : ""}${isScanning ? " ⟳" : ""} `;
      } else {
        el.style.display = "none";
      }
    }
  } catch {}
}

async function showRefPopup(): Promise<void> {
  const popup = document.getElementById("ref-status-popup");
  if (!popup) return;
  try {
    const r = await fetch("http://localhost:3001/api/references");
    const j = await r.json();
    if (j?.ok && j.data?.documents) {
      popup.innerHTML = j.data.documents.map((d: any) =>
        `<div class="ref-file">📄 ${escapeHtml(d.filename)} (${d.chunkCount} chunks)</div>`
      ).join("");
      if (j.data.documents.length === 0) popup.innerHTML = "<div class=\"ref-file\">No documents indexed</div>";
    }
  } catch {
    popup.innerHTML = "<div class=\"ref-file\">Unable to fetch</div>";
  }
}

function setupRefStatus(): void {
  const el = document.getElementById("ref-status");
  if (!el) return;
  el.addEventListener("click", (e) => {
    e.stopPropagation();
    const popup = document.getElementById("ref-status-popup");
    if (!popup) return;
    const isVisible = popup.classList.contains("visible");
    popup.classList.toggle("visible", !isVisible);
    if (!isVisible) showRefPopup();
  });
  document.addEventListener("click", () => {
    document.getElementById("ref-status-popup")?.classList.remove("visible");
  });
  // Poll every 30s
  updateRefStatus();
  setInterval(updateRefStatus, 30000);
}

// ── Track Changes / YOLO Mode ──
let trackChangesMode = false;

function setupTrackChangesToggle(): void {
  const modeToggle = document.getElementById("mode-toggle");
  if (!modeToggle) return;

  modeToggle.addEventListener("click", async () => {
    trackChangesMode = !trackChangesMode;
    modeToggle.textContent = trackChangesMode ? "🔍 Track" : "⚡ YOLO";
    modeToggle.classList.toggle("tracking", trackChangesMode);

    // Tell Word to enable/disable track changes
    try {
      await Word.run(async (context) => {
        if (trackChangesMode) {
          context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        } else {
          context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        }
        await context.sync();
      });
    } catch (e) {
      console.warn("changeTrackingMode not available:", e);
    }

    // Tell the server
    fetch("http://localhost:3001/api/settings/mode", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ trackChanges: trackChangesMode }),
    }).catch(() => {});

    // Persist preference
    localStorage.setItem("sidebar-track-changes", String(trackChangesMode));
  });

  
  // Tighten button — tightens selected text
  const tightenBtn = document.getElementById("tighten-btn");
  tightenBtn?.addEventListener("click", async () => {
    const prompt = `Tighten the following selected text. Make it more concise without losing any substance, meaning, or persuasive force. Reduce word count significantly. Do not change the tone or weaken any arguments. Replace the selection with the tightened version.

Rules:
- Cut filler words, redundancies, and throat-clearing
- Combine sentences where possible
- Prefer active voice
- Keep every substantive point
- Maintain the same level of formality
- Do not add new content`;
    
    const input = document.getElementById("prompt-input") as HTMLTextAreaElement;
    if (input) {
      input.value = prompt;
      const sendBtn = document.getElementById("prompt-send") as HTMLButtonElement;
      sendBtn?.click();
    }
  });

  // Restore on load
  const savedMode = localStorage.getItem("sidebar-track-changes");
  if (savedMode === "true") {
    trackChangesMode = true;
    modeToggle.textContent = "🔍 Track";
    modeToggle.classList.add("tracking");
    Word.run(async (context) => {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await context.sync();
    }).catch(() => {});
    // Also tell server on load
    fetch("http://localhost:3001/api/settings/mode", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ trackChanges: true }),
    }).catch(() => {});
  }

  // Restore original state when leaving
  window.addEventListener("beforeunload", () => {
    if (trackChangesMode) {
      Word.run(async (context) => {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        await context.sync();
      }).catch(() => {});
    }
  });
}

// ══════════════════════════════════════
// QUICK ACTIONS
// ══════════════════════════════════════

interface QuickAction {
  name: string;
  prompt: string;
}

interface QuickActionCategory {
  label: string;
  actions: QuickAction[];
}

const QUICK_ACTIONS_DEFAULTS: QuickActionCategory[] = [
  {
    label: "📋 Cite",
    actions: [
      { name: "Check Bluebook", prompt: "Review all citations in this document for Bluebook formatting errors. List each error with the paragraph number, the incorrect citation, and the corrected version." },
      { name: "Long/Short Cites", prompt: "Check that every short citation in this document has a corresponding full citation earlier in the document. Flag any short cites that don't match a prior long cite, and any long cites that are repeated unnecessarily." },
      { name: "TOA Pages", prompt: "Review the Table of Authorities and verify that the page numbers are correct. List any discrepancies." },
      { name: "Find Uncited", prompt: "Identify any factual assertions in this document that lack a supporting citation or footnote." },
    ]
  },
  {
    label: "📝 Format",
    actions: [
      { name: "Defined Terms", prompt: "Check all defined terms in this document for consistency. Verify that each term is defined before first use, that definitions match usage, and that capitalization is consistent." },
      { name: "House Style", prompt: "Review this document for common style issues: Oxford comma consistency, 'shall' vs 'will' vs 'must', passive voice, legalese that could be simplified, and inconsistent formatting." },
      { name: "Cross-References", prompt: "Check all internal cross-references (e.g., 'See Section 3.2', 'as defined in Article IV') and verify they point to the correct sections." },
      { name: "Clean Up", prompt: "Clean up formatting issues: remove double spaces, fix inconsistent paragraph spacing, normalize quotation marks (smart quotes), and fix any broken numbering." },
    ]
  },
  {
    label: "🔍 Review",
    actions: [
      { name: "Summarize", prompt: "Provide a concise summary of this document, including: (1) the type of document, (2) the parties involved, (3) the key terms and obligations, and (4) any notable provisions or risks." },
      { name: "Ambiguities", prompt: "Identify any ambiguous language in this document that could lead to disputes or multiple interpretations. For each instance, explain the ambiguity and suggest clearer language." },
      { name: "Missing Definitions", prompt: "Find any terms in this document that are used as if they are defined terms (e.g., capitalized nouns) but lack a formal definition. List them with suggested definitions." },
      { name: "Risk Analysis", prompt: "Analyze this document for potential legal risks, unfavorable terms, or provisions that may be problematic. Prioritize by severity." },
    ]
  },
];

function getBuiltinOverrides(): Record<string, string> {
  try { return JSON.parse(localStorage.getItem("sidebar-builtin-overrides") || "{}"); } catch { return {}; }
}
function saveBuiltinOverrides(overrides: Record<string, string>): void {
  localStorage.setItem("sidebar-builtin-overrides", JSON.stringify(overrides));
}

function getCustomPrompts(): QuickAction[] {
  try { return JSON.parse(localStorage.getItem("sidebar-custom-prompts") || "[]"); } catch { return []; }
}
function saveCustomPrompts(prompts: QuickAction[]): void {
  localStorage.setItem("sidebar-custom-prompts", JSON.stringify(prompts));
  // Also sync to server (best-effort)
  fetch("http://localhost:3001/api/prompts/custom", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(prompts),
  }).catch(() => {});
}

function getEffectivePrompt(actionName: string, defaultPrompt: string): string {
  const overrides = getBuiltinOverrides();
  return overrides[actionName] ?? defaultPrompt;
}

function getDefaultPrompt(actionName: string): string | null {
  for (const cat of QUICK_ACTIONS_DEFAULTS) {
    for (const a of cat.actions) {
      if (a.name === actionName) return a.prompt;
    }
  }
  return null;
}

function isBuiltinAction(actionName: string): boolean {
  return getDefaultPrompt(actionName) !== null;
}

function isBuiltinModified(actionName: string): boolean {
  const overrides = getBuiltinOverrides();
  return actionName in overrides;
}

async function resolvePromptVariables(prompt: string): Promise<string> {
  let resolved = prompt;
  if (resolved.includes("{{selection}}")) {
    try {
      const sel = await Word.run(async (ctx) => {
        const s = ctx.document.getSelection();
        s.load("text");
        await ctx.sync();
        return s.text;
      });
      resolved = resolved.replace(/\{\{selection\}\}/g, sel || "(no text selected)");
    } catch {
      resolved = resolved.replace(/\{\{selection\}\}/g, "(unable to read selection)");
    }
  }
  resolved = resolved.replace(/\{\{document\}\}/g, "[full document context included automatically]");
  return resolved;
}

async function executeQuickAction(prompt: string): Promise<void> {
  const resolved = await resolvePromptVariables(prompt);
  const input = document.getElementById("prompt-input") as HTMLTextAreaElement;
  if (input) {
    input.value = resolved;
    input.dispatchEvent(new Event("input")); // trigger auto-resize
  }
  const sendBtn = document.getElementById("prompt-send") as HTMLButtonElement;
  if (sendBtn) sendBtn.click();
}

// ── Modal State ──
let qaModalMode: "edit-builtin" | "edit-custom" | "add-custom" = "add-custom";
let qaModalActionName: string | null = null;
let qaModalCustomIndex: number = -1;

function openPromptEditor(opts: {
  mode: typeof qaModalMode;
  name?: string;
  prompt?: string;
  defaultPrompt?: string;
  customIndex?: number;
}): void {
  qaModalMode = opts.mode;
  qaModalActionName = opts.name || null;
  qaModalCustomIndex = opts.customIndex ?? -1;

  const overlay = document.getElementById("qa-modal-overlay");
  const titleEl = document.getElementById("qa-modal-title") as HTMLElement;
  const nameInput = document.getElementById("qa-modal-name") as HTMLInputElement;
  const promptInput = document.getElementById("qa-modal-prompt") as HTMLTextAreaElement;
  const resetBtn = document.getElementById("qa-modal-reset") as HTMLButtonElement;

  if (!overlay) return;

  if (opts.mode === "add-custom") {
    titleEl.textContent = "New Custom Prompt";
    nameInput.value = "";
    promptInput.value = "";
    nameInput.readOnly = false;
    resetBtn.style.display = "none";
  } else if (opts.mode === "edit-builtin") {
    titleEl.textContent = "Edit: " + (opts.name || "");
    nameInput.value = opts.name || "";
    nameInput.readOnly = true;
    promptInput.value = opts.prompt || opts.defaultPrompt || "";
    resetBtn.style.display = isBuiltinModified(opts.name || "") ? "inline-block" : "none";
  } else {
    titleEl.textContent = "Edit: " + (opts.name || "");
    nameInput.value = opts.name || "";
    nameInput.readOnly = false;
    promptInput.value = opts.prompt || "";
    resetBtn.style.display = "none";
  }

  overlay.classList.add("visible");
}

function closePromptEditor(): void {
  document.getElementById("qa-modal-overlay")?.classList.remove("visible");
}

function setupQuickActions(): void {
  const bar = document.getElementById("quick-actions-bar");
  if (!bar) return;

  function closeAllDropdowns() {
    bar.querySelectorAll(".qa-dropdown.visible").forEach(d => d.classList.remove("visible"));
    bar.querySelectorAll(".qa-pill.active").forEach(p => p.classList.remove("active"));
  }

  function renderBar() {
    bar.innerHTML = "";
    const overrides = getBuiltinOverrides();
    const customPrompts = getCustomPrompts();

    // Built-in categories
    for (const cat of QUICK_ACTIONS_DEFAULTS) {
      const pill = document.createElement("div");
      pill.className = "qa-pill";
      pill.textContent = cat.label;

      const dropdown = document.createElement("div");
      dropdown.className = "qa-dropdown";

      for (const action of cat.actions) {
        const effectivePrompt = getEffectivePrompt(action.name, action.prompt);
        const modified = isBuiltinModified(action.name);

        const row = document.createElement("div");
        row.className = "qa-action";

        const nameSpan = document.createElement("span");
        nameSpan.className = "qa-action-name";
        nameSpan.textContent = action.name;
        nameSpan.addEventListener("click", (e) => {
          e.stopPropagation();
          closeAllDropdowns();
          executeQuickAction(effectivePrompt);
        });

        const icons = document.createElement("span");
        icons.className = "qa-action-icons";
        if (modified) {
          const dot = document.createElement("span");
          dot.className = "qa-modified-dot";
          dot.title = "Customized";
          icons.appendChild(dot);
        }
        const editBtn = document.createElement("span");
        editBtn.className = "qa-action-edit";
        editBtn.textContent = "✏️";
        editBtn.title = "Edit prompt";
        editBtn.addEventListener("click", (e) => {
          e.stopPropagation();
          closeAllDropdowns();
          openPromptEditor({
            mode: "edit-builtin",
            name: action.name,
            prompt: effectivePrompt,
            defaultPrompt: action.prompt,
          });
        });
        icons.appendChild(editBtn);

        row.appendChild(nameSpan);
        row.appendChild(icons);
        dropdown.appendChild(row);
      }

      pill.appendChild(dropdown);
      pill.addEventListener("click", (e) => {
        if ((e.target as HTMLElement).closest(".qa-action")) return;
        const wasActive = pill.classList.contains("active");
        closeAllDropdowns();
        if (!wasActive) {
          pill.classList.add("active");
          dropdown.classList.add("visible");
        }
      });
      bar.appendChild(pill);
    }

    // Custom category
    const customPill = document.createElement("div");
    customPill.className = "qa-pill";
    customPill.textContent = "⭐ Custom";

    const customDropdown = document.createElement("div");
    customDropdown.className = "qa-dropdown";

    for (let i = 0; i < customPrompts.length; i++) {
      const cp = customPrompts[i];
      const row = document.createElement("div");
      row.className = "qa-action";

      const nameSpan = document.createElement("span");
      nameSpan.className = "qa-action-name";
      nameSpan.textContent = cp.name;
      nameSpan.addEventListener("click", (e) => {
        e.stopPropagation();
        closeAllDropdowns();
        executeQuickAction(cp.prompt);
      });

      const icons = document.createElement("span");
      icons.className = "qa-action-icons";
      const editBtn = document.createElement("span");
      editBtn.className = "qa-action-edit";
      editBtn.textContent = "✏️";
      editBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        closeAllDropdowns();
        openPromptEditor({ mode: "edit-custom", name: cp.name, prompt: cp.prompt, customIndex: i });
      });
      const delBtn = document.createElement("span");
      delBtn.className = "qa-action-delete";
      delBtn.textContent = "✕";
      delBtn.title = "Delete";
      delBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        const prompts = getCustomPrompts();
        prompts.splice(i, 1);
        saveCustomPrompts(prompts);
        renderBar();
      });
      icons.appendChild(editBtn);
      icons.appendChild(delBtn);

      row.appendChild(nameSpan);
      row.appendChild(icons);
      customDropdown.appendChild(row);
    }

    // Add custom prompt button
    const addRow = document.createElement("div");
    addRow.className = "qa-add-custom";
    if (customPrompts.length > 0) addRow.classList.add("qa-custom-actions");
    addRow.textContent = "+ Add Custom Prompt";
    addRow.addEventListener("click", (e) => {
      e.stopPropagation();
      closeAllDropdowns();
      openPromptEditor({ mode: "add-custom" });
    });
    customDropdown.appendChild(addRow);

    customPill.appendChild(customDropdown);
    customPill.addEventListener("click", (e) => {
      if ((e.target as HTMLElement).closest(".qa-action, .qa-add-custom")) return;
      const wasActive = customPill.classList.contains("active");
      closeAllDropdowns();
      if (!wasActive) {
        customPill.classList.add("active");
        customDropdown.classList.add("visible");
      }
    });
    bar.appendChild(customPill);
  }

  // Close dropdowns when clicking outside
  document.addEventListener("click", (e) => {
    if (!(e.target as HTMLElement).closest(".qa-pill")) {
      closeAllDropdowns();
    }
  });

  // Modal handlers
  document.getElementById("qa-modal-cancel")?.addEventListener("click", closePromptEditor);
  document.getElementById("qa-modal-overlay")?.addEventListener("click", (e) => {
    if (e.target === document.getElementById("qa-modal-overlay")) closePromptEditor();
  });

  document.getElementById("qa-modal-reset")?.addEventListener("click", () => {
    if (qaModalMode === "edit-builtin" && qaModalActionName) {
      const overrides = getBuiltinOverrides();
      delete overrides[qaModalActionName];
      saveBuiltinOverrides(overrides);
      closePromptEditor();
      renderBar();
    }
  });

  document.getElementById("qa-modal-save")?.addEventListener("click", () => {
    const nameInput = document.getElementById("qa-modal-name") as HTMLInputElement;
    const promptInput = document.getElementById("qa-modal-prompt") as HTMLTextAreaElement;
    const name = nameInput.value.trim();
    const prompt = promptInput.value.trim();
    if (!name || !prompt) return;

    if (qaModalMode === "edit-builtin" && qaModalActionName) {
      const defaultPrompt = getDefaultPrompt(qaModalActionName);
      const overrides = getBuiltinOverrides();
      if (prompt === defaultPrompt) {
        delete overrides[qaModalActionName];
      } else {
        overrides[qaModalActionName] = prompt;
      }
      saveBuiltinOverrides(overrides);
    } else if (qaModalMode === "edit-custom" && qaModalCustomIndex >= 0) {
      const prompts = getCustomPrompts();
      if (qaModalCustomIndex < prompts.length) {
        prompts[qaModalCustomIndex] = { name, prompt };
        saveCustomPrompts(prompts);
      }
    } else if (qaModalMode === "add-custom") {
      const prompts = getCustomPrompts();
      prompts.push({ name, prompt });
      saveCustomPrompts(prompts);
    }

    closePromptEditor();
    renderBar();
  });

  // Load custom prompts from server as backup (merge missing ones)
  fetch("http://localhost:3001/api/prompts/custom").then(r => r.json()).then(j => {
    if (j?.ok && Array.isArray(j.data) && j.data.length > 0) {
      const local = getCustomPrompts();
      const localNames = new Set(local.map(p => p.name));
      let added = false;
      for (const sp of j.data) {
        if (!localNames.has(sp.name)) {
          local.push(sp);
          added = true;
        }
      }
      if (added) {
        localStorage.setItem("sidebar-custom-prompts", JSON.stringify(local));
        renderBar();
      }
    }
  }).catch(() => {});

  renderBar();
}
