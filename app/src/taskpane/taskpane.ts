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


// ── Smooth Auto-Scroll (throttled via rAF) ──
let scrollPending = false;
function smoothScroll(el: HTMLElement) {
  if (!scrollPending) {
    scrollPending = true;
    requestAnimationFrame(() => {
      el.scrollTo({ top: el.scrollHeight, behavior: 'smooth' });
      scrollPending = false;
    });
  }
}

let thinkingTimerInterval: any = null;
let queryInProgress = false;

function setQueryInProgress(active: boolean): void {
  queryInProgress = active;
  const sendBtn = document.getElementById("prompt-send") as HTMLButtonElement | null;
  if (!sendBtn) return;
  if (active) {
    sendBtn.textContent = "■";
    sendBtn.title = "Stop";
    sendBtn.classList.add("stop-mode");
  } else {
    sendBtn.textContent = "↑";
    sendBtn.title = "Send";
    sendBtn.classList.remove("stop-mode");
  }
}

function addThinkingIndicator(): HTMLElement {
  const history = document.getElementById("prompt-history")!;
  setQueryInProgress(true);
  const el = document.createElement("div");
  el.className = "chat-thinking";
  el.id = "thinking-indicator";
  el.innerHTML = '<div class="thinking-dots"><div class="dot"></div><div class="dot"></div><div class="dot"></div></div><span class="thinking-elapsed"></span>';
  // Stop button is now in the send button area
  const legacyStop = document.createElement("span"); // hidden placeholder
  legacyStop.addEventListener("click", () => {
    // Abort current request via server endpoint
    fetch("http://localhost:3001/api/prompts/abort", { method: "POST" }).catch(() => {});
    removeThinkingIndicator();
    // Show cancelled message
    const hist = document.getElementById("prompt-history");
    if (hist) {
      const cancelEl = document.createElement("div");
      cancelEl.className = "chat-entry chat-assistant";
      cancelEl.innerHTML = '<div class="chat-role">The Sidebar</div><div class="chat-text" style="opacity:0.5;font-style:italic">Request cancelled.</div>';
      hist.appendChild(cancelEl);
      hist.scrollTo({ top: hist.scrollHeight, behavior: "smooth" });
    }
  });
  history.appendChild(el);
  history.scrollTop = history.scrollHeight;

  // Start elapsed timer
  const startTime = Date.now();
  const elapsedEl = el.querySelector(".thinking-elapsed") as HTMLElement;
  if (thinkingTimerInterval) clearInterval(thinkingTimerInterval);
  thinkingTimerInterval = setInterval(() => {
    const seconds = Math.floor((Date.now() - startTime) / 1000);
    if (seconds < 5) {
      elapsedEl.textContent = "";
    } else if (seconds < 60) {
      elapsedEl.textContent = `${seconds}s`;
    } else {
      const min = Math.floor(seconds / 60);
      const sec = seconds % 60;
      elapsedEl.textContent = `${min}m ${sec.toString().padStart(2, "0")}s`;
    }
  }, 1000);

  return el;
}

function removeThinkingIndicator(): void {
  setQueryInProgress(false);
  if (thinkingTimerInterval) { clearInterval(thinkingTimerInterval); thinkingTimerInterval = null; }
  const el = document.getElementById("thinking-indicator");
  if (el) el.remove();
}




function resolveUserEntryForPrompt(promptId: number): HTMLElement | null {
  const history = document.getElementById("prompt-history")!;
  let userEl = history.querySelector(`[data-prompt-id="${promptId}"]`) as HTMLElement | null;
  if (userEl) return userEl;

  // Fallback: if progress arrives before ack, attach this promptId to the most recent pending user bubble.
  const pending = Array.from(pendingPromptEls.values());
  if (pending.length > 0) {
    userEl = pending[pending.length - 1] as HTMLElement;
    userEl.setAttribute("data-prompt-id", String(promptId));
    return userEl;
  }

  return null;
}


function compactStreamingPreview(text: string): string {
  const cleaned = (text || "").replace(/\s+/g, " ").trim();
  if (!cleaned) return "Drafting…";
  if (cleaned.length <= 140) return cleaned;
  return "…" + cleaned.slice(-140);
}

function summarizeStreamingState(progressLabel?: string, progressText?: string): string {
  if (progressLabel && progressLabel.trim()) return progressLabel.trim();
  if (progressText && progressText.trim()) return compactStreamingPreview(progressText);
  return "Working…";
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
    setupTrayIcons();
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

/** Initialize or resume a document session (with retry for slow doc loads) */
async function initSession(retries = 3, delayMs = 1500): Promise<void> {
  for (let attempt = 1; attempt <= retries; attempt++) {
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
          console.log("[session] Resumed:", existingId, `(attempt ${attempt})`);
          return;
        }
      } catch (e) {
        console.error("[session] Resume failed:", e);
      }
    }
    // Doc property not ready yet - wait and retry
    if (attempt < retries) {
      console.log(`[session] No session ID found, retrying in ${delayMs}ms (attempt ${attempt}/${retries})`);
      await new Promise(resolve => setTimeout(resolve, delayMs));
    }
  }
  // All retries exhausted - start fresh session
  console.log("[session] No existing session found after retries, creating new");
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
    
    // Get current selection + stable anchor metadata for context
    let context = "";
    try {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text,isEmpty");
        const selPara = sel.paragraphs.getFirstOrNullObject();
        selPara.load("text,isListItem");
        const selLi = selPara.listItemOrNullObject;
        selLi.load("listString");

        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();

        let paragraphIndex = -1;
        if (!selPara.isNullObject) {
          for (let i = 0; i < paragraphs.items.length; i++) {
            if (paragraphs.items[i] === selPara) {
              paragraphIndex = i;
              break;
            }
          }
        }

        if (!sel.isEmpty && sel.text) {
          const payload: any = {
            selectedText: sel.text.substring(0, 1000),
            selectionParagraphIndex: paragraphIndex >= 0 ? paragraphIndex : undefined,
            selectionParagraphListString: !selLi.isNullObject ? selLi.listString : undefined,
            selectionParagraphPrefix: !selPara.isNullObject ? (selPara.text || "").substring(0, 120) : undefined,
            selectionAnchorRule: "Treat selectionParagraphIndex/listString as authoritative anchor. Do not locate by prefix search.",
          };
          context = JSON.stringify(payload);
        }
      });
    } catch {}
    
    const clientId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const model = modelSelect?.value || selectedModel || undefined;
    socket.send(JSON.stringify({ type: "prompt", clientId, text, model, context: context || undefined }));

    // Show clean label for quick actions, full text for manual prompts
    const displayText = (input as any)._displayLabel || text;
    delete (input as any)._displayLabel;
    const el = appendChatEntry("user", displayText, "Sending…");
    el.setAttribute("data-client-id", clientId);
    pendingPromptEls.set(clientId, el);
    input.value = "";
  };

  btn.addEventListener("click", () => {
    if (queryInProgress) {
      // Stop the current query
      fetch("http://localhost:3001/api/prompts/abort", { method: "POST" }).catch(() => {});
      removeThinkingIndicator();
      const hist = document.getElementById("prompt-history");
      if (hist) {
        const cancelEl = document.createElement("div");
        cancelEl.className = "chat-entry chat-assistant";
        cancelEl.innerHTML = '<div class="chat-role">The Sidebar</div><div class="chat-text" style="opacity:0.5;font-style:italic">Request cancelled.</div>';
        hist.appendChild(cancelEl);
      }
      return;
    }
    sendPrompt();
  });
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
        const toolStatus = msg.status as string | undefined;
        const toolName = msg.toolName as string | undefined;
        if (promptId === undefined) return;

        const history = document.getElementById("prompt-history")!;

        // Handle reasoning content from thinking models (DeepSeek, QwQ, etc.)
        if (toolStatus === "reasoning") {
          removeThinkingIndicator();
          const reasoningContent = msg.content as string | undefined;
          if (!reasoningContent) return;
          // Get or create the activity block
          let activityBlock = history.querySelector(`[data-activity-for="${promptId}"]`) as HTMLElement | null;
          if (!activityBlock) {
            activityBlock = document.createElement("div");
            activityBlock.className = "activity-block";
            activityBlock.setAttribute("data-activity-for", String(promptId));
            const header = document.createElement("div");
            header.className = "activity-header";
            header.innerHTML = '<span class="activity-icon">💭</span> <span class="activity-label">Thinking…</span>';
            header.addEventListener("click", () => activityBlock!.classList.toggle("collapsed"));
            activityBlock.appendChild(header);
            const body = document.createElement("div");
            body.className = "activity-body";
            activityBlock.appendChild(body);
            const userEl = resolveUserEntryForPrompt(promptId);
            if (userEl) userEl.insertAdjacentElement("afterend", activityBlock);
            else history.appendChild(activityBlock);
          }
          let body = activityBlock.querySelector(".activity-body") as HTMLElement;
          let reasoningEl = body.querySelector(".reasoning-content") as HTMLElement;
          if (!reasoningEl) {
            reasoningEl = document.createElement("div");
            reasoningEl.className = "reasoning-content";
            body.insertBefore(reasoningEl, body.firstChild);
          }
          // Truncate display to last ~500 chars for performance, full content scrollable
          const maxDisplay = 500;
          const display = reasoningContent.length > maxDisplay
            ? "…" + reasoningContent.slice(-maxDisplay)
            : reasoningContent;
          reasoningEl.textContent = display;
          reasoningEl.title = "Model reasoning (click header to collapse)";
          smoothScroll(history);
          return;
        }

        // Handle tool phase — model sent text but is now executing tools
        if (toolStatus === "tool_phase") {
          // Re-add thinking indicator with "Executing..." label
          removeThinkingIndicator();
          const indicator = addThinkingIndicator();
          const elapsed = indicator.querySelector(".thinking-elapsed") as HTMLElement;
          if (elapsed) elapsed.textContent = "Executing edits...";
          return;
        }

        // Handle tool execution progress — put in activity block
        if (toolStatus === "tool" && toolName) {
          removeThinkingIndicator();
          // Get or create activity block
          let activityBlock = history.querySelector(`[data-activity-for="${promptId}"]`) as HTMLElement | null;
          if (!activityBlock) {
            activityBlock = document.createElement("div");
            activityBlock.className = "activity-block";
            activityBlock.setAttribute("data-activity-for", String(promptId));
            const header = document.createElement("div");
            header.className = "activity-header";
            header.innerHTML = '<span class="activity-icon">⚙️</span> <span class="activity-label">Working…</span>';
            header.addEventListener("click", () => activityBlock!.classList.toggle("collapsed"));
            activityBlock.appendChild(header);
            const body = document.createElement("div");
            body.className = "activity-body";
            activityBlock.appendChild(body);
            const userEl = resolveUserEntryForPrompt(promptId);
            if (userEl) userEl.insertAdjacentElement("afterend", activityBlock);
            else history.appendChild(activityBlock);
          }
          // Update header to show tool activity
          const label = activityBlock.querySelector(".activity-label") as HTMLElement;
          if (label) label.textContent = "Working…";
          const icon = activityBlock.querySelector(".activity-icon") as HTMLElement;
          if (icon) icon.textContent = "⚙️";

          const body = activityBlock.querySelector(".activity-body") as HTMLElement;
          let toolContainer = body.querySelector(".tool-progress-container") as HTMLElement;
          if (!toolContainer) {
            toolContainer = document.createElement("div");
            toolContainer.className = "tool-progress-container";
            body.appendChild(toolContainer);
          }
          const line = document.createElement("div");
          line.className = "tool-progress-line";
          line.setAttribute("data-tool", toolName);
          line.innerHTML = `<span class="spinner">⟳</span> ${escapeHtml(progressLabel || toolName)}`;
          toolContainer.appendChild(line);
          smoothScroll(history);
          return;
        }

        if (toolStatus === "tool_complete" && toolName) {
          const activityBlock = history.querySelector(`[data-activity-for="${promptId}"]`) as HTMLElement | null;
          if (activityBlock) {
            const lines = activityBlock.querySelectorAll(`.tool-progress-line[data-tool="${toolName}"]:not(.complete)`);
            const line = lines[lines.length - 1] as HTMLElement | undefined;
            if (line) {
              line.classList.add("complete");
              const spinner = line.querySelector(".spinner");
              if (spinner) spinner.textContent = "✓";
            }
          }
          // Also check old-style streamEl for backward compat
          const streamEl = history.querySelector(`[data-streaming-for="${promptId}"]`) as HTMLElement | null;
          if (streamEl) {
            const lines = streamEl.querySelectorAll(`.tool-progress-line[data-tool="${toolName}"]:not(.complete)`);
            const line = lines[lines.length - 1] as HTMLElement | undefined;
            if (line) line.classList.add("complete");
          }
          return;
        }

        // If we have a descriptive progress label (tool activity), show it in thinking indicator
        if (progressLabel && !progressText) {
          const thinkingEl = document.getElementById("thinking-indicator");
          if (thinkingEl) {
            let labelEl = thinkingEl.querySelector(".progress-text") as HTMLElement;
            if (!labelEl) {
              thinkingEl.innerHTML = "";
              labelEl = document.createElement("span");
              labelEl.className = "progress-text";
              thinkingEl.appendChild(labelEl);
            }
            const newLabel = labelEl.cloneNode(false) as HTMLElement;
            newLabel.textContent = progressLabel;
            labelEl.replaceWith(newLabel);
          }
        }

        if (!progressText) return;
        // Update or create streaming entry — use plain text during streaming for speed
        let streamEl = history.querySelector(`[data-streaming-for="${promptId}"]`) as HTMLElement | null;
        if (!streamEl) {
          removeThinkingIndicator();
          streamEl = document.createElement("div");
          streamEl.className = "chat-entry chat-assistant";
          streamEl.setAttribute("data-streaming-for", String(promptId));
          const roleEl = document.createElement("div");
          roleEl.className = "chat-role";
          roleEl.textContent = "Assistant";
          const textEl = document.createElement("div");
          textEl.className = "chat-text streaming-cursor";
          textEl.textContent = summarizeStreamingState(progressLabel, progressText);
          streamEl.appendChild(roleEl);
          streamEl.appendChild(textEl);
          const toolContainer = document.createElement("div");
          toolContainer.className = "tool-progress-container";
          streamEl.appendChild(toolContainer);
          const userEl = resolveUserEntryForPrompt(promptId);
          if (userEl) userEl.insertAdjacentElement("afterend", streamEl);
          else history.appendChild(streamEl);
        } else {
          const textEl = streamEl.querySelector(".chat-text") as HTMLElement;
          if (textEl) {
            textEl.textContent = summarizeStreamingState(progressLabel, progressText);
            textEl.classList.add("streaming-cursor");
          }
        }
        smoothScroll(history);
        return;
      }
      if (msg.type === "prompt_response") {
        // Refresh memory count (extraction happens async after each exchange)
        setTimeout(refreshMemoryCount, 3000);
        removeThinkingIndicator();
        const promptId = msg.promptId as number | undefined;
        const responseText = msg.text as string | undefined;
        if (!responseText) return;
        // Collapse the activity block (reasoning + tool calls)
        if (promptId !== undefined) {
          const activityBlock = document.getElementById("prompt-history")?.querySelector(`[data-activity-for="${promptId}"]`) as HTMLElement | null;
          if (activityBlock) {
            activityBlock.classList.add("collapsed", "done");
            const label = activityBlock.querySelector(".activity-label") as HTMLElement;
            // Summarize: count tool lines
            const toolCount = activityBlock.querySelectorAll(".tool-progress-line").length;
            const hasReasoning = !!activityBlock.querySelector(".reasoning-content");
            const parts: string[] = [];
            if (hasReasoning) parts.push("reasoned");
            if (toolCount > 0) parts.push(`${toolCount} tool${toolCount > 1 ? "s" : ""}`);
            if (label && parts.length) label.textContent = parts.join(" + ");
            else if (label) label.textContent = "Done";
          }
        }
        // Capture tool progress lines from streaming preview before removing
        let toolProgressHtml = "";
        if (promptId !== undefined) {
          const streamEl = document.getElementById("prompt-history")?.querySelector(`[data-streaming-for="${promptId}"]`);
          if (streamEl) {
            const toolContainer = streamEl.querySelector(".tool-progress-container");
            if (toolContainer && toolContainer.children.length > 0) {
              toolProgressHtml = toolContainer.outerHTML;
            }
            streamEl.remove();
          }
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
            // Re-attach tool progress lines from streaming
            if (toolProgressHtml) {
              const temp = document.createElement("div");
              temp.innerHTML = toolProgressHtml;
              const container = temp.firstElementChild;
              if (container) {
                // Mark all lines as complete
                container.querySelectorAll(".tool-progress-line:not(.complete)").forEach(l => l.classList.add("complete"));
                el.appendChild(container);
              }
            }
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
      throw new Error("replaceSelection is disabled for safety. Use anchored edits (replaceParagraph/findReplace) to avoid selection drift.");

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

        const p = paragraphs.items[paraIndex];
        p.load("text,style");
        await ctx.sync();

        const original = p.text;
        const desired = normalizeSmartQuotes(text);
        if (original === desired) return { index: paraIndex, changed: false };

        // Compute minimal middle replacement window.
        let prefixLen = 0;
        while (prefixLen < original.length && prefixLen < desired.length && original[prefixLen] === desired[prefixLen]) prefixLen++;

        let suffixLen = 0;
        while (
          suffixLen < (original.length - prefixLen) &&
          suffixLen < (desired.length - prefixLen) &&
          original[original.length - 1 - suffixLen] === desired[desired.length - 1 - suffixLen]
        ) suffixLen++;

        const oldMiddle = original.substring(prefixLen, original.length - suffixLen);
        const newMiddle = desired.substring(prefixLen, desired.length - suffixLen);
        const trackMode = !!trackChangesMode;

        // Strict tracked-edit guardrail: block broad rewrites in Track mode.
        const changedChars = oldMiddle.length;
        const totalChars = Math.max(1, original.length);
        const changedRatio = changedChars / totalChars;
        const isBroadRewrite = changedChars === totalChars || changedRatio > 0.6;
        if (trackMode && isBroadRewrite) {
          throw new Error("Track mode blocks broad paragraph rewrites. Please apply smaller sentence/phrase edits.");
        }

        if (oldMiddle.length > 0 && oldMiddle.length < original.length) {
          // Best case: patch only the changed span.
          const searchResults = p.search(oldMiddle, { matchCase: true });
          searchResults.load("items");
          await ctx.sync();
          if (searchResults.items.length > 0) {
            searchResults.items[0].insertText(newMiddle, Word.InsertLocation.replace);
            await ctx.sync();
          } else if (!trackMode) {
            const range = p.getRange(Word.RangeLocation.content);
            range.insertText(desired, Word.InsertLocation.replace);
            await ctx.sync();
          } else {
            throw new Error("Could not apply a granular tracked edit. Please retry with a narrower change.");
          }
        } else if (!trackMode) {
          // YOLO mode: allow full paragraph replacement.
          const range = p.getRange(Word.RangeLocation.content);
          range.insertText(desired, Word.InsertLocation.replace);
          await ctx.sync();
        } else {
          // Track mode: avoid one-shot full replacement. Do delete + insert as separate revisions.
          const contentRange = p.getRange(Word.RangeLocation.content);
          contentRange.load("text");
          await ctx.sync();

          if (contentRange.text && contentRange.text.length > 0) {
            contentRange.insertText("", Word.InsertLocation.replace);
            await ctx.sync();
          }

          const refreshed = p.getRange(Word.RangeLocation.content);
          refreshed.insertText(desired, Word.InsertLocation.start);
          await ctx.sync();
        }

        pushUndo({ command: "replaceParagraph", original, replacement: desired, paragraphIndex: paraIndex, style: p.style, timestamp: Date.now() });
        return { original, replacement: desired, paragraphIndex: paraIndex, undoAvailable: true, trackMode };
      });

    case "editSelection":
      throw new Error("editSelection is disabled for safety. Use anchored edits (replaceParagraph/findReplace) to avoid selection drift.");

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
        const normalizedReplacement = normalizeSmartQuotes(replacement);
        const results = ctx.document.body.search(text, { matchCase: matchCase ?? false, matchWholeWord: matchWholeWord ?? false });
        results.load("items"); await ctx.sync();
        const count = results.items.length;
        for (const item of results.items) item.insertText(normalizedReplacement, Word.InsertLocation.replace);
        await ctx.sync();
        return { replacedCount: count };
      });

    case "insert":
      return Word.run(async (ctx) => {
        const { text, location, paragraphIndex, style, allowEnd } = params || {};
        if (!text) throw new Error("params.text required");
        let paragraph: Word.Paragraph;
        if (paragraphIndex !== undefined) {
          const paragraphs = ctx.document.body.paragraphs; paragraphs.load("items"); await ctx.sync();
          if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) throw new Error(`paragraphIndex out of range`);
          paragraph = paragraphs.items[paragraphIndex].insertParagraph(normalizeSmartQuotes(text), location === "before" ? Word.InsertLocation.before : Word.InsertLocation.after);
        } else if (location === "start") {
          paragraph = ctx.document.body.insertParagraph(normalizeSmartQuotes(text), Word.InsertLocation.start);
        } else if (location === "end" && allowEnd === true) {
          paragraph = ctx.document.body.insertParagraph(normalizeSmartQuotes(text), Word.InsertLocation.end);
        } else {
          throw new Error('Unsafe insert blocked: provide paragraphIndex (+ before/after) for targeted insertion. End-of-document insert requires { location: "end", allowEnd: true }.');
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
        if (text !== undefined) p.insertText(normalizeSmartQuotes(text), Word.InsertLocation.replace);
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
        footnotes.items[index].body.insertText(normalizeSmartQuotes(text), Word.InsertLocation.replace);
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

    // ── Tables ──
    case "getTables":
      return Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load("count");
        await ctx.sync();
        const result: any[] = [];
        for (let i = 0; i < tables.count; i++) {
          const t = tables.items[i];
          t.load("rowCount,headerRowCount");
        }
        await ctx.sync();
        return { count: tables.count, tables: tables.items.map((t, i) => ({ index: i, rowCount: t.rowCount, headerRows: t.headerRowCount })) };
      });

    case "readTable":
      return Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load("count");
        await ctx.sync();
        const idx = params?.index ?? 0;
        if (idx >= tables.count) throw new Error("Table index out of range");
        const table = tables.items[idx];
        table.load("rowCount,values,headerRowCount");
        await ctx.sync();
        return { rowCount: table.rowCount, headerRowCount: table.headerRowCount, values: table.values };
      });

    case "insertTable":
      return Word.run(async (ctx) => {
        const body = ctx.document.body;
        const rows = params?.rows || 2;
        const cols = params?.columns || 2;
        const table = body.insertTable(rows, cols, Word.InsertLocation.end, params?.values || []);
        if (params?.style) table.styleBuiltIn = params.style;
        await ctx.sync();
        return { success: true, rows, columns: cols };
      });

    case "updateTableCell":
      return Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load("count");
        await ctx.sync();
        const table = tables.items[params?.tableIndex || 0];
        const cell = table.getCell(params.row, params.column);
        cell.body.clear();
        cell.body.insertText(normalizeSmartQuotes(params.text), Word.InsertLocation.start);
        await ctx.sync();
        return { success: true };
      });

    case "addTableRow":
      return Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load("count");
        await ctx.sync();
        const table = tables.items[params?.tableIndex || 0];
        table.addRows(params?.position === "start" ? Word.InsertLocation.start : Word.InsertLocation.end, 1, params?.values ? [params.values] : []);
        await ctx.sync();
        return { success: true };
      });

    case "addTableColumn":
      return Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load("count");
        await ctx.sync();
        const table = tables.items[params?.tableIndex || 0];
        table.addColumns(params?.position === "start" ? Word.InsertLocation.start : Word.InsertLocation.end, 1, params?.values ? [params.values] : []);
        await ctx.sync();
        return { success: true };
      });

    // ── Headers & Footers ──
    case "getHeaderFooter":
      return Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();
        const section = sections.items[params?.sectionIndex || 0];
        const headerType = params?.headerType === "firstPage" ? Word.HeaderFooterType.firstPage
          : params?.headerType === "evenPages" ? Word.HeaderFooterType.evenPages
          : Word.HeaderFooterType.primary;
        const hf = params?.type === "footer"
          ? section.getFooter(headerType)
          : section.getHeader(headerType);
        hf.load("text");
        await ctx.sync();
        return { text: hf.text, type: params?.type || "header", headerType: params?.headerType || "primary" };
      });

    case "setHeaderFooter":
      return Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();
        const section = sections.items[params?.sectionIndex || 0];
        const headerType = params?.headerType === "firstPage" ? Word.HeaderFooterType.firstPage
          : params?.headerType === "evenPages" ? Word.HeaderFooterType.evenPages
          : Word.HeaderFooterType.primary;
        const hf = params?.type === "footer"
          ? section.getFooter(headerType)
          : section.getHeader(headerType);
        hf.clear();
        hf.insertText(normalizeSmartQuotes(params.text), Word.InsertLocation.start);
        await ctx.sync();
        return { success: true };
      });

    // ── Delete Paragraph ──
    case "deleteParagraph":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const idx = params?.index;
        if (idx === undefined || idx < 0 || idx >= paragraphs.items.length) throw new Error("Index out of range");
        paragraphs.items[idx].delete();
        await ctx.sync();
        return { success: true };
      });

    // ── Breaks ──
    case "insertBreak":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const afterIdx = params?.afterParagraph ?? paragraphs.items.length - 1;
        const para = paragraphs.items[afterIdx];
        const breakType = params?.breakType === "section" ? Word.BreakType.sectionNext
          : params?.breakType === "sectionContinuous" ? Word.BreakType.sectionContinuous
          : Word.BreakType.page;
        para.insertBreak(breakType, Word.InsertLocation.after);
        await ctx.sync();
        return { success: true };
      });

    // ── Lists ──
    case "setListFormat":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const from = params?.fromIndex;
        const to = params?.toIndex ?? from;
        if (from === undefined) throw new Error("fromIndex required");
        for (let i = from; i <= to && i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          if (params?.type === "bullet") {
            para.startNewList();
          } else if (params?.type === "numbered") {
            para.startNewList();
          } else if (params?.type === "none") {
            try { para.detachFromList(); } catch {}
          }
        }
        await ctx.sync();
        return { success: true };
      });

    // ── Bookmarks ──
    case "getBookmarks":
      return Word.run(async (ctx) => {
        try {
          const bookmarks = ctx.document.body.getRange().getBookmarks();
          await ctx.sync();
          return { bookmarks: bookmarks.value };
        } catch {
          return { bookmarks: [], note: "Bookmarks API not available in this Word version" };
        }
      });

    // ── Highlight & Font Color ──
    case "highlightText":
      return Word.run(async (ctx) => {
        const results = ctx.document.body.search(params.text, { matchCase: params?.matchCase || false });
        results.load("items");
        await ctx.sync();
        for (const item of results.items) {
          item.font.highlightColor = params?.color || "yellow";
        }
        await ctx.sync();
        return { count: results.items.length };
      });

    case "setFontColor":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const para = paragraphs.items[params.index];
        if (params?.text) {
          const results = para.search(params.text, { matchCase: true });
          results.load("items");
          await ctx.sync();
          for (const r of results.items) r.font.color = params.color || "black";
        } else {
          para.font.color = params.color || "black";
        }
        await ctx.sync();
        return { success: true };
      });

    // ── Paragraph Format ──
    case "setParagraphFormat":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const para = paragraphs.items[params.index];
        if (params.spaceBefore !== undefined) para.spaceBefore = params.spaceBefore;
        if (params.spaceAfter !== undefined) para.spaceAfter = params.spaceAfter;
        if (params.lineSpacing !== undefined) para.lineSpacing = params.lineSpacing;
        if (params.leftIndent !== undefined) para.leftIndent = params.leftIndent;
        if (params.rightIndent !== undefined) para.rightIndent = params.rightIndent;
        if (params.firstLineIndent !== undefined) para.firstLineIndent = params.firstLineIndent;
        if (params.alignment !== undefined) para.alignment = params.alignment;
        await ctx.sync();
        return { success: true };
      });

    // ── Tracked Changes ──
    case "getTrackedChanges":
      return Word.run(async (ctx) => {
        try {
          const body = ctx.document.body;
          const trackedChanges = body.getTrackedChanges();
          trackedChanges.load("items");
          await ctx.sync();
          for (const tc of trackedChanges.items) tc.load("text,type");
          await ctx.sync();
          return { count: trackedChanges.items.length, changes: trackedChanges.items.map((tc: any, i: number) => ({ index: i, text: tc.text, type: tc.type })) };
        } catch { return { error: "Tracked changes API not available in this version of Word", changes: [] }; }
      });

    case "acceptTrackedChange":
      return Word.run(async (ctx) => {
        try {
          const trackedChanges = ctx.document.body.getTrackedChanges();
          trackedChanges.load("items");
          await ctx.sync();
          if (params?.all) {
            trackedChanges.acceptAll();
          } else if (params?.index !== undefined) {
            trackedChanges.items[params.index].accept();
          }
          await ctx.sync();
          return { success: true };
        } catch { return { error: "Tracked changes API not available" }; }
      });

    case "rejectTrackedChange":
      return Word.run(async (ctx) => {
        try {
          const trackedChanges = ctx.document.body.getTrackedChanges();
          trackedChanges.load("items");
          await ctx.sync();
          if (params?.all) {
            trackedChanges.rejectAll();
          } else if (params?.index !== undefined) {
            trackedChanges.items[params.index].reject();
          }
          await ctx.sync();
          return { success: true };
        } catch { return { error: "Tracked changes API not available" }; }
      });

    // ── Style Operations ──
    case "applyStyle":
      return Word.run(async (ctx) => {
        const { fromIndex, toIndex, styleName } = params || {};
        if (fromIndex === undefined || !styleName) throw new Error("fromIndex and styleName required");
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const to = toIndex ?? fromIndex;
        for (let i = fromIndex; i <= to && i < paragraphs.items.length; i++) {
          paragraphs.items[i].style = styleName;
        }
        await ctx.sync();
        return { success: true, applied: to - fromIndex + 1 };
      });

    case "createStyle":
      return Word.run(async (ctx) => {
        const { name, basedOn, fontName, fontSize, bold, italic, color, spaceBefore, spaceAfter, lineSpacing, alignment } = params || {};
        if (!name) throw new Error("params.name required");
        const style = ctx.document.addStyle(name, Word.StyleType.paragraph);
        if (basedOn) style.baseStyle = basedOn;
        if (fontName) style.font.name = fontName;
        if (fontSize) style.font.size = fontSize;
        if (bold !== undefined) style.font.bold = bold;
        if (italic !== undefined) style.font.italic = italic;
        if (color) style.font.color = color;
        if (spaceBefore !== undefined) style.paragraphFormat.spaceBefore = spaceBefore;
        if (spaceAfter !== undefined) style.paragraphFormat.spaceAfter = spaceAfter;
        if (lineSpacing !== undefined) style.paragraphFormat.lineSpacing = lineSpacing;
        if (alignment !== undefined) style.paragraphFormat.alignment = alignment;
        await ctx.sync();
        return { success: true, name };
      });

    case "modifyStyle":
      return Word.run(async (ctx) => {
        const { styleName, fontName, fontSize, bold, italic, color, spaceBefore, spaceAfter, lineSpacing, alignment } = params || {};
        if (!styleName) throw new Error("params.styleName required");
        const style = ctx.document.getStyles().getByNameOrNullObject(styleName);
        style.load("nameLocal");
        await ctx.sync();
        if (style.isNullObject) throw new Error(`Style "${styleName}" not found`);
        if (fontName) style.font.name = fontName;
        if (fontSize) style.font.size = fontSize;
        if (bold !== undefined) style.font.bold = bold;
        if (italic !== undefined) style.font.italic = italic;
        if (color) style.font.color = color;
        if (spaceBefore !== undefined) style.paragraphFormat.spaceBefore = spaceBefore;
        if (spaceAfter !== undefined) style.paragraphFormat.spaceAfter = spaceAfter;
        if (lineSpacing !== undefined) style.paragraphFormat.lineSpacing = lineSpacing;
        if (alignment !== undefined) style.paragraphFormat.alignment = alignment;
        await ctx.sync();
        return { success: true, styleName };
      });

    case "getStyleDetails":
      return Word.run(async (ctx) => {
        const { styleName } = params || {};
        if (!styleName) throw new Error("params.styleName required");
        const style = ctx.document.getStyles().getByNameOrNullObject(styleName);
        style.load("nameLocal,type,builtIn,baseStyle");
        await ctx.sync();
        if (style.isNullObject) throw new Error(`Style "${styleName}" not found`);
        style.font.load("name,size,bold,italic,color,underline");
        style.paragraphFormat.load("spaceBefore,spaceAfter,lineSpacing,alignment,leftIndent,rightIndent,firstLineIndent");
        await ctx.sync();
        return {
          name: style.nameLocal, type: style.type, builtIn: style.builtIn,
          baseStyle: style.baseStyle,
          font: { name: style.font.name, size: style.font.size, bold: style.font.bold, italic: style.font.italic, color: style.font.color, underline: style.font.underline },
          paragraphFormat: {
            spaceBefore: style.paragraphFormat.spaceBefore, spaceAfter: style.paragraphFormat.spaceAfter,
            lineSpacing: style.paragraphFormat.lineSpacing, alignment: style.paragraphFormat.alignment,
            leftIndent: style.paragraphFormat.leftIndent, rightIndent: style.paragraphFormat.rightIndent,
            firstLineIndent: style.paragraphFormat.firstLineIndent,
          },
        };
      });

    // ── Footnotes (expanded) ──
    case "deleteFootnote":
      return Word.run(async (ctx) => {
        const footnotes = ctx.document.body.footnotes;
        footnotes.load("items");
        await ctx.sync();
        const idx = params?.index;
        if (idx === undefined || idx < 0 || idx >= footnotes.items.length) throw new Error("Footnote index out of range");
        footnotes.items[idx].delete();
        await ctx.sync();
        return { success: true };
      });

    case "getFootnoteBody":
      return Word.run(async (ctx) => {
        const footnotes = ctx.document.body.footnotes;
        footnotes.load("items");
        await ctx.sync();
        const idx = params?.index;
        if (idx === undefined || idx < 0 || idx >= footnotes.items.length) throw new Error("Footnote index out of range");
        const fn = footnotes.items[idx];
        fn.body.load("text");
        fn.body.paragraphs.load("items");
        await ctx.sync();
        for (const p of fn.body.paragraphs.items) p.load("text,style");
        await ctx.sync();
        return {
          index: idx,
          text: fn.body.text,
          paragraphs: fn.body.paragraphs.items.map((p: any, i: number) => ({ index: i, text: p.text, style: p.style })),
        };
      });

    case "insertFootnoteWithFormat":
      return Word.run(async (ctx) => {
        const { anchorText, footnoteText, matchCase } = params || {};
        if (!anchorText || !footnoteText) throw new Error("anchorText and footnoteText required");
        const results = ctx.document.body.search(anchorText, { matchCase: matchCase ?? true });
        results.load("items");
        await ctx.sync();
        if (results.items.length === 0) throw new Error(`Anchor text "${anchorText}" not found`);
        const range = results.items[0].getRange(Word.RangeLocation.end);
        const fn = range.insertFootnote(footnoteText);
        fn.body.load("text");
        await ctx.sync();
        return { body: fn.body.text };
      });

    case "reorderFootnotes":
      return Word.run(async (ctx) => {
        const footnotes = ctx.document.body.footnotes;
        footnotes.load("items");
        await ctx.sync();
        const result: any[] = [];
        for (let i = 0; i < footnotes.items.length; i++) {
          const fn = footnotes.items[i];
          fn.body.load("text");
          fn.reference.load("text");
        }
        await ctx.sync();
        for (let i = 0; i < footnotes.items.length; i++) {
          const fn = footnotes.items[i];
          result.push({ index: i, body: fn.body.text, referenceText: fn.reference.text });
        }
        return { footnotes: result, count: result.length };
      });

    // ── Citations / Table of Authorities ──
    case "markCitation":
      return Word.run(async (ctx) => {
        const { shortCite, longCite, category, searchText } = params || {};
        if (!shortCite || !longCite) throw new Error("shortCite and longCite required");
        // Category: 1=Cases, 2=Statutes, 3=Other Authorities, 4=Rules
        const cat = category || 1;
        // Find the text to mark
        const anchor = searchText || shortCite;
        const results = ctx.document.body.search(anchor, { matchCase: true });
        results.load("items");
        await ctx.sync();
        if (results.items.length === 0) throw new Error(`Text "${anchor}" not found in document`);
        // Insert TA field code: { TA \l "longCite" \s "shortCite" \c category }
        const fieldCode = `TA \\l "${longCite}" \\s "${shortCite}" \\c ${cat}`;
        const range = results.items[0].getRange(Word.RangeLocation.end);
        range.insertText(" ", Word.InsertLocation.after);
        const fieldRange = range.getRange(Word.RangeLocation.after);
        try {
          fieldRange.insertField(Word.InsertLocation.end, Word.FieldType.empty, fieldCode, true);
          await ctx.sync();
          return { success: true, shortCite, longCite, category: cat };
        } catch {
          // Fallback: insert as hidden text field code marker
          const marker = `{${fieldCode}}`;
          range.insertText(marker, Word.InsertLocation.after);
          await ctx.sync();
          return { success: true, shortCite, longCite, category: cat, note: "Inserted as text marker (insertField API not available)" };
        }
      });

    case "insertTableOfAuthorities":
      return Word.run(async (ctx) => {
        const { category, paragraphIndex } = params || {};
        // Insert a TOA field: { TOA \c category }
        const cat = category || 0; // 0 = all categories
        const fieldCode = cat > 0 ? `TOA \\c ${cat}` : `TOA`;
        let insertRange: Word.Range;
        if (paragraphIndex !== undefined) {
          const paragraphs = ctx.document.body.paragraphs;
          paragraphs.load("items");
          await ctx.sync();
          insertRange = paragraphs.items[paragraphIndex].getRange(Word.RangeLocation.after);
        } else {
          insertRange = ctx.document.body.getRange(Word.RangeLocation.end);
        }
        try {
          insertRange.insertField(Word.InsertLocation.end, Word.FieldType.empty, fieldCode, true);
          await ctx.sync();
          return { success: true, category: cat };
        } catch {
          insertRange.insertText(`[Table of Authorities${cat > 0 ? ` — Category ${cat}` : ""}]\n{${fieldCode}}`, Word.InsertLocation.end);
          await ctx.sync();
          return { success: true, category: cat, note: "Inserted as text placeholder (insertField API not available)" };
        }
      });

    // ── Cross-References ──
    case "insertCrossReference":
      return Word.run(async (ctx) => {
        const { type, target, text, paragraphIndex } = params || {};
        if (!type || !target) throw new Error("type and target required");
        // type: "heading", "footnote", "bookmark"
        let refText = text || "";
        if (type === "heading") {
          // Find the heading paragraph and generate reference text
          const paragraphs = ctx.document.body.paragraphs;
          paragraphs.load("items");
          await ctx.sync();
          for (const p of paragraphs.items) p.load("text,style");
          await ctx.sync();
          const heading = paragraphs.items.find((p: any) => p.text.includes(target) && p.style.toLowerCase().includes("heading"));
          if (!heading) throw new Error(`Heading containing "${target}" not found`);
          if (!refText) refText = heading.text.trim();
        } else if (type === "footnote") {
          if (!refText) refText = `footnote ${target}`;
        } else if (type === "bookmark") {
          if (!refText) refText = target;
        }
        // Insert a cross-reference field
        const fieldCode = type === "heading" ? `REF "${target}" \\h`
          : type === "bookmark" ? `REF ${target} \\h`
          : `NOTEREF ${target} \\h`;
        let insertRange: Word.Range;
        if (paragraphIndex !== undefined) {
          const paragraphs = ctx.document.body.paragraphs;
          paragraphs.load("items");
          await ctx.sync();
          insertRange = paragraphs.items[paragraphIndex].getRange(Word.RangeLocation.end);
        } else {
          const sel = ctx.document.getSelection();
          insertRange = sel.getRange(Word.RangeLocation.end);
        }
        try {
          insertRange.insertField(Word.InsertLocation.end, Word.FieldType.empty, fieldCode, true);
          await ctx.sync();
          return { success: true, type, target, fieldCode };
        } catch {
          // Fallback: insert as styled text
          insertRange.insertText(refText, Word.InsertLocation.end);
          await ctx.sync();
          return { success: true, type, target, note: "Inserted as plain text (insertField API not available)" };
        }
      });

    case "validateCrossReferences":
      return Word.run(async (ctx) => {
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        // Collect all heading texts
        const headings: { index: number; text: string; level: number; style: string }[] = [];
        const allText: { index: number; text: string }[] = [];
        const batchSize = 200;
        for (let i = 0; i < paragraphs.items.length; i += batchSize) {
          const batch = paragraphs.items.slice(i, i + batchSize);
          for (const p of batch) p.load("text,style");
          await ctx.sync();
          for (let j = 0; j < batch.length; j++) {
            const p = batch[j];
            allText.push({ index: i + j, text: p.text });
            const sLow = p.style.toLowerCase();
            if (sLow.startsWith("heading") || sLow.includes("heading")) {
              const m = p.style.match(/\d+/);
              headings.push({ index: i + j, text: p.text.trim(), level: m ? parseInt(m[0], 10) : 1, style: p.style });
            }
          }
        }
        // Patterns to validate
        const patterns = [
          /Section\s+(\d+[\.\d]*)/gi,
          /Article\s+([IVXLCDM]+|\d+)/gi,
          /see\s+supra\s+(?:Section\s+)?(\S+)/gi,
          /see\s+infra\s+(?:Section\s+)?(\S+)/gi,
          /¶\s*(\d+)/gi,
          /Part\s+([IVXLCDM]+|\d+)/gi,
        ];
        const issues: { paragraphIndex: number; text: string; reference: string; pattern: string; found: boolean }[] = [];
        for (const para of allText) {
          for (const pat of patterns) {
            pat.lastIndex = 0;
            let match;
            while ((match = pat.exec(para.text)) !== null) {
              const ref = match[0];
              const refTarget = match[1];
              // Check if any heading contains this reference
              const found = headings.some(h => h.text.includes(refTarget) || h.text.includes(ref));
              if (!found) {
                issues.push({
                  paragraphIndex: para.index,
                  text: para.text.substring(Math.max(0, match.index - 20), match.index + ref.length + 20),
                  reference: ref,
                  pattern: pat.source,
                  found: false,
                });
              }
            }
          }
        }
        return { headingCount: headings.length, issueCount: issues.length, issues, headings: headings.map(h => ({ index: h.index, text: h.text, level: h.level })) };
      });
    case "exportPdf": {
      return new Promise((resolve, _reject) => {
        Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 65536 }, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            return resolve({ error: result.error.message });
          }
          const file = result.value;
          const sliceCount = file.sliceCount;
          const slices: Uint8Array[] = [];
          let totalBytes = 0;

          const getSlice = (index: number) => {
            file.getSliceAsync(index, (sliceResult) => {
              if (sliceResult.status === Office.AsyncResultStatus.Failed) {
                file.closeAsync();
                return resolve({ error: sliceResult.error.message });
              }

              const bytes = new Uint8Array(sliceResult.value.data);
              slices.push(bytes);
              totalBytes += bytes.length;

              if (index + 1 === sliceCount) {
                // Concatenate bytes in-order, then base64 encode once.
                const combined = new Uint8Array(totalBytes);
                let offset = 0;
                for (const part of slices) {
                  combined.set(part, offset);
                  offset += part.length;
                }

                // Convert to base64 in chunks to avoid call stack / arg limits.
                const CHUNK = 0x8000;
                let binary = "";
                for (let i = 0; i < combined.length; i += CHUNK) {
                  const sub = combined.subarray(i, i + CHUNK);
                  binary += String.fromCharCode(...Array.from(sub));
                }

                file.closeAsync();
                resolve({ pdf: btoa(binary), slices: sliceCount, bytes: totalBytes });
              } else {
                getSlice(index + 1);
              }
            });
          };

          getSlice(0);
        });
      });
    }

    case "getToaEntries": {
      return Word.run(async (ctx) => {
        const body = ctx.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await ctx.sync();
        const entries: Array<{text: string, pages: string}> = [];
        let inToa = false;
        for (const para of paragraphs.items) {
          para.load("text,style");
        }
        await ctx.sync();
        for (const para of paragraphs.items) {
          const text = para.text.trim();
          if (text.toUpperCase().includes("TABLE OF AUTHORITIES")) {
            inToa = true;
            continue;
          }
          if (inToa && (text.toUpperCase().includes("TABLE OF CONTENTS") ||
              para.style?.startsWith("Heading 1") ||
              text.toUpperCase() === "INTRODUCTION" ||
              text.toUpperCase() === "PRELIMINARY STATEMENT" ||
              text.toUpperCase() === "ARGUMENT")) {
            break;
          }
          if (inToa && text.length > 0) {
            const dotMatch = text.match(/^(.+?)\s*[.\u2026\u00b7]{2,}\s*(.+)$/);
            if (dotMatch) {
              entries.push({ text: dotMatch[1].trim(), pages: dotMatch[2].trim() });
            } else {
              const pageMatch = text.match(/^(.+?)\s+((?:\d+(?:,\s*)?)+|passim)\s*$/);
              if (pageMatch) {
                entries.push({ text: pageMatch[1].trim(), pages: pageMatch[2].trim() });
              }
            }
          }
        }
        return { entries, count: entries.length };
      });
    }

    case "getPageSetup": {
      return Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();
        const section = sections.items[data.sectionIndex || 0];
        section.load("headerDistance,footerDistance");
        const body = section.body;
        body.load("style");
        await ctx.sync();
        try {
          (section as any).load("pageSetup");
          await ctx.sync();
          const ps = (section as any).pageSetup;
          return {
            topMargin: ps?.topMargin,
            bottomMargin: ps?.bottomMargin,
            leftMargin: ps?.leftMargin,
            rightMargin: ps?.rightMargin,
            gutter: ps?.gutter,
            paperSize: ps?.paperSize,
            headerDistance: section.headerDistance,
            footerDistance: section.footerDistance
          };
        } catch {
          return {
            headerDistance: section.headerDistance,
            footerDistance: section.footerDistance,
            note: "Full page setup (margins) requires WordApi 1.5+. Available properties returned."
          };
        }
      });
    }

    case "setPageSetup": {
      return Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();
        const section = sections.items[data.sectionIndex || 0];
        if (data.headerDistance !== undefined) section.headerDistance = data.headerDistance;
        if (data.footerDistance !== undefined) section.footerDistance = data.footerDistance;
        try {
          const ps = (section as any).pageSetup;
          if (data.topMargin !== undefined) ps.topMargin = data.topMargin;
          if (data.bottomMargin !== undefined) ps.bottomMargin = data.bottomMargin;
          if (data.leftMargin !== undefined) ps.leftMargin = data.leftMargin;
          if (data.rightMargin !== undefined) ps.rightMargin = data.rightMargin;
          if (data.gutter !== undefined) ps.gutter = data.gutter;
          if (data.orientation !== undefined) ps.orientation = data.orientation;
          if (data.paperSize !== undefined) ps.paperSize = data.paperSize;
          await ctx.sync();
          return { success: true };
        } catch {
          await ctx.sync();
          return { success: true, note: "Only headerDistance/footerDistance set. Full margins require WordApi 1.5+" };
        }
      });
    }

    case "getPageNumbers": {
      return Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();
        for (const s of sections.items) {
          s.load("headerDistance,footerDistance");
          s.body.load("text");
        }
        await ctx.sync();
        return {
          sectionCount: sections.items.length,
          sections: sections.items.map((s, i) => ({
            index: i,
            headerDistance: s.headerDistance,
            footerDistance: s.footerDistance,
            bodyLength: s.body.text.length
          }))
        };
      });
    }

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

    // Populate workspace + precedent
    const workspaceInput = document.getElementById("settings-workspace-path") as HTMLInputElement;
    const precedentInput = document.getElementById("settings-precedent-path") as HTMLInputElement;
    if (workspaceInput) workspaceInput.value = data.workspacePath || "";
    if (precedentInput) precedentInput.value = data.precedentPath || "";
    updateTrayIcons(data.workspacePath, data.precedentPath);

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
function normalizeSmartQuotes(input: string): string {
  let out = input;
  // First convert clearly paired straight double quotes
  out = out.replace(/"([^"]+)"/g, "“$1”");
  // Convert apostrophes in words (can't -> can’t)
  out = out.replace(/(\w)'(\w)/g, "$1’$2");
  // Opening single quotes after whitespace/punctuation
  out = out.replace(/(^|[\s([{<])'/g, "$1‘");
  // Remaining single quotes become closing
  out = out.replace(/'/g, "’");
  // Opening double quotes after whitespace/punctuation
  out = out.replace(/(^|[\s([{<])"/g, "$1“");
  // Remaining straight double quotes become closing
  out = out.replace(/"/g, "”");
  return out;
}

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

  // OpenAI — fetch available models dynamically
  if (settings.openaiApiKey) {
    try {
      const r = await fetch("http://localhost:3001/api/models/openai");
      const j = await r.json();
      if (j?.ok && j.data?.length) {
        for (const m of j.data) {
          models.push({ id: `openai:${m.id}`, label: `${m.id} (OpenAI)`, backend: "openai" });
        }
      }
    } catch {
      // Fallback to common models
      models.push(
        { id: "openai:gpt-4o", label: "GPT-4o (OpenAI)", backend: "openai" },
        { id: "openai:gpt-4o-mini", label: "GPT-4o Mini (OpenAI)", backend: "openai" },
      );
    }
  }

  // Anthropic — fetch available models dynamically
  if (settings.anthropicApiKey) {
    try {
      const r = await fetch("http://localhost:3001/api/models/anthropic");
      const j = await r.json();
      if (j?.ok && j.data?.length) {
        for (const m of j.data) {
          models.push({ id: `anthropic:${m.id}`, label: `${m.name || m.id} (Anthropic)`, backend: "anthropic" });
        }
      }
    } catch {
      // Fallback
      models.push(
        { id: "anthropic:claude-sonnet-4-20250514", label: "Claude Sonnet 4 (Anthropic)", backend: "anthropic" },
      );
    }
  }

  // Local endpoints — fetch models dynamically
  const endpoints = settings.localEndpoints || [];
  for (const ep of endpoints) {
    if (ep.name && ep.baseUrl) {
      try {
        const r = await fetch(`http://localhost:3001/api/models/local?baseUrl=${encodeURIComponent(ep.baseUrl)}`);
        const j = await r.json();
        if (j?.ok && j.data?.length) {
          for (const m of j.data) {
            models.push({ id: `local:${ep.baseUrl}:${m.id}`, label: `${m.id} (${ep.name})`, backend: "local" });
          }
        } else {
          models.push({ id: `local:${ep.baseUrl}:default`, label: `${ep.name} (Local)`, backend: "local" });
        }
      } catch {
        models.push({ id: `local:${ep.baseUrl}:default`, label: `${ep.name} (Local)`, backend: "local" });
      }
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

  const workspacePath = (document.getElementById("settings-workspace-path") as HTMLInputElement)?.value?.trim();
  const precedentPath = (document.getElementById("settings-precedent-path") as HTMLInputElement)?.value?.trim();

  const body: Record<string, any> = {};
  body.workspacePath = workspacePath || "";
  body.precedentPath = precedentPath || "";
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
    if (j?.ok) updateTrayIcons(workspacePath, precedentPath);
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

function setModeToggleVisual(modeToggle: HTMLElement, enabled: boolean): void {
  modeToggle.textContent = enabled ? "🔍 Track" : "⚡ YOLO";
  modeToggle.classList.toggle("tracking", enabled);
}

async function applyTrackChangesMode(desired: boolean): Promise<boolean> {
  try {
    return await Word.run(async (context) => {
      context.document.changeTrackingMode = desired
        ? Word.ChangeTrackingMode.trackAll
        : Word.ChangeTrackingMode.off;
      context.document.load("changeTrackingMode");
      await context.sync();
      const actual = context.document.changeTrackingMode;
      if (desired) return actual === Word.ChangeTrackingMode.trackAll;
      return actual === Word.ChangeTrackingMode.off;
    });
  } catch (e) {
    console.warn("changeTrackingMode not available:", e);
    return false;
  }
}

function setupTrackChangesToggle(): void {
  const modeToggle = document.getElementById("mode-toggle");
  if (!modeToggle) return;

  modeToggle.addEventListener("click", async () => {
    const desired = !trackChangesMode;
    const confirmed = await applyTrackChangesMode(desired);

    trackChangesMode = confirmed ? desired : false;
    setModeToggleVisual(modeToggle, trackChangesMode);

    if (!confirmed && desired) {
      appendChatEntry("assistant", "Track Changes could not be enabled via Word API on this setup. Please turn on Track Changes in Word manually, then try again.");
      localStorage.removeItem("sidebar-track-changes");
    }

    // Tell the server the confirmed state
    fetch("http://localhost:3001/api/settings/mode", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ trackChanges: trackChangesMode }),
    }).catch(() => {});

    // Persist preference
    if (trackChangesMode) localStorage.setItem("sidebar-track-changes", "true");
    else localStorage.removeItem("sidebar-track-changes");
  });

  
  // Tighten button — tightens selected text
  const tightenBtn = document.getElementById("tighten-btn");
  tightenBtn?.addEventListener("click", async () => {
    const prompt = `Tighten the selected text now. Do NOT ask follow-up questions.

Rules:
- Preserve meaning and tone.
- Cut filler, redundancy, and throat-clearing.
- Prefer active voice and tighter phrasing.
- Keep formatting intact.
- Use granular edits (phrase/sentence level), not full-selection replacement.
- In Track mode, avoid broad paragraph rewrites.

After making edits, respond with EXACTLY:
**Tightened:** [one-sentence summary of what changed]
**Saved:** [X] words ([Y]% reduction)

Count words in original vs final to compute savings.`;
    
    const input = document.getElementById("prompt-input") as HTMLTextAreaElement;
    if (input) {
      input.value = prompt;
      (input as any)._displayLabel = "✂️ Tighten selection";
      const sendBtn = document.getElementById("prompt-send") as HTMLButtonElement;
      sendBtn?.click();
    }
  });

  // Restore on load (verify Word accepted it before showing Track mode)
  const savedMode = localStorage.getItem("sidebar-track-changes");
  if (savedMode === "true") {
    applyTrackChangesMode(true).then((confirmed) => {
      trackChangesMode = confirmed;
      setModeToggleVisual(modeToggle, trackChangesMode);
      if (!confirmed) {
        localStorage.removeItem("sidebar-track-changes");
        appendChatEntry("assistant", "Track Changes preference was saved, but Word did not accept Track mode in this session. Using YOLO until Track Changes is enabled in Word.");
      }
      fetch("http://localhost:3001/api/settings/mode", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ trackChanges: trackChangesMode }),
      }).catch(() => {});
    });
  } else {
    setModeToggleVisual(modeToggle, false);
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

const QUICK_ACTIONS_DEFAULTS: QuickActionCategory[] = [];

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


async function runDirectAction(action: string, displayLabel: string): Promise<void> {
  const history = document.getElementById("prompt-history");
  if (!history) return;

  // Show user message
  appendChatEntry("user", displayLabel);

  if (action === "checkToaPages") {
    // Phase 1: Export PDF and get page map
    addThinkingIndicator();
    const thinkingEl = document.getElementById("thinking-indicator");
    const elapsed = thinkingEl?.querySelector(".thinking-elapsed") as HTMLElement | null;

    try {
      // Update status
      if (elapsed) elapsed.textContent = "Exporting PDF...";

      const res = await fetch("http://localhost:3001/api/toa/check", { method: "POST" });
      const raw = await res.json();
      const payload = raw?.data ?? raw;

      if (!raw?.ok || !payload?.pageMap) {
        removeThinkingIndicator();
        appendChatEntry("assistant", raw?.error || "Failed to export PDF and extract pages.");
        return;
      }

      if (elapsed) elapsed.textContent = "Sending to AI for analysis...";

      // Phase 2: Build prompt with REAL page data and send to LLM
      const toaList = payload.toaEntries.map((e: any) => `  - ${e.text} → listed pages: ${e.pages}`).join("\n");

      const pageMapText = payload.pageMap.map((p: any) =>
        `--- PAGE ${p.page} ---\n${p.text}`
      ).join("\n\n");

      const prompt = `You are auditing a Table of Authorities (TOA) for a legal brief.

Below are the TOA entries with their LISTED page numbers, followed by the ACTUAL document text organized by REAL page number (extracted from a PDF export of the document).

Your task: For each TOA entry, find where that authority is actually cited in the document body (using the page-mapped text below), and compare the REAL page numbers against the LISTED page numbers in the TOA.

Handle citation variants: short cites, supra/infra references, party-name-only references, different reporter formats.

IMPORTANT: The page numbers in the page map below are REAL document page numbers from the PDF. Use ONLY these to determine actual pages. Do NOT guess or infer page numbers from section headings or TOC entries.

## TOA Entries (with listed pages):
${toaList}

## Document Text by Page:
${pageMapText}

## Output Format (use bullet lists, NOT tables):

**1) High-Confidence Errors**
For each:
**Authority name**
- TOA lists: X, Y → Actually on: X, Z
- Confidence: High
- Rationale: one sentence

**2) Possible Mismatches (Medium Confidence)**
Same format

**3) Entries That Appear Correct**
Brief list

**4) Could Not Verify**
Entries where the citation wasn't found in the body text

Keep it compact. Be precise about page numbers.

After listing all findings, IMMEDIATELY fix any incorrect page numbers in the TOA by editing the document directly. Use the replaceParagraph or findReplace tools to update the wrong page numbers in the Table of Authorities section. Do NOT tell the user to regenerate fields, update fields manually, or press F9. Make the corrections yourself.`;

      // Send to LLM via the normal prompt mechanism
      const input = document.getElementById("prompt-input") as HTMLTextAreaElement;
      if (input) {
        removeThinkingIndicator();
        input.value = prompt;
        (input as any)._displayLabel = "Analyzing TOA with real page data...";
        const sendBtn = document.getElementById("prompt-send") as HTMLButtonElement;
        sendBtn?.click();
      }
    } catch (e: any) {
      removeThinkingIndicator();
      appendChatEntry("assistant", "Error: " + (e.message || "TOA check failed"));
    }
    return;
  }

  // Generic direct action fallback
  appendChatEntry("assistant", "Unknown direct action: " + action);
}


async function executeQuickAction(prompt: string, displayLabel?: string): Promise<void> {
  // Handle direct server actions (no LLM needed)
  if (prompt.startsWith("__DIRECT_ACTION__")) {
    const action = prompt.substring("__DIRECT_ACTION__".length);
    await runDirectAction(action, displayLabel || action);
    return;
  }
  const resolved = await resolvePromptVariables(prompt);
  const input = document.getElementById("prompt-input") as HTMLTextAreaElement;
  if (input) {
    input.value = resolved;
    // Store display label so the chat shows a clean message instead of the raw prompt
    if (displayLabel) {
      (input as any)._displayLabel = displayLabel;
    }
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

function updateTrayIcons(workspace?: string, precedent?: string): void {
  const wsTray = document.getElementById("workspace-tray");
  const prTray = document.getElementById("precedent-tray");
  if (wsTray) {
    const hasWs = !!workspace;
    wsTray.style.opacity = hasWs ? "0.8" : "0.4";
    wsTray.title = hasWs ? `Workspace: ${workspace}` : "Workspace: not set (click to set)";
  }
  if (prTray) {
    const hasPr = !!precedent;
    prTray.style.opacity = hasPr ? "0.8" : "0.4";
    prTray.title = hasPr ? `Precedent: ${precedent}` : "Precedent: not set (click to set)";
  }
}

async function pickAndSetFolder(field: "workspacePath" | "precedentPath", title: string): Promise<void> {
  try {
    const r = await fetch("http://localhost:3001/api/pick-folder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title }),
    });
    const j = await r.json();
    if (!j?.ok || !j.path) return; // cancelled or error

    // Save to settings
    const settingsR = await fetch("http://localhost:3001/api/settings");
    const settingsJ = await settingsR.json();
    const current = settingsJ?.data || {};

    const body: Record<string, any> = {};
    body[field] = j.path;
    // Preserve other settings
    if (current.workspacePath && field !== "workspacePath") body.workspacePath = current.workspacePath;
    if (current.precedentPath && field !== "precedentPath") body.precedentPath = current.precedentPath;

    await fetch("http://localhost:3001/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    // Update tray + settings input
    const ws = field === "workspacePath" ? j.path : (current.workspacePath || "");
    const pr = field === "precedentPath" ? j.path : (current.precedentPath || "");
    updateTrayIcons(ws, pr);

    // Also update settings input if panel is open
    const inputId = field === "workspacePath" ? "settings-workspace-path" : "settings-precedent-path";
    const input = document.getElementById(inputId) as HTMLInputElement;
    if (input) input.value = j.path;
  } catch (e) {
    console.error("Folder pick failed:", e);
  }
}

function setupTrayIcons(): void {
  const wsTray = document.getElementById("workspace-tray");
  const prTray = document.getElementById("precedent-tray");
  wsTray?.addEventListener("click", () => pickAndSetFolder("workspacePath", "Select Workspace Folder"));
  prTray?.addEventListener("click", () => pickAndSetFolder("precedentPath", "Select Precedent Folder"));
  // Memory tray — click to view/manage memories
  const memTray = document.getElementById("memory-tray");
  memTray?.addEventListener("click", showMemoryPanel);
  
  // Load initial state
  fetch("http://localhost:3001/api/settings").then(r => r.json()).then(j => {
    if (j?.ok) updateTrayIcons(j.data?.workspacePath, j.data?.precedentPath);
  }).catch(() => {});
  refreshMemoryCount();
}

async function refreshMemoryCount(): Promise<void> {
  try {
    const r = await fetch("http://localhost:3001/api/memory");
    const j = await r.json();
    const count = j?.entries?.length || 0;
    const tray = document.getElementById("memory-tray");
    if (tray) {
      tray.style.opacity = count > 0 ? "0.8" : "0.4";
      tray.title = `Memory: ${count} item${count !== 1 ? "s" : ""} learned`;
    }
  } catch {}
}

async function showMemoryPanel(): Promise<void> {
  try {
    const r = await fetch("http://localhost:3001/api/memory");
    const j = await r.json();
    const entries = j?.entries || [];
    
    // Build a simple overlay
    let existing = document.getElementById("memory-overlay");
    if (existing) existing.remove();
    
    const overlay = document.createElement("div");
    overlay.id = "memory-overlay";
    overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,0.4);z-index:9999;display:flex;align-items:center;justify-content:center;";
    overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });
    
    const panel = document.createElement("div");
    panel.style.cssText = "background:var(--bg-primary,#fff);border-radius:8px;padding:16px;max-width:90%;max-height:80%;overflow-y:auto;min-width:280px;box-shadow:0 8px 32px rgba(0,0,0,0.2);";
    
    let html = "<h3 style=\"margin:0 0 12px;font-size:14px;\">🧠 Learned Memory</h3>";
    if (entries.length === 0) {
      html += "<p style=\"color:var(--text-secondary,#888);font-size:12px;\">No memories yet. As you work, I\'ll learn your preferences and corrections.</p>";
    } else {
      for (const e of entries) {
        const icon = e.category === "correction" ? "🔴" : e.category === "preference" ? "🟢" : e.category === "style" ? "🟡" : "🔵";
        const scope = e.scope === "global" ? "all docs" : "this doc";
        html += `<div style="display:flex;align-items:start;gap:6px;margin-bottom:8px;font-size:11px;line-height:1.4;">
          <span>${icon}</span>
          <span style="flex:1;">${e.text} <span style="opacity:0.5;">(${scope})</span></span>
          <button data-mem-id="${e.id}" style="background:none;border:none;cursor:pointer;opacity:0.4;font-size:10px;" title="Delete">✕</button>
        </div>`;
      }
    }
    html += "<div style=\"margin-top:12px;text-align:right;\"><button id=\"memory-close\" style=\"font-size:11px;padding:4px 12px;border:1px solid var(--border,#ddd);border-radius:4px;background:none;cursor:pointer;\">Close</button></div>";
    
    panel.innerHTML = html;
    overlay.appendChild(panel);
    document.body.appendChild(overlay);
    
    document.getElementById("memory-close")?.addEventListener("click", () => overlay.remove());
    
    // Delete handlers
    panel.querySelectorAll("[data-mem-id]").forEach(btn => {
      btn.addEventListener("click", async (e) => {
        const id = (e.currentTarget as HTMLElement).getAttribute("data-mem-id");
        if (!id) return;
        await fetch(`http://localhost:3001/api/memory/${id}`, { method: "DELETE" });
        overlay.remove();
        refreshMemoryCount();
        showMemoryPanel(); // Refresh
      });
    });
  } catch (e) {
    console.error("Memory panel error:", e);
  }
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
          executeQuickAction(effectivePrompt, action.name);
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
          // Position dropdown using fixed positioning
          const rect = pill.getBoundingClientRect();
          dropdown.style.left = rect.left + "px";
          // Try to show above the pill; if not enough room, show below
          dropdown.classList.add("visible");
          const ddRect = dropdown.getBoundingClientRect();
          if (rect.top - ddRect.height > 0) {
            dropdown.style.top = (rect.top - ddRect.height - 4) + "px";
          } else {
            dropdown.style.top = (rect.bottom + 4) + "px";
          }
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
        const rect = customPill.getBoundingClientRect();
        customDropdown.style.left = rect.left + "px";
        customDropdown.classList.add("visible");
        const ddRect = customDropdown.getBoundingClientRect();
        if (rect.top - ddRect.height > 0) {
          customDropdown.style.top = (rect.top - ddRect.height - 4) + "px";
        } else {
          customDropdown.style.top = (rect.bottom + 4) + "px";
        }
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
