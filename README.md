# The Sidebar

**Connect your AI to Microsoft Word.** Open source, model-agnostic, runs entirely on your machine.

Use [OpenClaw](https://openclaw.ai), OpenAI, Anthropic, Ollama, LM Studio — or any OpenAI-compatible API — as a fully autonomous document editor inside Word. 60+ tools including live web research. Your keys. Your data. Nothing leaves your machine.

Built by a litigation partner who uses it daily.

<!-- TODO: screenshot -->
![Screenshot](docs/screenshot.png)

---

### 🦞 OpenClaw + Word

**The Sidebar is the bridge between [OpenClaw](https://github.com/openclaw/openclaw) and Microsoft Word.** If you run OpenClaw, The Sidebar gives your agent full document control — read, edit, footnote, format, cite-check, and rewrite — without copy-pasting between chat and your doc.

Point The Sidebar at your OpenClaw gateway (`http://localhost:18789`) and your agent gets 55+ document tools, encrypted per-document sessions, reference folder RAG, and a real-time task pane UI inside Word.

**Async architecture** — OpenClaw requests run in the background. You get an immediate response while the agent works (researching, browsing, editing), with text streaming progressively into the task pane as results arrive. No more waiting for long agentic loops to complete.

**Two-track editing** — The agent automatically picks the right approach:
- **Sidebar API** for surgical edits (fix a typo, add a footnote, find/replace)
- **python-docx + reload** for bulk operations (rewrite a section, restructure the document) — edits the .docx directly, then tells Word to refresh via `POST /api/document/reload`

**Web research built in** — The Sidebar can search the web and fetch URLs without leaving Word. Ask it to find market data, verify a quote, look up a case summary, or pull manufacturer representations — and it inserts the results with source footnotes automatically.

**Not using OpenClaw?** No problem. The Sidebar works standalone with any LLM provider.

---

## Why This Exists

Legal work lives in Word. Every AI tool either (a) makes you copy-paste between a chat window and your document, or (b) is Microsoft Copilot — expensive, inflexible, and surprisingly bad at actual document manipulation.

The Sidebar sits next to your document in Word's task pane and does what Copilot can't: find/replace with regex, insert footnotes, apply styles, add comments, format tables, check citations, tighten prose — autonomously, in sequence, without you babysitting each step.

Microsoft Copilot costs $30/month, locks you into one model, and can't even add a footnote. The Sidebar gives your AI 55+ document tools and lets you pick any model — including one running on your own hardware.

## Features

### 🤖 AI & Models
- **Model-agnostic** — OpenAI, Anthropic, OpenClaw, Ollama, LM Studio, or any OpenAI-compatible API
- **OpenClaw integration** — connect your OpenClaw agent directly to Word with async streaming and full document tool access
- **BYOK (Bring Your Own Keys)** — your credentials, your choice of provider, your data
- **Agentic loop** — AI plans and executes multi-step document edits autonomously (proper structured tool messages for OpenAI and Anthropic tool-calling formats)
- **Token-by-token streaming** — see responses as they generate, with live tool execution ticker
- **Smart context management** — dynamic context window detection per model, configurable budget (default 40%)
- **Prompt caching** — Anthropic `cache_control` + OpenAI-optimized message ordering for up to 90% input cost reduction

### 📝 Document Tools (55+)
- Read, edit, insert, delete paragraphs
- Find and replace (literal + regex)
- **Footnotes** — add, edit, delete, read body, reorder
- **Comments** — add and list
- **Styles** — apply, create, modify, inspect
- **Tables** — create, read, edit cells, add rows/columns
- **Headers/footers** — read and edit
- **Page setup** — margins, gutters, orientation, paper size
- **Track changes** — accept/reject individual or all
- **Citations** — mark TA fields, insert Table of Authorities
- **Cross-references** — insert and validate
- **Page/section breaks**, lists, bookmarks, highlighting, font colors, paragraph formatting
- **Web research** — `webSearch` and `webFetch` tools let the AI search the web and read URLs mid-document, inserting sourced content with footnotes

### ⚡ Quick Actions
- **One-click legal tools** — Cite Check, Long/Short Cites, TOA Pages, Defined Terms, House Style, Risk Analysis, and more
- **✂️ Tighten** — select text, click once, AI rewrites it tighter without losing substance. Preserves footnotes and formatting. In Track mode, applies targeted phrase-level edits for clean redlines; in YOLO mode, does a full paragraph replace
- **Editable prompts** — customize any built-in action, save your own custom prompts with `{{selection}}` and `{{document}}` variables
- **TOA Page Checker** — exports to PDF for real pagination, uses LLM intelligence for citation variant matching

### 📁 Reference Documents (RAG)
- **Designate reference folders** — point The Sidebar at your case folders and every document becomes searchable context
- **Automatic indexing** — recursively scans .docx, .pdf, .txt, .md files
- **Smart retrieval** — embeds and retrieves only the most relevant chunks
- **Multiple embedding backends** — OpenAI, local endpoints, or built-in TF-IDF fallback (works offline, zero API keys)
- **OpenClaw-aware** — passes folder paths as filesystem hints instead of RAG

### 💬 Conversation
- **Per-document sessions** — each document gets its own conversation thread
- **Session restoration** — auto-generated recap on resume, full history search
- **Revert system** — per-exchange undo, one click to roll back
- **Inline mini-diffs** — every edit shows a compact before/after summary
- **Stop button** — cancel in-flight requests instantly, with real abort signal propagation that kills the underlying HTTP request
- **Async OpenClaw mode** — long-running agent tasks stream results progressively instead of blocking, with no timeout (runs until done or stopped)
- **Elapsed timer** — see how long the model has been thinking, with ⌛ progress counter for OpenClaw
- **Completion indicator** — red ■ Done badge at both the top and bottom of long tool chains, auto-scrolls into view when the agent finishes

### ✏️ Editing Experience
- **Track Changes / YOLO mode** — toggle between tracked changes and direct edits, synced with Word's actual state
- **Live streaming** — typewriter cursor, tool execution ticker with spinners
- **Smooth auto-scroll** — `requestAnimationFrame` throttled
- **Pre-warm on connect** — document indexed before first prompt
- **Batch operations** — consecutive edits batched for speed

### 🖥️ App
- **macOS menu bar app** — Electron tray app, starts on login
- **Dark & light mode** — auto-detects system preference, manual toggle
- **Localhost only** — HTTP server on port 3001, nothing exposed to the network

## 🔒 Security & Privacy

This is a tool for lawyers. It's built like one.

- **Zero cloud dependency** — server runs on `localhost:3001`. No data leaves your machine except outbound LLM API calls to the provider *you* choose
- **Encrypted session storage** — AES-256-GCM with a machine-specific key. Moving a document to another computer does not expose prior conversations
- **Attorney work product protection** — session data is encrypted and machine-bound
- **No telemetry** — zero tracking, zero analytics, zero phoning home
- **BYOK** — you control which AI provider sees your document content. Use a local model for maximum confidentiality
- **Configurable session TTL** — default 30 days, configurable including "keep forever"
- **Remove AI Traces** — one-click removal of all Sidebar metadata from a document. Clean for discovery and filing

## Architecture

```
┌─────────────────────────────────────┐
│         Microsoft Word               │
│  ┌───────────────────────────────┐  │
│  │    Word Add-in (Task Pane)    │  │
│  │    Chat UI — 55+ doc tools    │  │
│  └──────────────┬────────────────┘  │
└─────────────────┼───────────────────┘
                  │ WebSocket (localhost:3001)
┌─────────────────┼───────────────────┐
│  ┌──────────────▼────────────────┐  │
│  │      Express + WSS Server     │  │
│  │  Agent Loop │ Sessions (AES)  │  │
│  │  References │ Context Mgmt    │  │
│  └───────┬──────────────────────┘  │
│          │                          │
│  ┌───────▼──────────────────────┐  │
│  │    LLM Router                │  │
│  │  OpenClaw │ OpenAI │ Anthropic│  │
│  │  Ollama   │ LM Studio│ Local │  │
│  └──────────────────────────────┘  │
│     Electron menu bar app           │
│     Your machine (localhost)        │
└─────────────────────────────────────┘
```

## Install

### Quick Start (macOS)

1. Download the latest `.dmg` from [Releases](https://github.com/yavarb/thesidebar/releases)
2. Drag to Applications
3. Right-click → Open (first launch only — no Apple Developer cert yet)
4. Launch — it lives in your menu bar
5. Open Word → The Sidebar appears in the task pane
6. Enter your API key (or OpenClaw gateway URL) in settings

### From Source

```bash
git clone https://github.com/yavarb/thesidebar.git
cd thesidebar
npm install
npm run dev
```

### Connect to OpenClaw

1. Make sure your OpenClaw gateway is running (`openclaw gateway start`)
2. In The Sidebar settings, set:
   - **Provider**: OpenClaw
   - **Gateway URL**: `http://localhost:18789`
   - **Token**: your OpenClaw gateway token
3. Select "openclaw" in the model picker
4. Start prompting — your OpenClaw agent now has full document control

## Configuration

Settings are accessible from the task pane UI or the menu bar icon.

| Setting | Default | Description |
|---|---|---|
| **AI Provider** | OpenClaw | Which LLM backend to use |
| **API Key / Token** | — | Your key for the selected provider |
| **Model** | Auto-detected | Model name (auto-detects context window) |
| **Context Budget** | 40% | Max % of context window to use |
| **Session TTL** | 30 days | How long conversations persist (`0` = forever) |
| **Theme** | System | Dark, light, or match system preference |
| **Reference Folders** | — | Case folders for RAG / filesystem hints |

## Contributing

PRs welcome. If you're a lawyer who codes (or a coder who lawyers), even better.

## License

MIT — do whatever you want with it.

---

*Built for lawyers, by a lawyer. Because your AI should work as hard as you do.*
