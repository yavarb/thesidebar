# The Sidebar

**An AI copilot for lawyers that actually edits your documents.** Open source, model-agnostic, runs entirely on your machine. Built by a litigation partner who uses it daily.

Microsoft Copilot costs $30/month, locks you into one model, and can't even add a footnote. The Sidebar gives your AI 37+ document tools and lets you pick any model — including one running on your own hardware.

<!-- TODO: screenshot -->
![Screenshot](docs/screenshot.png)

## Why This Exists

Legal work lives in Word. Every AI tool either (a) makes you copy-paste between a chat window and your document, or (b) is Microsoft Copilot — expensive, inflexible, and surprisingly bad at actual document manipulation.

The Sidebar sits next to your document in Word's task pane and does what Copilot can't: find/replace with regex, insert footnotes, apply styles, add comments, format tables — autonomously, in sequence, without you babysitting each step.

You pick the model. You hold the keys. Nothing leaves your machine except the API calls you explicitly authorize.

## Features

### 🤖 AI & Models
- **Model-agnostic** — OpenAI, Anthropic, OpenClaw, Ollama, LM Studio, or any OpenAI-compatible API
- **BYOK (Bring Your Own Keys)** — your credentials, your choice of provider, your data
- **Agentic loop** — AI plans and executes multi-step document edits autonomously
- **Smart context management** — dynamic context window detection per model, configurable budget (default 40%) to control API costs
- **Prompt caching** — Anthropic `cache_control` + OpenAI-optimized message ordering for up to 90% input cost reduction

### 📝 Document Tools (37+)
- Read document text, paragraphs, selections
- Find and replace (literal + regex)
- Insert, edit, and delete paragraphs
- Add and manage **footnotes**
- Insert and resolve **comments**
- Apply styles and formatting
- Table manipulation
- And more — the AI discovers available tools automatically

### 💬 Conversation
- **Per-document sessions** — each document gets its own conversation thread
- **Conversation history** — maintains context across exchanges within a session
- **Revert system** — per-exchange undo. Don't like what the AI did? One click to roll it back
- **Markdown rendering** — AI responses rendered with proper formatting

### 🖥️ App
- **macOS menu bar app** — Electron tray app, starts on login
- **Dark & light mode** — auto-detects system preference, manual toggle
- **Auto-updates** — via GitHub releases
- **Localhost only** — HTTP server on port 3001, nothing exposed to the network

## 🔒 Security & Privacy

This is a tool for lawyers. It's built like one.

| Safeguard | Detail |
|---|---|
| **Zero cloud dependency** | Server runs on `localhost:3001`. No data leaves your machine except outbound LLM API calls to the provider *you* choose. |
| **Encrypted session storage** | Conversation history encrypted at rest using **AES-256-GCM** with a machine-specific key. Moving a document to another computer does **not** expose prior conversations. |
| **Attorney work product protection** | Session data is encrypted and machine-bound. Even if someone obtains your document file, they cannot access the AI conversation history. |
| **No telemetry** | The Sidebar does not phone home, track usage, or send analytics. Zero. |
| **BYOK model** | You control which AI provider sees your document content. Use a local model (Ollama/LM Studio) for maximum confidentiality. |
| **Configurable session TTL** | Conversations auto-expire after 30 days by default. Configurable — including "keep forever." |
| **Machine-bound keys** | The encryption key (`~/.thesidebar/.machine-key`) is generated locally and never leaves your machine. |
| **Remove AI Traces** | One-click removal of all Sidebar metadata from a document. Strips custom properties, deletes session files. Produces a clean document with no trace of AI usage — critical for discovery and filing. |

**The short version:** Your documents stay on your machine. Your conversations are encrypted. You choose who sees what. There is no "our servers."

## Architecture

```
┌─────────────────────────────────────┐
│         Microsoft Word               │
│  ┌───────────────────────────────┐  │
│  │    Word Add-in (Task Pane)    │  │
│  │    React UI — chat + tools    │  │
│  └──────────────┬────────────────┘  │
└─────────────────┼───────────────────┘
                  │ HTTP (localhost:3001)
┌─────────────────┼───────────────────┐
│  ┌──────────────▼────────────────┐  │
│  │      Express Server           │  │
│  │  ┌─────────┐  ┌───────────┐  │  │
│  │  │ Agentic │  │  Session   │  │  │
│  │  │  Loop   │  │  Manager   │  │  │
│  │  │         │  │ (AES-256) │  │  │
│  │  └────┬────┘  └───────────┘  │  │
│  └───────┼──────────────────────┘  │
│          │                          │
│  ┌───────▼──────────────────────┐  │
│  │    LLM Provider Adapters     │  │
│  │  OpenAI │ Anthropic │ Local  │  │
│  └───────┬──────────────────────┘  │
│     Your machine (localhost)        │
└──────────┼──────────────────────────┘
           │ outbound HTTPS only
           ▼
    ┌──────────────┐
    │  LLM API     │
    │  (your pick) │
    └──────────────┘
```

## Install

### Quick Start (macOS)

1. Download the latest `.dmg` from [Releases](https://github.com/ybbathaee/thesidebar/releases)
2. Drag to Applications
3. Launch — it lives in your menu bar
4. Open Word → The Sidebar appears in the task pane
5. Enter your API key in settings and start working

### From Source

```bash
git clone https://github.com/ybbathaee/thesidebar.git
cd thesidebar
npm install
npm run dev
```

## Configuration

Settings are accessible from the task pane UI or the menu bar icon.

| Setting | Default | Description |
|---|---|---|
| **AI Provider** | OpenAI | Which LLM to use |
| **API Key** | — | Your key for the selected provider |
| **Model** | `gpt-4o` | Model name (auto-detects context window) |
| **Context Budget** | 40% | Max % of context window to use (controls cost) |
| **Session TTL** | 30 days | How long conversations persist (`0` = forever) |
| **Theme** | System | Dark, light, or match system preference |

### Local Models

For maximum confidentiality — nothing leaves your machine at all:

1. Install [Ollama](https://ollama.ai) or [LM Studio](https://lmstudio.ai)
2. Set provider to "OpenAI Compatible"
3. Point the base URL to your local server (e.g., `http://localhost:11434/v1`)
4. No API key needed

## Development

```bash
# Install dependencies
npm install

# Run dev server (hot reload)
npm run dev

# Build the Electron app
npm run build

# Run tests
npm test
```

### Project Structure

```
src/
├── server/         # Express server, agentic loop, LLM adapters
├── addin/          # Word task pane add-in (React)
├── electron/       # Menu bar app, auto-updater
└── shared/         # Types, crypto, session management
```

## Contributing

PRs welcome. If you're a lawyer who codes (or a coder who lawyers), even better.

1. Fork the repo
2. Create a feature branch
3. Make your changes
4. Open a PR with a clear description

Please open an issue first for large changes so we can discuss approach.

## License

MIT — do whatever you want with it.

---

*Built for lawyers, by a lawyer. Because your AI should work as hard as you do.*
