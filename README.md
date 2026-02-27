# The Sidebar

**An AI copilot for lawyers who work in Microsoft Word.**

Open source. Model-agnostic. Runs on your machine. Your keys, your data, your models.

Built by a litigation partner who got tired of paying Microsoft $30/seat/month for a copilot locked to one model that can't even do footnotes right.

---

## What It Does

The Sidebar is a Word task pane add-in powered by an agentic AI loop. You chat with it. It edits your document. It has **37+ tools** for manipulating Word documents — find/replace, footnotes, styles, formatting, tables, headers, track changes, and more.

It doesn't just suggest edits. It *makes* them. Multi-step, planned, executed.

### Features

- **Model-agnostic** — OpenAI, Anthropic, local models via Ollama/LM Studio, any OpenAI-compatible API
- **BYOK** (bring your own key) — nothing touches our servers because there are no servers
- **Runs entirely on localhost** — Express server on port 3001, nothing exposed inbound
- **Agentic loop** — AI plans and executes multi-step document edits autonomously
- **37+ document tools** — find/replace, footnotes, styles, formatting, tables, headers, comments, track changes, bookmarks, PDF export, and more
- **macOS menu bar app** — Electron tray app, starts on login, stays out of your way
- **Real-time streaming** — responses stream via WebSocket as the AI types
- **Settings UI** — configure API keys, models, and endpoints from the task pane

## Architecture

```
┌─────────────────────────────────────────────────┐
│                Microsoft Word                     │
│  ┌───────────────────────────────────────────┐   │
│  │           Task Pane (The Sidebar)          │   │
│  │  Chat UI → WebSocket → localhost:3001      │   │
│  └───────────────────────────────────────────┘   │
└─────────────────────────────────────────────────┘
                        ↕ WebSocket
┌─────────────────────────────────────────────────┐
│              Server (localhost:3001)              │
│                                                   │
│  Express + WebSocket server                       │
│  ├── LLM Router → OpenAI / Anthropic / Local     │
│  ├── Agent Loop (plan → tool calls → execute)    │
│  └── 37+ API endpoints → Office.js in Word       │
│                                                   │
│  Config: ~/.thesidebar/config.json               │
└─────────────────────────────────────────────────┘
                        ↕ HTTPS
┌─────────────────────────────────────────────────┐
│          LLM Backend (your choice)               │
│  OpenAI / Anthropic / Ollama / LM Studio / etc.  │
└─────────────────────────────────────────────────┘
```

Everything runs on your machine. The only outbound connections are to whatever LLM API you configure.

## Install

### Prerequisites

- macOS (Windows/Linux support planned)
- Node.js 18+
- Microsoft Word for Mac

### Quick Start

```bash
git clone https://github.com/yavarb/thesidebar.git
cd thesidebar

# Install dependencies
npm run install:all

# Generate self-signed certs and create config directory
./install.sh

# Build everything
npm run build

# Start the server
npm start
```

Then open Word → Insert → Add-ins → My Add-ins → The Sidebar.

### Configuration

Settings live in `~/.thesidebar/config.json`:

```json
{
  "anthropicApiKey": "sk-ant-...",
  "openaiApiKey": "sk-...",
  "model": "claude-sonnet-4-20250514",
  "customEndpoint": "http://localhost:11434/v1"
}
```

Or configure everything from the ⚙️ settings panel in the task pane.

## Development

```bash
# Run server in dev mode (auto-reload)
npm run dev

# Run the Word add-in dev server
cd app && npm run dev-server

# Run tests
npm test
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the full guide.

## Why This Exists

Microsoft Copilot for Word costs $30/user/month, only works with GPT-4, and can't do half the document operations a lawyer needs. No footnote support. No real find/replace. No agentic multi-step editing.

This is the tool I wanted. Now it's yours too.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT
