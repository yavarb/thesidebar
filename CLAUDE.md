# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

The Sidebar is an AI-powered legal writing assistant for Microsoft Word. It's a macOS Electron menu bar app that runs an Express server locally, connecting a Word task pane add-in to various LLM providers (OpenClaw, OpenAI, Anthropic, Ollama/LM Studio). It provides 55+ document manipulation tools via an agentic loop.

## Architecture

Three-workspace monorepo:

- **`server/`** — Express + WebSocket backend (port 3001). Core modules: `index.ts` (routes, WebSocket, prompt orchestration), `llm-router.ts` (multi-provider LLM routing via `AsyncGenerator<string>`), `agent-loop.ts` (multi-turn tool calling, max 15 iterations), `tools.ts` (55+ tool definitions + endpoint mappings), `sessions.ts` (AES-256-GCM encrypted per-document sessions), `references.ts` (RAG with TF-IDF/embeddings), `context.ts` (dynamic context window budgeting), `memory.ts` (global + per-document learned facts).
- **`app/`** — Office Add-in task pane (Webpack + Babel). `taskpane.ts` is the main UI (~3500 lines): chat interface, settings panel, model selector, quick actions, Word document API integration via Office.js. Communicates with server via WebSocket.
- **`electron/`** — macOS menu bar tray app. `main.ts` forks the server as a child process, handles first-run setup (cert generation, manifest sideloading), auto-update, and logging.

Data flow: User prompt → Task Pane (WebSocket) → Express Server → LLM Router → Provider → [Agent Loop if tool calls] → Tool execution against document API → Response streamed back via WebSocket.

LLM backends are selected by model string prefix: `openai:`, `anthropic:`, `openclaw:`, `local:`. All return `AsyncGenerator<string>` for uniform streaming.

## Build & Development Commands

```bash
# Install all workspaces
npm run install:all

# Build individual workspaces
npm run build:server          # tsc (server/src → server/dist)
npm run build:app             # webpack (app/src → app/dist)
npm run build:electron        # tsc (electron/ → electron/dist)
npm run build:all             # all three

# Production build (all + electron-builder macOS DMG)
npm run build

# Development
npm run dev                   # build:electron + launch Electron
cd server && npm run dev      # tsx watch server (hot reload)
cd app && npm run dev-server  # webpack-dev-server on port 3000

# Testing
npm test                      # runs server tests
cd server && npm test         # server unit tests (vitest/jest)
bash test.sh                  # integration tests (curl-based, 29+ endpoint tests)

# Linting (app only)
cd app && npm run lint
cd app && npm run lint:fix
```

## Code Conventions

- TypeScript everywhere, 2-space indentation, no semicolons in imports
- Prefer `const` over `let`
- Use async generators (`AsyncGenerator<string>`) for all streaming responses
- JSDoc all exported functions
- Config stored at `~/.thesidebar/config.json` (600 permissions), sessions at `~/.thesidebar/sessions/`

## Adding a New Tool

1. Add `ToolDefinition` in `server/src/tools.ts`
2. Add `ToolEndpoint` mapping in the same file
3. Add test in `server/src/tools.test.ts`
4. Update tool table in `README.md`

## Adding a New LLM Backend

1. Create a `route*` async generator in `server/src/llm-router.ts`
2. Add to the `Backend` type union
3. Update `resolveModel()` for the new prefix
4. Add tests in `server/src/llm-router.test.ts`

## Key Patterns

- WebSocket is the primary client-server transport (JSON messages with `type` field: `prompt`, `prompt_ack`, `prompt_progress`, `prompt_response`)
- The agent loop executes tools by making HTTP calls against the server's own REST API endpoints
- Settings GET responses mask sensitive fields with `"****"`; the merge logic in POST prevents masked values from overwriting real keys
- Session encryption uses a machine-specific key at `~/.thesidebar/.machine-key`
- Document index (paragraph cache) has a 30-second TTL
- The app targets ES5 (IE11 compat for Office.js) while server targets ES2020
