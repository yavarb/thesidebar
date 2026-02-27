# Contributing to The Sidebar

## Project Structure

```
thesidebar/
├── app/              # Office Add-in (Word task pane)
│   └── src/taskpane/ # Task pane UI (HTML, CSS, TypeScript)
├── server/           # Express + WebSocket server
│   └── src/
│       ├── index.ts          # Main server, routes, WebSocket
│       ├── llm-router.ts     # LLM routing (OpenAI, Anthropic, Local, OpenClaw)
│       ├── tools.ts          # Tool definitions for agentic loop
│       ├── agent-loop.ts     # Multi-turn tool-calling loop
│       └── settings.ts       # Config management
├── menubar/          # macOS menu bar app (Electron)
├── certs/            # Self-signed SSL certificates
├── install.sh        # Installation script
├── uninstall.sh      # Uninstallation script
├── start.sh          # Start server
├── stop.sh           # Stop server
└── test.sh           # Integration test suite
```

## Development Workflow

1. **TypeScript everywhere** — all code is TypeScript
2. **Test before commit** — run relevant test files and `bash test.sh`
3. **JSDoc all exports** — every exported function needs documentation
4. **Never break existing tests** — the 29 integration tests in `test.sh` must pass

## Adding a New Tool

1. Add the `ToolDefinition` in `server/src/tools.ts`
2. Add the `ToolEndpoint` mapping in the same file
3. Update `server/src/tools.test.ts` to validate the new tool
4. Update the tool table in `README.md`

## Adding a New LLM Backend

1. Create a new `route*` async generator function in `server/src/llm-router.ts`
2. Add the backend to the `Backend` type
3. Update `resolveModel()` to handle the new prefix
4. Add tests in `llm-router.test.ts`

## Code Style

- 2-space indentation
- No semicolons in imports
- Prefer `const` over `let`
- Use async generators for streaming responses
- Keep functions focused and well-documented
