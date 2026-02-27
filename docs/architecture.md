# The Sidebar Architecture

## Data Flow

```
User Prompt → Task Pane → WebSocket → Server → LLM Router → [Backend]
                                         ↓
                                    Agent Loop (if tool calls)
                                         ↓
                                    Execute Tools → API Endpoints → WebSocket → Task Pane → Word
                                         ↓
                                    Feed results back to LLM
                                         ↓
                                    Final Response → Task Pane
```

## Components

### Server (Express + WebSocket)
- **Port**: 3001 (configurable via SIDEBAR_PORT)
- **Protocol**: HTTPS with self-signed certs (falls back to HTTP)
- **WebSocket**: Bidirectional communication with task pane
- **37+ API endpoints** for document manipulation

### LLM Router
- Resolves model strings to backends via prefix convention
- All backends return `AsyncGenerator<string>` for streaming
- SSE stream parsing for OpenAI and Anthropic formats
- Tool call detection and accumulation

### Agent Loop
- Multi-turn conversation with tool calling
- Executes tools against The Sidebar's own API
- Max 15 iterations safety limit
- Supports both OpenAI and Anthropic tool-calling formats

### Settings
- Stored in `~/.thesidebar/config.json`
- File permissions: 600 (owner read/write only)
- API keys masked in GET responses
- Merge logic prevents masked values from overwriting real keys

### Task Pane
- Runs inside Word via Office.js
- WebSocket connection with auto-reconnect
- Settings UI panel (⚙️ gear icon)
- Model selector dropdown
- Chat-style prompt/response display

### Menu Bar App
- Electron tray app (hidden from dock)
- Polls server status every 5 seconds
- LaunchAgent plist for auto-start on login
