/**
 * @module agent-loop.test
 * Tests for the agentic tool-calling loop.
 * Uses mock HTTP servers to simulate both the LLM and the The Sidebar API.
 */

import http from "http";
import { executeTool, ToolCall } from "./agent-loop";
import { parseSSELines } from "./llm-router";

let passed = 0;
let failed = 0;

function assert(cond: boolean, msg: string) {
  if (cond) { console.log(`  ✅ ${msg}`); passed++; }
  else { console.log(`  ❌ ${msg}`); failed++; }
}

function createMockServer(handler: (req: http.IncomingMessage, res: http.ServerResponse) => void): Promise<{ server: http.Server; port: number }> {
  return new Promise((resolve) => {
    const srv = http.createServer(handler);
    srv.listen(0, "127.0.0.1", () => {
      const addr = srv.address() as any;
      resolve({ server: srv, port: addr.port });
    });
  });
}

async function testExecuteTool() {
  console.log("\n── executeTool ──");

  // Mock The Sidebar server
  const { server, port } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => body += c);
    req.on("end", () => {
      if (req.url === "/api/document") {
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ ok: true, data: { text: "Hello world document" } }));
      } else if (req.url === "/api/navigate") {
        const parsed = JSON.parse(body);
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ ok: true, data: { navigatedTo: parsed.index } }));
      } else if (req.url?.startsWith("/api/paragraph/")) {
        const idx = req.url.split("/").pop();
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ ok: true, data: { index: parseInt(idx!), text: `Paragraph ${idx}` } }));
      } else {
        res.writeHead(404);
        res.end(JSON.stringify({ ok: false, error: "not found" }));
      }
    });
  });

  try {
    const baseUrl = `http://127.0.0.1:${port}`;

    // readDocument
    const r1 = await executeTool({ id: "call_1", name: "readDocument", arguments: {} }, baseUrl);
    assert(r1.result?.text === "Hello world document", "readDocument returns document text");
    assert(r1.error === undefined, "readDocument no error");

    // readParagraph (uses mapArgs)
    const r2 = await executeTool({ id: "call_2", name: "readParagraph", arguments: { index: 5 } }, baseUrl);
    assert(r2.result?.text === "Paragraph 5", "readParagraph returns correct paragraph");

    // navigateTo
    const r3 = await executeTool({ id: "call_3", name: "navigateTo", arguments: { index: 10 } }, baseUrl);
    assert(r3.result?.navigatedTo === 10, "navigateTo sends correct index");

    // Unknown tool
    const r4 = await executeTool({ id: "call_4", name: "nonexistentTool", arguments: {} }, baseUrl);
    assert(r4.error !== undefined && r4.error.includes("Unknown tool"), "unknown tool returns error");
  } finally {
    server.close();
  }
}

async function testAnthropicToolCallParsing() {
  console.log("\n── Anthropic Tool Call Parsing ──");

  const { server, port } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => body += c);
    req.on("end", () => {
      res.writeHead(200, { "Content-Type": "text/event-stream" });
      // Anthropic-style tool use events
      res.write('data: {"type":"content_block_start","content_block":{"type":"tool_use","id":"toolu_1","name":"readDocument"}}\n\n');
      res.write('data: {"type":"content_block_delta","delta":{"type":"input_json_delta","partial_json":"{}"}}\n\n');
      res.write('data: {"type":"message_delta","delta":{"stop_reason":"tool_use"}}\n\n');
      res.end();
    });
  });

  try {
    // Simulate Anthropic streaming directly
    const response = await new Promise<http.IncomingMessage>((resolve) => {
      const req = http.request(`http://127.0.0.1:${port}`, { method: "POST", headers: { "Content-Type": "application/json" } }, resolve);
      req.write("{}");
      req.end();
    });

    // Parse as if it were Anthropic streaming
    const toolCalls: ToolCall[] = [];
    let currentTool: { id: string; name: string; argsJson: string } | null = null;
    let buffer = "";

    for await (const chunk of response) {
      buffer += chunk.toString();
      const { events, remainder } = parseSSELines(buffer);
      buffer = remainder;

      for (const event of events) {
        try {
          const p = JSON.parse(event);
          if (p.type === "content_block_start" && p.content_block?.type === "tool_use") {
            currentTool = { id: p.content_block.id, name: p.content_block.name, argsJson: "" };
          } else if (p.type === "content_block_delta" && p.delta?.type === "input_json_delta") {
            if (currentTool) currentTool.argsJson += p.delta.partial_json;
          } else if (p.type === "message_delta" && p.delta?.stop_reason === "tool_use") {
            if (currentTool) {
              let args = {};
              try { args = JSON.parse(currentTool.argsJson); } catch {}
              toolCalls.push({ id: currentTool.id, name: currentTool.name, arguments: args });
              currentTool = null;
            }
          }
        } catch {}
      }
    }

    assert(toolCalls.length === 1, "parsed one Anthropic tool call");
    assert(toolCalls[0].name === "readDocument", "correct tool name");
    assert(toolCalls[0].id === "toolu_1", "correct tool ID");
  } finally {
    server.close();
  }
}

async function main() {
  console.log("🎀 Agent Loop Test Suite\n════════════════════════");

  await testExecuteTool();
  await testAnthropicToolCallParsing();

  console.log(`\n════════════════════════`);
  console.log(`Results: ${passed} passed, ${failed} failed`);
  process.exit(failed > 0 ? 1 : 0);
}

main().catch((e) => { console.error("Test error:", e); process.exit(1); });
