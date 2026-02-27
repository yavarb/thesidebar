/**
 * @module agent-loop.test
 * Tests for the agentic tool-calling loop.
 * Uses mock HTTP servers to simulate both the LLM and the The Sidebar API.
 */

import http from "http";
import { executeTool, collectLLMResponse, ToolCall } from "./agent-loop";
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

async function testCollectLLMResponse_TextOnly() {
  console.log("\n── collectLLMResponse (text only) ──");

  // Mock LLM server that returns plain text
  const { server, port } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => body += c);
    req.on("end", () => {
      res.writeHead(200, { "Content-Type": "text/event-stream" });
      res.write('data: {"choices":[{"delta":{"content":"The document"}}]}\n\n');
      res.write('data: {"choices":[{"delta":{"content":" looks good."}}]}\n\n');
      res.write("data: [DONE]\n\n");
      res.end();
    });
  });

  try {
    const config = { localEndpoints: [{ name: "test", baseUrl: `http://127.0.0.1:${port}/v1` }] };
    const { text, toolCalls } = await collectLLMResponse("Summarize", "local:test-model", {}, config);
    assert(text === "The document looks good.", "collects text response");
    assert(toolCalls.length === 0, "no tool calls in text-only response");
  } finally {
    server.close();
  }
}

async function testCollectLLMResponse_WithToolCalls() {
  console.log("\n── collectLLMResponse (with tool calls) ──");

  // Mock LLM server returning OpenAI-format tool calls
  const { server, port } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => body += c);
    req.on("end", () => {
      res.writeHead(200, { "Content-Type": "text/event-stream" });
      // Tool call: readDocument
      res.write('data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_abc","function":{"name":"readDocument","arguments":""}}]}}]}\n\n');
      res.write('data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{}"}}]}}]}\n\n');
      res.write("data: [DONE]\n\n");
      res.end();
    });
  });

  try {
    const config = { localEndpoints: [{ name: "test", baseUrl: `http://127.0.0.1:${port}/v1` }] };
    const { text, toolCalls } = await collectLLMResponse("Read it", "local:model", { tools: [{ type: "function", function: { name: "readDocument", description: "Read doc", parameters: { type: "object", properties: {} } } }] }, config);
    assert(toolCalls.length === 1, "extracts one tool call");
    assert(toolCalls[0].name === "readDocument", "correct tool name");
    assert(toolCalls[0].id === "call_abc", "correct tool call ID");
  } finally {
    server.close();
  }
}

async function testMultiTurnConversation() {
  console.log("\n── Multi-turn Tool-Calling Conversation (mock) ──");

  // Simulate a 2-turn conversation:
  // Turn 1: LLM asks to readDocument
  // Turn 2: LLM asks to getDocumentStats
  // Turn 3: LLM gives final text response

  let requestCount = 0;
  const { server: llmServer, port: llmPort } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => body += c);
    req.on("end", () => {
      requestCount++;
      res.writeHead(200, { "Content-Type": "text/event-stream" });

      if (requestCount === 1) {
        // First call: return tool call for readDocument
        res.write('data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","function":{"name":"readDocument","arguments":"{}"}}]}}]}\n\n');
      } else if (requestCount === 2) {
        // Second call: return tool call for getDocumentStats
        res.write('data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_2","function":{"name":"getDocumentStats","arguments":"{}"}}]}}]}\n\n');
      } else {
        // Third call: final text response
        res.write('data: {"choices":[{"delta":{"content":"The document has 500 words and 20 paragraphs."}}]}\n\n');
      }
      res.write("data: [DONE]\n\n");
      res.end();
    });
  });

  // Mock The Sidebar API server
  const { server: apiServer, port: apiPort } = await createMockServer((req, res) => {
    res.writeHead(200, { "Content-Type": "application/json" });
    if (req.url === "/api/document") {
      res.end(JSON.stringify({ ok: true, data: { text: "Lorem ipsum dolor sit amet..." } }));
    } else if (req.url === "/api/document/stats") {
      res.end(JSON.stringify({ ok: true, data: { words: 500, paragraphs: 20 } }));
    } else {
      res.end(JSON.stringify({ ok: true, data: {} }));
    }
  });

  try {
    // We can't easily test runAgentLoop end-to-end because it uses routePrompt
    // which hardcodes API URLs. Instead, test the components:

    // Verify executeTool works against our mock API
    const r1 = await executeTool({ id: "call_1", name: "readDocument", arguments: {} }, `http://127.0.0.1:${apiPort}`);
    assert(r1.result?.text === "Lorem ipsum dolor sit amet...", "turn 1: readDocument executed");

    const r2 = await executeTool({ id: "call_2", name: "getDocumentStats", arguments: {} }, `http://127.0.0.1:${apiPort}`);
    assert(r2.result?.words === 500, "turn 2: getDocumentStats executed");
    assert(r2.result?.paragraphs === 20, "turn 2: stats are correct");

    // Verify the LLM server gets called 3 times with increasing context
    const config = { localEndpoints: [{ name: "test", baseUrl: `http://127.0.0.1:${llmPort}/v1` }] };

    // Turn 1
    const t1 = await collectLLMResponse("Analyze the doc", "local:model", {}, config);
    assert(t1.toolCalls.length === 1 && t1.toolCalls[0].name === "readDocument", "turn 1: LLM requests readDocument");

    // Turn 2 (with tool results in messages)
    const t2 = await collectLLMResponse("Continue", "local:model", {
      messages: [
        { role: "assistant", content: "[Tool calls: readDocument]" },
        { role: "user", content: 'Tool results:\n\nTool "readDocument" (call_1):\n{"text":"Lorem ipsum..."}' },
      ],
    }, config);
    assert(t2.toolCalls.length === 1 && t2.toolCalls[0].name === "getDocumentStats", "turn 2: LLM requests getDocumentStats");

    // Turn 3 (final response)
    const t3 = await collectLLMResponse("Continue", "local:model", {
      messages: [
        { role: "assistant", content: "[Tool calls: readDocument]" },
        { role: "user", content: 'Tool results:\n\nTool "readDocument"...' },
        { role: "assistant", content: "[Tool calls: getDocumentStats]" },
        { role: "user", content: 'Tool results:\n\nTool "getDocumentStats"...' },
      ],
    }, config);
    assert(t3.toolCalls.length === 0, "turn 3: no more tool calls");
    assert(t3.text.includes("500 words"), "turn 3: final response includes stats");

    assert(requestCount === 3, `LLM called ${requestCount} times (expected 3)`);
  } finally {
    llmServer.close();
    apiServer.close();
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
    const config = { anthropicApiKey: "test-key" };

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
  await testCollectLLMResponse_TextOnly();
  await testCollectLLMResponse_WithToolCalls();
  await testMultiTurnConversation();
  await testAnthropicToolCallParsing();

  console.log(`\n════════════════════════`);
  console.log(`Results: ${passed} passed, ${failed} failed`);
  process.exit(failed > 0 ? 1 : 0);
}

main().catch((e) => { console.error("Test error:", e); process.exit(1); });
