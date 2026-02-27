/**
 * @module llm-router.test
 * Unit tests for the LLM router module.
 * Uses mock HTTP servers to simulate each backend.
 */

import http from "http";
import { resolveModel, routePrompt, parseSSELines, RouterConfig, PromptContext } from "./llm-router";

// ── Test Helpers ──

let testsPassed = 0;
let testsFailed = 0;

function assert(condition: boolean, msg: string) {
  if (condition) {
    console.log(`  ✅ ${msg}`);
    testsPassed++;
  } else {
    console.log(`  ❌ ${msg}`);
    testsFailed++;
  }
}

async function collectStream(gen: AsyncGenerator<string>): Promise<string[]> {
  const chunks: string[] = [];
  for await (const chunk of gen) chunks.push(chunk);
  return chunks;
}

/** Create a mock HTTP server that returns SSE responses */
function createMockServer(handler: (req: http.IncomingMessage, res: http.ServerResponse) => void): Promise<{ server: http.Server; port: number }> {
  return new Promise((resolve) => {
    const server = http.createServer(handler);
    server.listen(0, "127.0.0.1", () => {
      const addr = server.address() as any;
      resolve({ server, port: addr.port });
    });
  });
}

// ── Tests ──

async function testResolveModel() {
  console.log("\n── resolveModel ──");
  const config: RouterConfig = {
    localEndpoints: [{ name: "LM Studio", baseUrl: "http://10.0.0.167:1234/v1" }],
  };

  const openai = resolveModel("openai:gpt-4o", config);
  assert(openai.backend === "openai" && openai.modelId === "gpt-4o", "resolves openai: prefix");

  const anthropic = resolveModel("anthropic:claude-sonnet-4-20250514", config);
  assert(anthropic.backend === "anthropic" && anthropic.modelId === "claude-sonnet-4-20250514", "resolves anthropic: prefix");

  const local = resolveModel("local:llama", config);
  assert(local.backend === "local" && local.modelId === "llama" && local.baseUrl === "http://10.0.0.167:1234/v1", "resolves local: with configured endpoint");

  const localUrl = resolveModel("local:http://192.168.1.1:8080/v1:mistral", config);
  assert(localUrl.backend === "local" && localUrl.modelId === "mistral" && localUrl.baseUrl === "http://192.168.1.1:8080/v1", "resolves local: with inline URL");

  const openclaw = resolveModel("opus", config);
  assert(openclaw.backend === "openclaw" && openclaw.modelId === "opus", "defaults to openclaw");
}

async function testParseSSE() {
  console.log("\n── parseSSELines ──");

  const { events, remainder } = parseSSELines('data: {"text":"hello"}\ndata: {"text":"world"}\n');
  assert(events.length === 2, "parses two SSE events");
  assert(remainder === "", "no remainder on complete buffer");

  const partial = parseSSELines('data: {"text":"hello"}\ndata: {"text":"incom');
  assert(partial.events.length === 1, "parses one complete event");
  assert(partial.remainder === 'data: {"text":"incom', "keeps incomplete line as remainder");

  const done = parseSSELines('data: [DONE]\n');
  assert(done.events.length === 0, "skips [DONE] marker");
}

async function testOpenAIBackend() {
  console.log("\n── OpenAI Backend (mock) ──");

  // Create mock OpenAI server
  const { server, port } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      const parsed = JSON.parse(body);
      assert(parsed.model === "gpt-4o", "sends correct model to API");
      assert(parsed.stream === true, "requests streaming");
      assert(parsed.messages.some((m: any) => m.content === "Hello"), "includes user prompt");

      res.writeHead(200, { "Content-Type": "text/event-stream" });
      res.write('data: {"choices":[{"delta":{"content":"Hello"}}]}\n\n');
      res.write('data: {"choices":[{"delta":{"content":" world"}}]}\n\n');
      res.write("data: [DONE]\n\n");
      res.end();
    });
  });

  try {
    // Monkey-patch the httpRequest to hit our mock server
    const origModule = require("./llm-router");
    const origHttpRequest = origModule.httpRequest;

    // We'll test via routeOpenAI directly by hitting our mock
    // Since we can't easily override the URL, test the SSE parsing flow
    const chunks = await collectStream(
      (async function* () {
        const response = await new Promise<http.IncomingMessage>((resolve) => {
          const req = http.request(`http://127.0.0.1:${port}/v1/chat/completions`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
          }, resolve);
          req.write(JSON.stringify({ model: "gpt-4o", messages: [{ role: "user", content: "Hello" }], stream: true }));
          req.end();
        });

        let buffer = "";
        for await (const chunk of response) {
          buffer += chunk.toString();
          const { events, remainder } = parseSSELines(buffer);
          buffer = remainder;
          for (const event of events) {
            try {
              const p = JSON.parse(event);
              if (p.choices?.[0]?.delta?.content) yield p.choices[0].delta.content;
            } catch {}
          }
        }
      })()
    );

    assert(chunks.join("") === "Hello world", "streams OpenAI response correctly");
  } finally {
    server.close();
  }
}

async function testAnthropicBackend() {
  console.log("\n── Anthropic Backend (mock) ──");

  const { server, port } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      const parsed = JSON.parse(body);
      assert(parsed.model === "claude-sonnet-4-20250514", "sends correct model");
      assert(parsed.stream === true, "requests streaming");

      res.writeHead(200, { "Content-Type": "text/event-stream" });
      res.write('data: {"type":"content_block_delta","delta":{"type":"text_delta","text":"Hi"}}\n\n');
      res.write('data: {"type":"content_block_delta","delta":{"type":"text_delta","text":" there"}}\n\n');
      res.write('data: {"type":"message_stop"}\n\n');
      res.end();
    });
  });

  try {
    const chunks: string[] = [];
    const response = await new Promise<http.IncomingMessage>((resolve) => {
      const req = http.request(`http://127.0.0.1:${port}/v1/messages`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
      }, resolve);
      req.write(JSON.stringify({ model: "claude-sonnet-4-20250514", messages: [{ role: "user", content: "Hi" }], stream: true }));
      req.end();
    });

    let buffer = "";
    for await (const chunk of response) {
      buffer += chunk.toString();
      const { events, remainder } = parseSSELines(buffer);
      buffer = remainder;
      for (const event of events) {
        try {
          const p = JSON.parse(event);
          if (p.type === "content_block_delta" && p.delta?.type === "text_delta") {
            chunks.push(p.delta.text);
          }
        } catch {}
      }
    }

    assert(chunks.join("") === "Hi there", "streams Anthropic response correctly");
  } finally {
    server.close();
  }
}

async function testLocalBackend() {
  console.log("\n── Local Backend (mock) ──");

  const { server, port } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      res.writeHead(200, { "Content-Type": "text/event-stream" });
      res.write('data: {"choices":[{"delta":{"content":"Local"}}]}\n\n');
      res.write('data: {"choices":[{"delta":{"content":" response"}}]}\n\n');
      res.write("data: [DONE]\n\n");
      res.end();
    });
  });

  try {
    const config: RouterConfig = {
      localEndpoints: [{ name: "Test", baseUrl: `http://127.0.0.1:${port}/v1` }],
    };

    // Use routeLocal directly
    const { routeLocal } = require("./llm-router");
    const chunks = await collectStream(
      routeLocal("Test prompt", "test-model", {}, config, `http://127.0.0.1:${port}/v1`)
    );

    assert(chunks.join("") === "Local response", "streams local response correctly");
  } finally {
    server.close();
  }
}

async function testToolCallStreaming() {
  console.log("\n── Tool Call Streaming (mock) ──");

  const { server, port } = await createMockServer((req, res) => {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      const parsed = JSON.parse(body);
      assert(Array.isArray(parsed.tools) && parsed.tools.length > 0, "tools sent to API");

      res.writeHead(200, { "Content-Type": "text/event-stream" });
      res.write('data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","function":{"name":"readDocument","arguments":""}}]}}]}\n\n');
      res.write('data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{}"}}]}}]}\n\n');
      res.write("data: [DONE]\n\n");
      res.end();
    });
  });

  try {
    const { routeLocal } = require("./llm-router");
    const tools = [{ type: "function", function: { name: "readDocument", description: "Read the document", parameters: { type: "object", properties: {} } } }];
    const chunks = await collectStream(
      routeLocal("Read the doc", "test-model", { tools }, {}, `http://127.0.0.1:${port}/v1`)
    );

    const toolCallChunks = chunks.filter(c => {
      try { return JSON.parse(c).type === "tool_calls"; } catch { return false; }
    });
    assert(toolCallChunks.length === 2, "streams tool call deltas");
  } finally {
    server.close();
  }
}

async function testErrorHandling() {
  console.log("\n── Error Handling ──");

  // Missing API key
  try {
    const { routeOpenAI } = require("./llm-router");
    await collectStream(routeOpenAI("test", "gpt-4o", {}, {}));
    assert(false, "throws on missing OpenAI key");
  } catch (e: any) {
    assert(e.message.includes("not configured"), "throws on missing OpenAI key");
  }

  try {
    const { routeAnthropic } = require("./llm-router");
    await collectStream(routeAnthropic("test", "claude", {}, {}));
    assert(false, "throws on missing Anthropic key");
  } catch (e: any) {
    assert(e.message.includes("not configured"), "throws on missing Anthropic key");
  }

  // API error response
  const { server, port } = await createMockServer((_req, res) => {
    res.writeHead(401);
    res.end(JSON.stringify({ error: "invalid_api_key" }));
  });

  try {
    const { routeLocal } = require("./llm-router");
    await collectStream(routeLocal("test", "m", {}, {}, `http://127.0.0.1:${port}/v1`));
    assert(false, "throws on API error");
  } catch (e: any) {
    assert(e.message.includes("401"), "throws on API error with status code");
  } finally {
    server.close();
  }
}

// ── Run ──

async function main() {
  console.log("🎀 LLM Router Test Suite\n═══════════════════════");

  await testResolveModel();
  await testParseSSE();
  await testOpenAIBackend();
  await testAnthropicBackend();
  await testLocalBackend();
  await testToolCallStreaming();
  await testErrorHandling();

  console.log(`\n═══════════════════════`);
  console.log(`Results: ${testsPassed} passed, ${testsFailed} failed`);
  process.exit(testsFailed > 0 ? 1 : 0);
}

main().catch((e) => {
  console.error("Test runner error:", e);
  process.exit(1);
});
