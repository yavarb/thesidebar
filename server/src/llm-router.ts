/**
 * @module llm-router
 * Unified LLM routing module with 4 backends:
 * 1. OpenClaw — OpenAI-compatible /v1/chat/completions on remote gateway
 * 2. OpenAI — Direct OpenAI API calls
 * 3. Anthropic — Direct Anthropic API calls
 * 4. Local — any OpenAI-compatible endpoint (e.g., LM Studio)
 *
 * All backends return AsyncGenerator<string> for streaming.
 */

import http from "http";
import https from "https";
import { URL } from "url";

// ── Types ──

/** Configuration for the LLM router */
export interface RouterConfig {
  /** OpenClaw gateway URL (e.g., "http://10.0.0.58:18789") */
  openclawUrl?: string;
  /** OpenClaw gateway auth token */
  openclawToken?: string;
  /** OpenAI API key for direct API calls */
  openaiApiKey?: string;
  /** Anthropic API key for direct API calls */
  anthropicApiKey?: string;
  /** Local OpenAI-compatible endpoints */
  localEndpoints?: LocalEndpoint[];
}

/** A local OpenAI-compatible endpoint */
export interface LocalEndpoint {
  /** Display name */
  name: string;
  /** Base URL (e.g., "http://10.0.0.167:1234/v1") */
  baseUrl: string;
}

/** Which backend to route to */
export type Backend = "openclaw" | "openai" | "anthropic" | "local" | "responses";

/** Model specification for routing */
export interface ModelSpec {
  /** Backend to use */
  backend: Backend;
  /** Model ID (e.g., "gpt-4o", "claude-sonnet-4-20250514", "local-llama") */
  modelId: string;
  /** For local backend: base URL of the endpoint */
  baseUrl?: string;
}

/** Context sent with the prompt */
export interface PromptContext {
  /** System prompt / instructions */
  systemPrompt?: string;
  /** Document context (e.g., current document text) */
  documentContext?: string;
  /** Conversation history */
  messages?: Array<{ role: "user" | "assistant" | "system"; content: string }>;
  /** Tool definitions (OpenAI function-calling format) */
  tools?: any[];
  /** Session user ID for OpenClaw context persistence */
  sessionUser?: string;
}

/** Cache statistics for prompt caching */
export interface CacheStats {
  hits: number;
  misses: number;
  anthropicCacheReadTokens: number;
  anthropicCacheCreationTokens: number;
  openaiCachedTokens: number;
}

/** Global cache stats singleton */
export const cacheStats: CacheStats = {
  hits: 0,
  misses: 0,
  anthropicCacheReadTokens: 0,
  anthropicCacheCreationTokens: 0,
  openaiCachedTokens: 0,
};

// ── Helpers ──

/**
 * Make an HTTP/HTTPS request and return the raw response stream.
 */
export function httpRequest(
  url: string,
  options: { method: string; headers: Record<string, string>; timeoutMs?: number },
  body?: any
): Promise<http.IncomingMessage> {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const mod = parsed.protocol === "https:" ? https : http;
    const req = mod.request(url, {
      method: options.method,
      headers: options.headers,
      rejectUnauthorized: false,
    }, (res) => resolve(res));
    req.on("error", reject);
    req.setTimeout(options.timeoutMs ?? 300000, () => {
      req.destroy(new Error(`Request timeout after ${options.timeoutMs ?? 300000}ms: ${url}`));
    });
    if (body !== undefined) {
      req.write(typeof body === "string" ? body : JSON.stringify(body));
    }
    req.end();
  });
}

/**
 * Parse Server-Sent Events (SSE) lines from a buffer, yielding data payloads.
 */
export function parseSSELines(buffer: string): { events: string[]; remainder: string } {
  const events: string[] = [];
  const lines = buffer.split("\n");
  let remainder = "";

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (i === lines.length - 1 && !buffer.endsWith("\n")) {
      remainder = line;
      break;
    }
    if (line.startsWith("data: ")) {
      const data = line.slice(6);
      if (data !== "[DONE]") {
        events.push(data);
      }
    }
  }

  return { events, remainder };
}

interface LLMExtract { text: string; reasoning?: string; }

function extractAssistantTextFromChatCompletion(parsed: any): string {
  return extractLLMResponse(parsed).text;
}

function extractLLMResponse(parsed: any): LLMExtract {
  const message = parsed?.choices?.[0]?.message;
  const content = message?.content;
  // reasoning_content is used by DeepSeek, QwQ, and other reasoning models
  const reasoning = message?.reasoning_content || message?.reasoning || undefined;

  let text = "";
  if (typeof content === "string") {
    text = content;
  } else if (Array.isArray(content)) {
    const parts = content
      .map((p: any) => {
        if (typeof p === "string") return p;
        if (typeof p?.text === "string") return p.text;
        if (typeof p?.content === "string") return p.content;
        return "";
      })
      .filter(Boolean);
    if (parts.length) text = parts.join("\n");
  } else if (typeof parsed?.choices?.[0]?.text === "string") {
    text = parsed.choices[0].text;
  }

  return { text, reasoning: typeof reasoning === "string" ? reasoning : undefined };
}

// ── Backend: OpenClaw ──

/**
 * Route prompt through OpenClaw's gateway via its OpenAI-compatible
 * /v1/chat/completions endpoint. The gateway must have chatCompletions
 * enabled (gateway.http.endpoints.chatCompletions.enabled = true).
 *
 * OpenClaw gateway docs: /opt/homebrew/lib/node_modules/openclaw/docs/gateway/openai-http-api.md
 *
 * @param prompt - User prompt text
 * @param model - Model string (passed as-is; OpenClaw resolves it)
 * @param context - System prompt, document context, conversation history
 * @param config - Must include openclawUrl (e.g., "http://10.0.0.58:18789")
 */
export async function* routeOpenClaw(
  prompt: string,
  model: string,
  context: PromptContext,
  config: RouterConfig
): AsyncGenerator<string> {
  const baseUrl = config.openclawUrl;
  if (!baseUrl) throw new Error("OpenClaw URL not configured. Set it in Settings → OpenClaw URL.");

  const messages: Array<{ role: string; content: string }> = [];
  if (context.systemPrompt) messages.push({ role: "system", content: context.systemPrompt });
  if (context.documentContext) messages.push({ role: "system", content: `Current document context:\n${context.documentContext}` });
  if (context.messages) messages.push(...context.messages);
  messages.push({ role: "user", content: prompt });

  const body: any = { model, messages, stream: true };
  if (context.sessionUser) body.user = context.sessionUser;
  if (context.tools?.length) body.tools = context.tools;

  const headers: Record<string, string> = { "Content-Type": "application/json", "x-openclaw-agent-id": "thesidebar" };

  if (config.openclawToken) {
    headers["Authorization"] = `Bearer ${config.openclawToken}`;
  }

  const url = baseUrl.replace(/\/$/, "") + "/v1/chat/completions";
  const res = await httpRequest(url, { method: "POST", headers }, body);

  if (res.statusCode && res.statusCode >= 400) {
    let errBody = "";
    for await (const chunk of res) errBody += chunk.toString();
    throw new Error(`OpenClaw API error ${res.statusCode}: ${errBody.slice(0, 500)}`);
  }

  // Stream SSE — same delta format as OpenAI chat completions
  let buffer = "";
  for await (const chunk of res) {
    buffer += chunk.toString();
    const { events, remainder } = parseSSELines(buffer);
    buffer = remainder;
    for (const event of events) {
      try {
        const parsed = JSON.parse(event);
        const delta = parsed.choices?.[0]?.delta;
        if (delta?.content) {
          yield delta.content;
        }
      } catch {}
    }
  }
}

// ── Backend: OpenAI ──

/**
 * Route prompt through OpenAI's API with streaming.
 */
export async function* routeOpenAI(
  prompt: string,
  model: string,
  context: PromptContext,
  config: RouterConfig
): AsyncGenerator<string> {
  if (!config.openaiApiKey) throw new Error("OpenAI API key not configured");

  const messages: Array<{ role: string; content: string }> = [];
  if (context.systemPrompt) messages.push({ role: "system", content: context.systemPrompt });
  if (context.documentContext) messages.push({ role: "system", content: `Current document context:\n${context.documentContext}` });
  if (context.messages) messages.push(...context.messages);
  messages.push({ role: "user", content: prompt });

  const body: any = { model, messages, stream: true, stream_options: { include_usage: true } };
  if (context.tools?.length) body.tools = context.tools;

  const res = await httpRequest("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${config.openaiApiKey}` },
  }, body);

  if (res.statusCode && res.statusCode >= 400) {
    let errBody = "";
    for await (const chunk of res) errBody += chunk.toString();
    throw new Error(`OpenAI API error ${res.statusCode}: ${errBody.slice(0, 500)}`);
  }

  let buffer = "";
  for await (const chunk of res) {
    buffer += chunk.toString();
    const { events, remainder } = parseSSELines(buffer);
    buffer = remainder;
    for (const event of events) {
      try {
        const parsed = JSON.parse(event);
        // Track OpenAI cache usage from final chunk
        if (parsed.usage?.prompt_tokens_details?.cached_tokens) {
          const cached = parsed.usage.prompt_tokens_details.cached_tokens;
          cacheStats.openaiCachedTokens += cached;
          cacheStats.hits++;
          console.log(`[cache] OpenAI cache hit: ${cached} cached tokens`);
        }
        const delta = parsed.choices?.[0]?.delta;
        if (delta?.tool_calls) {
          yield JSON.stringify({ type: "tool_calls", delta: delta.tool_calls });
        } else if (delta?.content) {
          yield delta.content;
        }
      } catch {}
    }
  }
}

// ── Backend: Anthropic ──

/**
 * Route prompt through Anthropic's API with streaming.
 */
export async function* routeAnthropic(
  prompt: string,
  model: string,
  context: PromptContext,
  config: RouterConfig
): AsyncGenerator<string> {
  if (!config.anthropicApiKey) throw new Error("Anthropic API key not configured");

  const messages: Array<{ role: string; content: string }> = [];
  if (context.messages) {
    // Anthropic rejects role:"system" in messages — filter them out
    for (const m of context.messages) {
      if (m.role !== "system") messages.push(m);
    }
  }
  messages.push({ role: "user", content: prompt });

  // Structure system as array of content blocks for optimal caching
  const systemBlocks: any[] = [];
  if (context.systemPrompt) {
    systemBlocks.push({ type: "text", text: context.systemPrompt, cache_control: { type: "ephemeral" } });
  }
  if (context.documentContext) {
    systemBlocks.push({ type: "text", text: `Current document context:\n${context.documentContext}`, cache_control: { type: "ephemeral" } });
  }

  const body: any = { model, messages, max_tokens: 4096, stream: true };
  if (systemBlocks.length) body.system = systemBlocks;

  if (context.tools?.length) {
    body.tools = context.tools.map((t: any) => ({
      name: t.function?.name || t.name,
      description: t.function?.description || t.description,
      input_schema: t.function?.parameters || t.parameters,
    }));
  }

  const res = await httpRequest("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": config.anthropicApiKey,
      "anthropic-version": "2023-06-01",
    },
  }, body);

  if (res.statusCode && res.statusCode >= 400) {
    let errBody = "";
    for await (const chunk of res) errBody += chunk.toString();
    throw new Error(`Anthropic API error ${res.statusCode}: ${errBody.slice(0, 500)}`);
  }

  let buffer = "";
  for await (const chunk of res) {
    buffer += chunk.toString();
    const { events, remainder } = parseSSELines(buffer);
    buffer = remainder;
    for (const event of events) {
      try {
        const parsed = JSON.parse(event);
        if (parsed.type === "message_start" && parsed.message?.usage) {
          const usage = parsed.message.usage;
          if (usage.cache_read_input_tokens) {
            cacheStats.anthropicCacheReadTokens += usage.cache_read_input_tokens;
            cacheStats.hits++;
            console.log(`[cache] Anthropic cache hit: ${usage.cache_read_input_tokens} tokens read from cache`);
          } else {
            cacheStats.misses++;
          }
          if (usage.cache_creation_input_tokens) {
            cacheStats.anthropicCacheCreationTokens += usage.cache_creation_input_tokens;
          }
        } else if (parsed.type === "message_delta" && parsed.usage) {
          // Final usage update
          if (parsed.usage.cache_read_input_tokens) {
            cacheStats.anthropicCacheReadTokens += parsed.usage.cache_read_input_tokens;
          }
        } else if (parsed.type === "content_block_delta") {
          if (parsed.delta?.type === "text_delta") yield parsed.delta.text;
          else if (parsed.delta?.type === "input_json_delta") {
            yield JSON.stringify({ type: "tool_input_delta", delta: parsed.delta.partial_json });
          }
        } else if (parsed.type === "content_block_start" && parsed.content_block?.type === "tool_use") {
          yield JSON.stringify({ type: "tool_use_start", id: parsed.content_block.id, name: parsed.content_block.name });
        } else if (parsed.type === "content_block_stop") {
          // content_block_stop fires after each content block completes
          // If we have a pending tool, this signals its completion
          yield JSON.stringify({ type: "content_block_stop" });
        } else if (parsed.type === "message_delta" && parsed.delta?.stop_reason === "tool_use") {
          yield JSON.stringify({ type: "tool_use_complete" });
        }
      } catch {}
    }
  }
}

// ── Backend: Local ──

/**
 * Route prompt through a local OpenAI-compatible endpoint with streaming.
 */
export async function* routeLocal(
  prompt: string,
  model: string,
  context: PromptContext,
  _config: RouterConfig,
  baseUrl: string
): AsyncGenerator<string> {
  // Many local models don't support the "system" role in their chat templates.
  // Fold system content into the first user message to avoid jinja errors.
  const systemParts: string[] = [];
  if (context.systemPrompt) systemParts.push(context.systemPrompt);
  if (context.documentContext) systemParts.push(`Current document context:\n${context.documentContext}`);

  const messages: Array<{ role: string; content: string }> = [];
  if (context.messages) messages.push(...context.messages.filter(m => m.role !== "system"));

  const userContent = systemParts.length
    ? `${systemParts.join("\n\n")}\n\n${prompt}`
    : prompt;
  messages.push({ role: "user", content: userContent });

  const body: any = { model, messages, stream: false };
  if (context.tools?.length) body.tools = context.tools;

  const base = baseUrl.replace(/\/$/, "");
  const url = base.endsWith("/v1") ? `${base}/chat/completions` : `${base}/v1/chat/completions`;

  const res = await httpRequest(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  }, body);

  let raw = "";
  for await (const chunk of res) raw += chunk.toString();

  if (res.statusCode && res.statusCode >= 400) {
    throw new Error(`Local LLM error ${res.statusCode}: ${raw.slice(0, 500)}`);
  }

  try {
    const parsed = JSON.parse(raw);
    const { text, reasoning } = extractLLMResponse(parsed);
    if (reasoning) {
      yield JSON.stringify({ type: "reasoning", content: reasoning });
    }

    // Parse tool calls from the response (OpenAI format)
    const message = parsed?.choices?.[0]?.message;
    if (message?.tool_calls && Array.isArray(message.tool_calls)) {
      // Yield tool calls in the same format as routeOpenAI streaming
      const delta = message.tool_calls.map((tc: any, i: number) => ({
        index: i,
        id: tc.id,
        function: {
          name: tc.function?.name,
          arguments: tc.function?.arguments || "{}",
        },
      }));
      yield JSON.stringify({ type: "tool_calls", delta });
    }

    // If content is empty but reasoning exists, use reasoning as the response
    const displayText = text || reasoning || "";
    if (displayText) {
      if (reasoning && !text) {
        // Model put everything in reasoning field — yield as reasoning event then as text
        yield JSON.stringify({ type: "reasoning", content: reasoning });
      }
      yield displayText;
    } else if (!message?.tool_calls?.length && raw.trim()) {
      yield raw.trim();
    }
  } catch {
    if (raw.trim()) yield raw.trim();
  }
}

// ── Backend: OpenAI Responses API (Codex models) ──

const CODEX_MODEL_PATTERNS = ["codex", "gpt-5.2", "gpt-5.3"];

export function isResponsesModel(modelId: string): boolean {
  const lower = modelId.toLowerCase();
  return CODEX_MODEL_PATTERNS.some(p => lower.includes(p));
}

export async function* routeResponses(
  prompt: string,
  model: string,
  context: PromptContext,
  config: RouterConfig
): AsyncGenerator<string> {
  const key = config.openaiApiKey;
  if (!key) throw new Error("No OpenAI API key configured");

  // Build input: system instructions + conversation history + current prompt
  const input: any[] = [];
  if (context.messages) {
    for (const m of context.messages) {
      if (m.role === "system") continue; // handled via instructions
      input.push({ role: m.role, content: m.content });
    }
  }
  input.push({ role: "user", content: prompt });

  // Convert tool definitions from chat format to Responses API format
  const tools = context.tools?.map((t: any) => ({
    type: "function" as const,
    name: t.function.name,
    description: t.function.description,
    parameters: t.function.parameters,
  }));

  // Build system instructions
  const instructions = [
    context.systemPrompt || "",
    context.documentContext ? `Current document context:\n${context.documentContext}` : "",
  ].filter(Boolean).join("\n\n");

  const body: any = {
    model,
    input,
    stream: false,
  };
  if (instructions) body.instructions = instructions;
  if (tools?.length) body.tools = tools;

  const res = await httpRequest("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${key}`,
    },
  }, body);

  let raw = "";
  for await (const chunk of res) raw += chunk.toString();

  if (res.statusCode && res.statusCode >= 400) {
    throw new Error(`OpenAI Responses API error ${res.statusCode}: ${raw.slice(0, 500)}`);
  }

  try {
    const parsed = JSON.parse(raw);
    const output = parsed.output || [];

    // Extract reasoning if present
    const reasoning = parsed.reasoning?.summary;
    if (reasoning) {
      yield JSON.stringify({ type: "reasoning", content: reasoning });
    }

    // Process output items
    const toolCalls: any[] = [];
    let textContent = "";

    for (const item of output) {
      if (item.type === "message") {
        for (const block of (item.content || [])) {
          if (block.type === "output_text") {
            textContent += block.text;
          }
        }
      } else if (item.type === "function_call") {
        toolCalls.push({
          index: toolCalls.length,
          id: item.call_id || item.id,
          function: {
            name: item.name,
            arguments: item.arguments || "{}",
          },
        });
      }
    }

    // Yield tool calls if any
    if (toolCalls.length > 0) {
      yield JSON.stringify({ type: "tool_calls", delta: toolCalls });
    }

    if (textContent) {
      yield textContent;
    }
  } catch {
    if (raw.trim()) yield raw.trim();
  }
}

// ── Main Router ──

/**
 * Resolve a model string to a backend + model ID.
 * Convention:
 *   "openai:gpt-4o" → backend=openai, modelId=gpt-4o
 *   "anthropic:claude-sonnet-4-20250514" → backend=anthropic
 *   "local:my-model" → backend=local (uses first local endpoint)
 *   "local:http://host:port/v1:model-name" → with specific endpoint
 *   anything else → backend=openclaw
 */
export function resolveModel(model: string, config: RouterConfig): ModelSpec {
  if (model.startsWith("openai:")) {
    const modelId = model.slice(7);
    if (isResponsesModel(modelId)) return { backend: "responses", modelId };
    return { backend: "openai", modelId };
  }
  if (model.startsWith("anthropic:")) {
    return { backend: "anthropic", modelId: model.slice(10) };
  }
  if (model.startsWith("local:")) {
    const rest = model.slice(6);
    const urlMatch = rest.match(/^(https?:\/\/[^:]+:\d+\/v\d+):(.+)$/);
    if (urlMatch) {
      return { backend: "local", modelId: urlMatch[2], baseUrl: urlMatch[1] };
    }
    const endpoint = config.localEndpoints?.[0];
    return { backend: "local", modelId: rest, baseUrl: endpoint?.baseUrl || "http://localhost:1234/v1" };
  }
  return { backend: "openclaw", modelId: model };
}

/**
 * Route a prompt to the appropriate LLM backend and stream the response.
 *
 * @param prompt - User prompt text
 * @param model - Model string (e.g., "openai:gpt-4o", "anthropic:claude-sonnet-4-20250514", "local:llama")
 * @param context - Additional context (system prompt, document, history, tools)
 * @param config - Router configuration (API keys, endpoints)
 * @yields Chunks of the response text
 */
export async function* routePrompt(
  prompt: string,
  model: string,
  context: PromptContext,
  config: RouterConfig
): AsyncGenerator<string> {
  const spec = resolveModel(model, config);

  switch (spec.backend) {
    case "openai":
      yield* routeOpenAI(prompt, spec.modelId, context, config);
      break;
    case "responses":
      yield* routeResponses(prompt, spec.modelId, context, config);
      break;
    case "anthropic":
      yield* routeAnthropic(prompt, spec.modelId, context, config);
      break;
    case "local":
      yield* routeLocal(prompt, spec.modelId, context, config, spec.baseUrl!);
      break;
    case "openclaw":
    default:
      yield* routeOpenClaw(prompt, spec.modelId, context, config);
      break;
  }
}
