/**
 * @module agent-loop
 * Agentic tool-calling loop for non-OpenClaw models.
 *
 * Flow:
 * 1. Send prompt + document context + tool definitions to the LLM
 * 2. If LLM returns tool_calls, execute them against The Sidebar's API
 * 3. Feed tool results back to the LLM
 * 4. Repeat until LLM returns a final text response (no more tool calls)
 * 5. Stream final response back
 *
 * Supports both OpenAI and Anthropic tool-calling formats.
 */

import http from "http";
import https from "https";
import { URL } from "url";
import { TOOL_DEFINITIONS, TOOL_ENDPOINTS, ToolEndpoint } from "./tools";
import { routePrompt, RouterConfig, PromptContext, httpRequest } from "./llm-router";

/** Maximum number of tool-calling iterations to prevent infinite loops */
const MAX_ITERATIONS = 15;

/** Default The Sidebar server base URL for executing tools */
const DEFAULT_SERVER_URL = "http://localhost:3001";

// ── Types ──

/** A tool call extracted from an LLM response */
export interface ToolCall {
  id: string;
  name: string;
  arguments: Record<string, any>;
}

/** Result of executing a tool */
export interface ToolResult {
  toolCallId: string;
  name: string;
  result: any;
  error?: string;
}

/** Options for the agentic loop */
export interface AgentLoopOptions {
  /** User prompt */
  prompt: string;
  /** Model string (e.g., "openai:gpt-4o") */
  model: string;
  /** Router config (API keys, etc.) */
  config: RouterConfig;
  /** System prompt for the LLM */
  systemPrompt?: string;
  /** Current document context to include */
  documentContext?: string;
  /** The Sidebar server URL for executing tools (default: https://localhost:3001) */
  serverUrl?: string;
  /** Max tool-call iterations (default: 15) */
  maxIterations?: number;
  /** Callback for tool execution events */
  onToolCall?: (call: ToolCall) => void;
  /** Callback for tool results */
  onToolResult?: (result: ToolResult) => void;
}

// ── Tool Execution ──

/**
 * Execute a tool call against the The Sidebar server API.
 *
 * @param call - The tool call to execute
 * @param serverUrl - Base URL of the The Sidebar server
 * @returns The tool result
 */
export async function executeTool(call: ToolCall, serverUrl: string): Promise<ToolResult> {
  const endpoint = TOOL_ENDPOINTS[call.name];
  if (!endpoint) {
    return { toolCallId: call.id, name: call.name, result: null, error: `Unknown tool: ${call.name}` };
  }

  try {
    // Resolve the actual path and body using mapArgs
    let path = endpoint.path;
    let body: any = undefined;
    let query: Record<string, string> = {};

    if (endpoint.mapArgs) {
      const mapped = endpoint.mapArgs(call.arguments);
      path = mapped.path;
      body = mapped.body;
      query = mapped.query || {};
    } else if (endpoint.method === "POST" || endpoint.method === "PUT") {
      body = call.arguments;
    }

    // Build URL with query params
    const urlObj = new URL(path, serverUrl);
    for (const [k, v] of Object.entries(query)) {
      urlObj.searchParams.set(k, v);
    }

    const res = await httpRequest(urlObj.toString(), {
      method: endpoint.method,
      headers: { "Content-Type": "application/json" },
    }, body);

    let resBody = "";
    for await (const chunk of res) resBody += chunk.toString();

    const parsed = JSON.parse(resBody);
    if (parsed.ok) {
      return { toolCallId: call.id, name: call.name, result: parsed.data };
    } else {
      return { toolCallId: call.id, name: call.name, result: null, error: parsed.error };
    }
  } catch (e: any) {
    return { toolCallId: call.id, name: call.name, result: null, error: e.message };
  }
}

// ── Response Parsing ──

/**
 * Collect a full non-streaming response from an LLM via the router.
 * Accumulates all streamed chunks and parses tool calls from structured outputs.
 *
 * For OpenAI-format responses, tool calls come as JSON chunks with type:"tool_calls".
 * For Anthropic-format, they come as tool_use_start + tool_input_delta + tool_use_complete.
 */
export async function collectLLMResponse(
  prompt: string,
  model: string,
  context: PromptContext,
  config: RouterConfig
): Promise<{ text: string; toolCalls: ToolCall[] }> {
  let text = "";
  const toolCalls: ToolCall[] = [];

  // For OpenAI tool call accumulation
  const openaiToolAccum: Map<number, { id: string; name: string; args: string }> = new Map();

  // For Anthropic tool call accumulation
  let anthropicCurrentTool: { id: string; name: string; argsJson: string } | null = null;

  for await (const chunk of routePrompt(prompt, model, context, config)) {
    // Try to parse as structured event
    try {
      const parsed = JSON.parse(chunk);

      // OpenAI format: tool_calls delta
      if (parsed.type === "tool_calls" && Array.isArray(parsed.delta)) {
        for (const tc of parsed.delta) {
          const idx = tc.index ?? 0;
          if (!openaiToolAccum.has(idx)) {
            openaiToolAccum.set(idx, { id: tc.id || "", name: tc.function?.name || "", args: "" });
          }
          const accum = openaiToolAccum.get(idx)!;
          if (tc.id) accum.id = tc.id;
          if (tc.function?.name) accum.name = tc.function.name;
          if (tc.function?.arguments) accum.args += tc.function.arguments;
        }
        continue;
      }

      // Anthropic format: tool_use_start
      if (parsed.type === "tool_use_start") {
        anthropicCurrentTool = { id: parsed.id, name: parsed.name, argsJson: "" };
        continue;
      }

      // Anthropic format: tool_input_delta
      if (parsed.type === "tool_input_delta") {
        if (anthropicCurrentTool) {
          anthropicCurrentTool.argsJson += parsed.delta;
        }
        continue;
      }

      // Anthropic format: tool_use_complete
      if (parsed.type === "tool_use_complete") {
        if (anthropicCurrentTool) {
          let args: Record<string, any> = {};
          try { args = JSON.parse(anthropicCurrentTool.argsJson); } catch {}
          toolCalls.push({ id: anthropicCurrentTool.id, name: anthropicCurrentTool.name, arguments: args });
          anthropicCurrentTool = null;
        }
        continue;
      }

      // OpenClaw queued response — not a tool call scenario
      if (parsed.type === "openclaw_queued") {
        text = chunk;
        continue;
      }
    } catch {
      // Not JSON — it's plain text content
    }

    text += chunk;
  }

  // Finalize any accumulated OpenAI tool calls
  for (const [, accum] of openaiToolAccum) {
    let args: Record<string, any> = {};
    try { args = JSON.parse(accum.args); } catch {}
    toolCalls.push({ id: accum.id, name: accum.name, arguments: args });
  }

  return { text, toolCalls };
}

// ── Agentic Loop ──

/**
 * Run the agentic tool-calling loop.
 *
 * Sends the prompt to the LLM with tool definitions. If the LLM responds
 * with tool calls, executes them against The Sidebar's API, feeds results
 * back, and repeats until the LLM returns a final text response.
 *
 * @param options - Loop configuration
 * @yields Chunks of the final text response
 */
export async function* runAgentLoop(options: AgentLoopOptions): AsyncGenerator<string> {
  const {
    prompt,
    model,
    config,
    systemPrompt,
    documentContext,
    serverUrl = DEFAULT_SERVER_URL,
    maxIterations = MAX_ITERATIONS,
    onToolCall,
    onToolResult,
  } = options;

  // Build conversation history for multi-turn
  const messages: Array<{ role: "user" | "assistant" | "system"; content: string }> = [];

  // Initial context
  const context: PromptContext = {
    systemPrompt: systemPrompt || "You are a helpful assistant that can read and edit Word documents. Use the provided tools to interact with the document when needed.",
    documentContext,
    tools: TOOL_DEFINITIONS,
    messages,
  };

  let currentPrompt = prompt;
  let iteration = 0;

  while (iteration < maxIterations) {
    iteration++;

    const { text, toolCalls } = await collectLLMResponse(currentPrompt, model, context, config);

    // No tool calls — this is the final response
    if (toolCalls.length === 0) {
      yield text;
      return;
    }

    // Add assistant's tool call response to history
    messages.push({ role: "assistant", content: text || `[Tool calls: ${toolCalls.map(tc => tc.name).join(", ")}]` });

    // Execute each tool call
    const results: ToolResult[] = [];
    for (const call of toolCalls) {
      if (onToolCall) onToolCall(call);
      const result = await executeTool(call, serverUrl);
      if (onToolResult) onToolResult(result);
      results.push(result);
    }

    // Format tool results as the next message
    const toolResultText = results
      .map((r) => {
        const resultStr = r.error
          ? `Error: ${r.error}`
          : JSON.stringify(r.result, null, 2);
        return `Tool "${r.name}" (${r.toolCallId}):\n${resultStr}`;
      })
      .join("\n\n");

    messages.push({ role: "user", content: `Tool results:\n\n${toolResultText}` });

    // Continue the loop — the LLM will see the tool results and either
    // make more tool calls or return a final response
    currentPrompt = "Based on the tool results above, continue with your task. If you need more information, use the tools. If you have enough information, provide your final response.";
  }

  // Max iterations reached
  yield "[Agent loop reached maximum iterations. Partial results may be available in the conversation above.]";
}
