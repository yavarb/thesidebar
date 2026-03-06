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

import { URL } from "url";
import { TOOL_DEFINITIONS, TOOL_ENDPOINTS } from "./tools";
import { routePrompt, RouterConfig, PromptContext, ConversationMessage, httpRequest } from "./llm-router";
import { isTrackChangesEnabled } from "./track-changes";

/** Models exit naturally when done (no tool calls = loop exits). No cap needed. */
const MAX_ITERATIONS = Infinity;

/** Tools that modify document content */
const MODIFYING_TOOL_NAMES = new Set([
  "updateParagraph", "replaceParagraph", "replaceSelection", "editSelection",
  "findReplace", "insert", "format", "setStyleFont", "addFootnote",
  "updateFootnote", "addComment", "batch",
]);

/** Tools that can be batched together */
const BATCHABLE_TOOLS = new Set(["updateParagraph", "replaceParagraph"]);

/** Navigate to a paragraph in Word before editing it */
async function navigateToParagraph(serverUrl: string, index: number): Promise<void> {
  try {
    await httpRequest(new URL("/api/navigate", serverUrl).toString(), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    }, { index });
    // Small delay so user sees the scroll
    await new Promise(r => setTimeout(r, 50));
  } catch {}
}

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

  /** Callback for tool execution events */
  onToolCall?: (call: ToolCall) => void;
  /** Callback for tool results */
  onToolResult?: (result: ToolResult) => void;
  /** Session user ID for OpenClaw context persistence */
  sessionUser?: string;
  /** Conversation history for non-OpenClaw models */
  conversationHistory?: { role: string; content: string }[];
  /** Abort signal — when triggered, the loop stops after the current iteration */
  signal?: AbortSignal;
}

/** Result of the agent loop including change summaries */
export interface AgentLoopResult {
  response: string;
  changeSummaries: string[];
}

// ── Change Summary Helpers ──

const PARAGRAPH_EDIT_TOOLS = new Set(["updateParagraph", "replaceParagraph"]);

/** Truncate text to maxLen chars, adding ellipsis if needed */
function truncate(text: string, maxLen: number = 50): string {
  if (!text) return "";
  const cleaned = text.replace(/\n/g, " ").trim();
  return cleaned.length > maxLen ? cleaned.substring(0, maxLen) + "..." : cleaned;
}

/** Fetch a paragraph's current text from the server (best-effort) */
async function fetchParagraphText(serverUrl: string, index: number): Promise<string | null> {
  try {
    const url = new URL(`/api/paragraph/${index}`, serverUrl);
    url.searchParams.set("compact", "true");
    const res = await httpRequest(url.toString(), { method: "GET", headers: {} });
    let body = "";
    for await (const chunk of res) body += chunk.toString();
    const parsed = JSON.parse(body);
    return parsed?.ok ? parsed.data?.text ?? null : null;
  } catch {
    return null;
  }
}

/** Generate a change summary string for a tool call */
function generateChangeSummary(toolName: string, args: Record<string, any>, beforeText: string | null, result: any): string | null {
  switch (toolName) {
    case "updateParagraph":
    case "replaceParagraph": {
      const idx = args.index ?? args.paragraphIndex ?? "?";
      const newText = args.text ?? "";
      if (beforeText !== null) {
        return `\u00b6 ${idx}: "${truncate(beforeText)}" \u2192 "${truncate(newText)}"`;
      }
      return `\u00b6 ${idx}: Updated to "${truncate(newText)}"`;
    }
    case "replaceSelection":
    case "editSelection": {
      const newText = args.text ?? args.replacement ?? "";
      if (beforeText !== null) {
        const wordsBefore = beforeText.split(/\s+/).filter(Boolean).length;
        const wordsAfter = newText.split(/\s+/).filter(Boolean).length;
        const saved = wordsBefore - wordsAfter;
        const pct = wordsBefore > 0 ? Math.round((saved / wordsBefore) * 100) : 0;
        if (saved > 0) {
          return `Selection: replaced with "${truncate(newText)}" — saved ${saved} words (${pct}% reduction)`;
        }
        return `Selection: replaced with "${truncate(newText)}" (${wordsBefore} → ${wordsAfter} words)`;
      }
      return `Selection: replaced with "${truncate(newText)}"`;
    }
    case "findReplace": {
      const count = result?.replacedCount ?? "?";
      return `Found ${count} replacements: "${truncate(args.text ?? "")}" \u2192 "${truncate(args.replacement ?? "")}"`;
    }
    case "insert": {
      const idx = args.paragraphIndex ?? "end";
      const loc = args.location ?? "after";
      return `+ \u00b6 ${idx} (${loc}): Inserted "${truncate(args.text ?? "")}"`;
    }
    case "addFootnote":
      return `Added footnote: "${truncate(args.footnoteText ?? "")}"`;
    case "updateFootnote": {
      const idx = args.index ?? "?";
      return `Updated footnote ${idx}: "${truncate(args.text ?? "")}"`;
    }
    case "addComment":
      return `Added comment: "${truncate(args.commentText ?? "")}"`;
    case "format": {
      const parts: string[] = [];
      if (args.bold !== undefined) parts.push(args.bold ? "bold" : "unbold");
      if (args.italic !== undefined) parts.push(args.italic ? "italic" : "unitalic");
      if (args.underline !== undefined) parts.push(args.underline ? "underline" : "no underline");
      if (args.style) parts.push(`style "${args.style}"`);
      if (args.color) parts.push(`color ${args.color}`);
      return `Format "${truncate(args.text ?? "")}": ${parts.join(", ") || "formatting applied"}`;
    }
    case "setStyleFont": {
      const parts: string[] = [];
      if (args.fontName) parts.push(args.fontName);
      if (args.fontSize) parts.push(`${args.fontSize}pt`);
      return `Style "${args.styleName ?? "?"}": ${parts.join(", ") || "font updated"}`;
    }
    case "batch": {
      const ops = args.operations;
      if (Array.isArray(ops)) {
        return `Batch: ${ops.length} operations (${ops.map((o: any) => o.command).join(", ")})`;
      }
      return `Batch operation`;
    }
    default:
      return null;
  }
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
    onToolCall,
    onToolResult,
    sessionUser,
    conversationHistory: priorHistory,
  } = options;

  // Build conversation history for multi-turn
  const messages: ConversationMessage[] = [];

  // Prepend prior conversation history if provided (for non-OpenClaw models)
  if (priorHistory?.length) {
    for (const msg of priorHistory) {
      messages.push({ role: msg.role as "user" | "assistant", content: msg.content });
    }
  }

  // Initial context — structured for optimal prompt caching:
  // systemPrompt is stable (cacheable), documentContext changes only when doc changes
  const context: PromptContext = {
    systemPrompt: (systemPrompt || "You are a helpful assistant that can read and edit Word documents. Use the provided tools to interact with the document when needed.") + (isTrackChangesEnabled() ? "\nNote: Track Changes is enabled in the document. The user will review your edits. Be precise about what you are changing and why." : ""),
    documentContext,
    tools: TOOL_DEFINITIONS,
    messages,
    sessionUser,
    signal: options.signal,
  };

  let currentPrompt = prompt;
  let iteration = 0;
  const changeSummaries: string[] = [];

  while (true) {
    if (options.signal?.aborted) return;
    iteration++;

    // Force tool use on first iteration so models act instead of narrate
    if (iteration === 1 && context.tools?.length) {
      context.tool_choice = "required";
    } else {
      context.tool_choice = "auto";
    }

    // Stream text chunks immediately while collecting tool calls
    let fullText = "";
    const toolCalls: ToolCall[] = [];
    const openaiToolAccum: Map<number, { id: string; name: string; args: string }> = new Map();
    let anthropicCurrentTool: { id: string; name: string; argsJson: string } | null = null;

    for await (const chunk of routePrompt(currentPrompt, model, context, config)) {
      if (options.signal?.aborted) return;
      let isToolEvent = false;
      try {
        const parsed = JSON.parse(chunk);
        if (parsed.type === "tool_calls" && Array.isArray(parsed.delta)) {
          isToolEvent = true;
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
        } else if (parsed.type === "tool_use_start") {
          isToolEvent = true;
          anthropicCurrentTool = { id: parsed.id, name: parsed.name, argsJson: "" };
          console.log("[agent] Anthropic tool_use_start:", parsed.name);
        } else if (parsed.type === "tool_input_delta") {
          isToolEvent = true;
          if (anthropicCurrentTool) anthropicCurrentTool.argsJson += parsed.delta;
        } else if (parsed.type === "tool_use_complete" || parsed.type === "content_block_stop") {
          isToolEvent = true;
          if (anthropicCurrentTool) {
            let args: Record<string, any> = {};
            try { args = JSON.parse(anthropicCurrentTool.argsJson); } catch {}
            toolCalls.push({ id: anthropicCurrentTool.id, name: anthropicCurrentTool.name, arguments: args });
            console.log("[agent] Anthropic tool complete:", anthropicCurrentTool.name, JSON.stringify(args));
            anthropicCurrentTool = null;
          }
        } else if (parsed.type === "reasoning") {
          isToolEvent = true;
          yield "\n__REASONING__" + JSON.stringify({ content: parsed.content });
        } else if (parsed.type === "openclaw_queued") {
          fullText = chunk;
          isToolEvent = true;
        }
      } catch {
        // Not JSON — plain text content
      }
      if (!isToolEvent) {
        fullText += chunk;
        yield chunk; // Stream text to client immediately
      }
    }

    // Finalize OpenAI tool calls
    for (const [, accum] of openaiToolAccum) {
      let args: Record<string, any> = {};
      try { args = JSON.parse(accum.args); } catch {}
      toolCalls.push({ id: accum.id, name: accum.name, arguments: args });
    }

    console.log("[agent] Iteration", iteration, "- toolCalls:", toolCalls.length, toolCalls.map(t => t.name).join(", "));

    // No tool calls — final response already streamed
    if (toolCalls.length === 0) {
      if (changeSummaries.length > 0) {
        yield "\n__CHANGE_SUMMARIES__" + JSON.stringify(changeSummaries);
      }
      return;
    }

    // Signal that tools are about to execute (UI should keep "working" indicator)
    yield "\n__TOOL_PHASE__" + JSON.stringify({ toolCount: toolCalls.length, tools: toolCalls.map(tc => tc.name) });

    // Record the user prompt that triggered these tool calls.
    // routePrompt received it as a separate param, not in messages[].
    if (currentPrompt) {
      messages.push({ role: "user", content: currentPrompt });
    }

    // Add assistant's response with structured tool_calls to history
    messages.push({
      role: "assistant",
      content: fullText || "",
      tool_calls: toolCalls.map(tc => ({
        id: tc.id,
        type: "function" as const,
        function: { name: tc.name, arguments: JSON.stringify(tc.arguments) },
      })),
    });

    // Execute tool calls — batch when possible, navigate before edits
    const results: ToolResult[] = [];

    // Check if we can batch: multiple batchable tools in the same response
    const batchableCalls = toolCalls.filter(tc => BATCHABLE_TOOLS.has(tc.name));
    const nonBatchableCalls = toolCalls.filter(tc => !BATCHABLE_TOOLS.has(tc.name));
    const shouldBatch = batchableCalls.length > 1;

    if (shouldBatch) {
      // Execute batchable calls as a single batch operation
      for (const call of batchableCalls) {
        if (onToolCall) onToolCall(call);
      }

      // Snapshot before-texts for all batchable calls
      const beforeTexts: (string | null)[] = [];
      for (const call of batchableCalls) {
        if (PARAGRAPH_EDIT_TOOLS.has(call.name)) {
          const idx = call.arguments.index ?? call.arguments.paragraphIndex;
          if (idx !== undefined && typeof idx === "number") {
            beforeTexts.push(await fetchParagraphText(serverUrl, idx));
          } else {
            beforeTexts.push(null);
          }
        } else {
          beforeTexts.push(null);
        }
      }

      // Build batch operations
      const operations = batchableCalls.map(call => ({
        command: call.name,
        params: call.arguments,
      }));

      const batchCall: ToolCall = {
        id: batchableCalls[0].id,
        name: "batch",
        arguments: { operations },
      };

      const batchResult = await executeTool(batchCall, serverUrl);
      // Map batch results back to individual tool calls
      const batchResults = batchResult.result?.results || [];
      for (let i = 0; i < batchableCalls.length; i++) {
        const call = batchableCalls[i];
        const individualResult: ToolResult = {
          toolCallId: call.id,
          name: call.name,
          result: batchResults[i] || null,
          error: batchResult.error,
        };
        if (onToolResult) onToolResult(individualResult);
        results.push(individualResult);
        const summary = generateChangeSummary(call.name, call.arguments, beforeTexts[i], individualResult.result);
        if (summary) changeSummaries.push(summary);
      }

      // Execute non-batchable calls individually
      for (const call of nonBatchableCalls) {
        if (onToolCall) onToolCall(call);
        const result = await executeTool(call, serverUrl);
        if (onToolResult) onToolResult(result);
        results.push(result);
        const summary = generateChangeSummary(call.name, call.arguments, null, result.result);
        if (summary) changeSummaries.push(summary);
      }
    } else {
      // Execute all calls individually with navigation
      for (const call of toolCalls) {
        if (onToolCall) onToolCall(call);

        // Navigate to paragraph before modifying it (visual feedback)
        if (MODIFYING_TOOL_NAMES.has(call.name)) {
          const idx = call.arguments.index ?? call.arguments.paragraphIndex;
          if (idx !== undefined && typeof idx === "number") {
            await navigateToParagraph(serverUrl, idx);
          }
        }

        // Before-fetch: snapshot text for diff summaries
        let beforeText: string | null = null;
        if (PARAGRAPH_EDIT_TOOLS.has(call.name)) {
          const idx = call.arguments.index ?? call.arguments.paragraphIndex;
          if (idx !== undefined && typeof idx === "number") {
            beforeText = await fetchParagraphText(serverUrl, idx);
          }
        }
        if (call.name === "editSelection" || call.name === "replaceSelection") {
          try {
            const selRes = await httpRequest(new URL("/api/selection", serverUrl).toString(), { method: "GET", headers: {} });
            let selRaw = "";
            for await (const ch of selRes) selRaw += ch.toString();
            const selData = JSON.parse(selRaw);
            beforeText = selData?.text || null;
          } catch {}
        }

        const result = await executeTool(call, serverUrl);
        if (onToolResult) onToolResult(result);
        results.push(result);

        // Generate change summary
        const summary = generateChangeSummary(call.name, call.arguments, beforeText, result.result);
        if (summary) changeSummaries.push(summary);
      }
    }

    // Add individual tool result messages (proper format for OpenAI/Anthropic)
    for (const r of results) {
      messages.push({
        role: "tool",
        tool_call_id: r.toolCallId,
        content: r.error ? `Error: ${r.error}` : JSON.stringify(r.result),
      });
    }

    // Continue the loop with no additional user message — the model
    // naturally continues after seeing tool results
    currentPrompt = "";
  }
}
