/**
 * @module context
 * Context window management for conversation history.
 *
 * Dynamically queries model context sizes from API endpoints,
 * caches results, and manages conversation history to fit within
 * a configurable budget of the context window.
 */

import http from "http";
import https from "https";
import { URL } from "url";

// ── Defaults ──

/** Conservative fallback when we can't determine context size */
const DEFAULT_CONTEXT_SIZE = 32000;

/** Default percentage of context window to use for conversation history */
const DEFAULT_CONTEXT_BUDGET_PERCENT = 40;

// ── Token Estimation ──

/** Rough token estimate: ~4 chars per token */
export function estimateTokens(text: string): number {
  return Math.ceil(text.length / 4);
}

/** Estimate tokens for a message array */
export function estimateMessagesTokens(messages: { role: string; content: string }[]): number {
  let total = 0;
  for (const m of messages) {
    total += 4 + estimateTokens(m.content);
  }
  return total;
}

// ── Context Size Cache ──

interface CacheEntry {
  size: number;
  ts: number;
}

const contextSizeCache = new Map<string, CacheEntry>();
const CACHE_TTL = 10 * 60 * 1000; // 10 minutes

function getCached(key: string): number | null {
  const entry = contextSizeCache.get(key);
  if (entry && Date.now() - entry.ts < CACHE_TTL) return entry.size;
  return null;
}

function setCache(key: string, size: number): void {
  contextSizeCache.set(key, { size, ts: Date.now() });
}

// ── HTTP Helper ──

function fetchJSON(url: string, headers: Record<string, string> = {}, timeoutMs = 5000): Promise<any> {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const mod = parsed.protocol === "https:" ? https : http;
    const req = mod.request(url, {
      method: "GET",
      headers: { "Content-Type": "application/json", ...headers },
      rejectUnauthorized: false,
    }, (res) => {
      let data = "";
      res.on("data", (chunk: any) => data += chunk.toString());
      res.on("end", () => {
        try { resolve(JSON.parse(data)); }
        catch { reject(new Error("Invalid JSON")); }
      });
      res.on("error", reject);
    });
    req.on("error", reject);
    req.setTimeout(timeoutMs, () => { req.destroy(); reject(new Error("timeout")); });
    req.end();
  });
}

// ── Dynamic Context Size Queries ──

/**
 * Query an OpenAI-compatible /v1/models endpoint for a model's context size.
 * Works for OpenAI, local (LM Studio, Ollama, etc.), and any compatible API.
 */
async function queryOpenAICompatibleContextSize(
  baseUrl: string,
  modelId: string,
  apiKey?: string,
): Promise<number | null> {
  try {
    const url = baseUrl.replace(/\/$/, "") + "/models";
    const headers: Record<string, string> = {};
    if (apiKey) headers["Authorization"] = `Bearer ${apiKey}`;

    const body = await fetchJSON(url, headers);
    const models = body?.data || body;
    if (!Array.isArray(models)) return null;

    // Try exact match first, then partial
    for (const m of models) {
      if (m.id === modelId) {
        const ctx = m.context_window ?? m.context_length ?? m.max_model_len ?? m.max_tokens;
        if (typeof ctx === "number" && ctx > 0) return ctx;
      }
    }
    for (const m of models) {
      if (m.id?.includes(modelId) || modelId.includes(m.id)) {
        const ctx = m.context_window ?? m.context_length ?? m.max_model_len ?? m.max_tokens;
        if (typeof ctx === "number" && ctx > 0) return ctx;
      }
    }
  } catch {
    // silently fall through
  }
  return null;
}

/**
 * Query Anthropic's models endpoint for context size.
 * Uses /v1/models if available.
 */
async function queryAnthropicContextSize(
  modelId: string,
  apiKey: string,
): Promise<number | null> {
  try {
    const body = await fetchJSON("https://api.anthropic.com/v1/models", {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    });
    const models = body?.data || body;
    if (!Array.isArray(models)) return null;

    for (const m of models) {
      if (m.id === modelId) {
        const ctx = m.context_window ?? m.max_tokens ?? m.context_length;
        if (typeof ctx === "number" && ctx > 0) return ctx;
      }
    }
  } catch {
    // Anthropic models endpoint may not exist — fall through
  }
  return null;
}

// ── Public API ──

export interface ContextSizeConfig {
  openaiApiKey?: string;
  anthropicApiKey?: string;
  localBaseUrl?: string;
}

/**
 * Get context window size for a model, querying APIs dynamically.
 * Results are cached in memory for 10 minutes.
 *
 * @param backend - "openai" | "anthropic" | "local" | "openclaw"
 * @param modelId - The model identifier (e.g., "gpt-4o", "claude-sonnet-4-20250514")
 * @param config - API keys and local endpoint URL
 * @returns Context window size in tokens
 */
export async function getContextSize(
  backend: string,
  modelId: string,
  config: ContextSizeConfig = {},
): Promise<number> {
  const cacheKey = `${backend}:${modelId}`;
  const cached = getCached(cacheKey);
  if (cached !== null) return cached;

  let size: number | null = null;

  switch (backend) {
    case "openai":
      size = await queryOpenAICompatibleContextSize(
        "https://api.openai.com/v1",
        modelId,
        config.openaiApiKey,
      );
      break;

    case "anthropic":
      if (config.anthropicApiKey) {
        size = await queryAnthropicContextSize(modelId, config.anthropicApiKey);
      }
      break;

    case "local":
      if (config.localBaseUrl) {
        size = await queryOpenAICompatibleContextSize(config.localBaseUrl, modelId);
      }
      break;

    case "openclaw":
      // OpenClaw manages its own context — we don't need to manage it
      size = null;
      break;
  }

  const result = size ?? DEFAULT_CONTEXT_SIZE;
  setCache(cacheKey, result);
  return result;
}

// ── Context Management ──

export interface ManagedContext {
  /** Messages ready to send (may include a summary message at the start) */
  messages: { role: string; content: string }[];
  /** Number of original messages that were compacted into summary */
  compactedCount: number;
  /** Estimated tokens used by the returned messages */
  estimatedTokens: number;
}

/**
 * Manage conversation history to fit within the context budget.
 *
 * @param history - Full conversation history
 * @param contextSize - Model's total context window in tokens
 * @param documentContext - Current document text (to account for its token usage)
 * @param contextBudgetPercent - Percentage of context window for history (default 40)
 */
export function manageContext(
  history: { role: string; content: string }[],
  contextSize: number,
  documentContext?: string,
  contextBudgetPercent: number = DEFAULT_CONTEXT_BUDGET_PERCENT,
): ManagedContext {
  if (history.length === 0) {
    return { messages: [], compactedCount: 0, estimatedTokens: 0 };
  }

  const budgetTokens = Math.floor(contextSize * contextBudgetPercent / 100);
  const docTokens = documentContext ? estimateTokens(documentContext) : 0;
  const availableTokens = budgetTokens - docTokens;

  if (availableTokens <= 0) {
    const last = history.slice(-2);
    return {
      messages: last,
      compactedCount: history.length - last.length,
      estimatedTokens: estimateMessagesTokens(last),
    };
  }

  // Walk backwards, accumulating messages that fit
  let tokenBudget = availableTokens;
  let cutoff = history.length;

  for (let i = history.length - 1; i >= 0; i--) {
    const msgTokens = 4 + estimateTokens(history[i].content);
    if (tokenBudget - msgTokens < 0) break;
    tokenBudget -= msgTokens;
    cutoff = i;
  }

  if (cutoff === 0) {
    return {
      messages: [...history],
      compactedCount: 0,
      estimatedTokens: availableTokens - tokenBudget,
    };
  }

  const kept = history.slice(cutoff);
  const dropped = history.slice(0, cutoff);

  // v1: Simple truncation summary
  // Future: call LLM to generate a proper summary
  const summaryBudget = Math.floor(availableTokens * 0.15);
  const summary = buildSimpleSummary(dropped, summaryBudget);

  const result: { role: string; content: string }[] = [];
  if (summary) {
    result.push({ role: "system", content: summary });
  }
  result.push(...kept);

  return {
    messages: result,
    compactedCount: dropped.length,
    estimatedTokens: estimateMessagesTokens(result),
  };
}

/**
 * Build a simple summary of compacted messages by concatenating and truncating.
 * v1 approach — a future version could call the LLM to summarize.
 */
function buildSimpleSummary(
  messages: { role: string; content: string }[],
  maxTokens: number,
): string {
  if (messages.length === 0) return "";

  const header = "Earlier in this conversation, the user and assistant discussed:\n\n";
  const maxChars = (maxTokens - estimateTokens(header)) * 4;

  if (maxChars <= 100) {
    return header + `[${messages.length} earlier messages omitted for context space]`;
  }

  let summary = "";
  for (const msg of messages) {
    const prefix = msg.role === "user" ? "User: " : "Assistant: ";
    const line = prefix + msg.content.replace(/\n+/g, " ").slice(0, 500) + "\n";
    if (summary.length + line.length > maxChars) {
      summary += "...\n";
      break;
    }
    summary += line;
  }

  return header + summary.trim();
}
