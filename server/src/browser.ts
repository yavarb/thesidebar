/**
 * @module browser
 * Web research tools for The Sidebar agent loop.
 * webSearch: search the web and return results
 * webFetch:  fetch a URL and extract readable text
 *
 * No external browser / Playwright needed — uses Node built-in fetch.
 */

import { JSDOM } from "jsdom";
import { Readability } from "@mozilla/readability";
import { readConfig } from "./settings";

// ── Types ──────────────────────────────────────────────────────────────────

export interface SearchResult {
  title: string;
  url: string;
  snippet: string;
}

export interface SearchResponse {
  query: string;
  results: SearchResult[];
  source: string;
}

export interface FetchResponse {
  url: string;
  title: string;
  content: string;
  wordCount: number;
  error?: string;
}

// ── Helpers ────────────────────────────────────────────────────────────────

const USER_AGENT =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 " +
  "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36";

async function fetchText(url: string, headers?: Record<string, string>): Promise<string> {
  const res = await fetch(url, {
    headers: { "User-Agent": USER_AGENT, ...headers },
    signal: AbortSignal.timeout(12000),
  });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return res.text();
}

// ── webSearch ──────────────────────────────────────────────────────────────

/** Search via Brave API if key configured, otherwise DuckDuckGo HTML fallback */
export async function webSearch(
  query: string,
  count: number = 5
): Promise<SearchResponse> {
  const config = readConfig();
  const braveKey = (config as any).braveApiKey as string | undefined;

  if (braveKey) {
    return searchViaBrave(query, count, braveKey);
  }
  return searchViaDuckDuckGo(query, count);
}

async function searchViaBrave(
  query: string,
  count: number,
  apiKey: string
): Promise<SearchResponse> {
  const url = `https://api.search.brave.com/res/v1/web/search?q=${encodeURIComponent(query)}&count=${Math.min(count, 10)}`;
  const raw = await fetchText(url, {
    Accept: "application/json",
    "Accept-Encoding": "gzip",
    "X-Subscription-Token": apiKey,
  });
  const data = JSON.parse(raw);
  const results: SearchResult[] = (data.web?.results || []).slice(0, count).map((r: any) => ({
    title: r.title || "",
    url: r.url || "",
    snippet: r.description || "",
  }));
  return { query, results, source: "brave" };
}

async function searchViaDuckDuckGo(
  query: string,
  count: number
): Promise<SearchResponse> {
  const url = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`;
  const html = await fetchText(url);
  const dom = new JSDOM(html);
  const doc = dom.window.document;

  const results: SearchResult[] = [];
  const links = doc.querySelectorAll("a.result__a");
  const snippets = doc.querySelectorAll(".result__snippet");

  links.forEach((link: Element, i: number) => {
    if (results.length >= count) return;
    const href = (link as HTMLAnchorElement).href || "";
    // DDG wraps URLs in redirects — extract from href or data attrs
    const realUrl = decodeURIComponent(href.replace(/.*?uddg=/, "").split("&")[0]) || href;
    results.push({
      title: link.textContent?.trim() || "",
      url: realUrl,
      snippet: (snippets[i] as HTMLElement)?.textContent?.trim() || "",
    });
  });

  return { query, results, source: "duckduckgo" };
}

// ── webFetch ───────────────────────────────────────────────────────────────

/** Fetch a URL and extract clean readable text via Readability */
export async function webFetch(url: string): Promise<FetchResponse> {
  try {
    const html = await fetchText(url);
    const dom = new JSDOM(html, { url });
    const reader = new Readability(dom.window.document);
    const article = reader.parse();

    const title = article?.title || dom.window.document.title || url;
    const content = article?.textContent?.replace(/\s+/g, " ").trim() || "";
    const wordCount = content.split(/\s+/).filter(Boolean).length;

    return { url, title, content: content.slice(0, 20000), wordCount };
  } catch (err: any) {
    return {
      url,
      title: "",
      content: "",
      wordCount: 0,
      error: err.message || String(err),
    };
  }
}
