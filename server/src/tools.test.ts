/**
 * @module tools.test
 * Validates tool definitions and endpoint mappings.
 */

import { TOOL_DEFINITIONS, TOOL_ENDPOINTS, getToolDefinitions, getToolEndpoint } from "./tools";

let passed = 0;
let failed = 0;

function assert(cond: boolean, msg: string) {
  if (cond) { console.log(`  ✅ ${msg}`); passed++; }
  else { console.log(`  ❌ ${msg}`); failed++; }
}

console.log("🎀 Tool Definitions Test Suite\n═══════════════════════════════");

// Validate all tool definitions
console.log("\n── Schema Validation ──");
const expectedTools = [
  "readDocument", "readParagraphs", "readParagraph", "replaceParagraph",
  "insertText", "findReplace", "readFootnotes", "addFootnote",
  "formatParagraph", "getStyles", "getDocumentStats", "getStructure",
  "readSelection", "navigateTo",
];

assert(TOOL_DEFINITIONS.length === expectedTools.length, `has ${expectedTools.length} tool definitions`);

for (const tool of TOOL_DEFINITIONS) {
  assert(tool.type === "function", `${tool.function.name}: type is "function"`);
  assert(typeof tool.function.name === "string" && tool.function.name.length > 0, `${tool.function.name}: has name`);
  assert(typeof tool.function.description === "string" && tool.function.description.length > 0, `${tool.function.name}: has description`);
  assert(tool.function.parameters.type === "object", `${tool.function.name}: parameters type is "object"`);
  assert(typeof tool.function.parameters.properties === "object", `${tool.function.name}: has properties object`);
}

// Verify all expected tools exist
console.log("\n── Expected Tools ──");
for (const name of expectedTools) {
  assert(TOOL_DEFINITIONS.some(t => t.function.name === name), `${name} defined`);
  assert(name in TOOL_ENDPOINTS, `${name} has endpoint mapping`);
}

// Validate endpoint mappings
console.log("\n── Endpoint Mappings ──");
for (const [name, ep] of Object.entries(TOOL_ENDPOINTS)) {
  assert(["GET", "POST", "PUT"].includes(ep.method), `${name}: valid HTTP method (${ep.method})`);
  assert(ep.path.startsWith("/api/"), `${name}: path starts with /api/`);
}

// Test mapArgs for tools that have them
console.log("\n── mapArgs ──");
const readParas = TOOL_ENDPOINTS.readParagraphs.mapArgs!({ from: 0, to: 5, compact: true });
assert(readParas.query?.from === "0" && readParas.query?.to === "5" && readParas.query?.compact === "true", "readParagraphs mapArgs");

const readPara = TOOL_ENDPOINTS.readParagraph.mapArgs!({ index: 3 });
assert(readPara.path === "/api/paragraph/3", "readParagraph mapArgs");

const nav = TOOL_ENDPOINTS.navigateTo.mapArgs!({ index: 10 });
assert(nav.body?.index === 10, "navigateTo mapArgs");

// Test getToolDefinitions filter
console.log("\n── Filtering ──");
const all = getToolDefinitions();
assert(all.length === expectedTools.length, "getToolDefinitions() returns all");

const filtered = getToolDefinitions(["readDocument", "getStyles"]);
assert(filtered.length === 2, "getToolDefinitions([...]) filters correctly");

const ep = getToolEndpoint("readDocument");
assert(ep !== undefined && ep.method === "GET", "getToolEndpoint returns correct endpoint");

assert(getToolEndpoint("nonexistent") === undefined, "getToolEndpoint returns undefined for unknown");

console.log(`\n═══════════════════════════════`);
console.log(`Results: ${passed} passed, ${failed} failed`);
process.exit(failed > 0 ? 1 : 0);
