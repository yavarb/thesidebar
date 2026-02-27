/**
 * @module settings.test
 * Tests for settings CRUD, masking, and merge logic.
 */

import fs from "fs";
import os from "os";
import path from "path";
import { readConfig, writeConfig, maskConfig, mergeConfig, SidebarConfig } from "./settings";

let passed = 0;
let failed = 0;

function assert(cond: boolean, msg: string) {
  if (cond) { console.log(`  ✅ ${msg}`); passed++; }
  else { console.log(`  ❌ ${msg}`); failed++; }
}

// Use a temp directory for tests
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "wr-settings-test-"));
const testConfigPath = path.join(tmpDir, "config.json");

function cleanup() {
  try { fs.rmSync(tmpDir, { recursive: true }); } catch {}
}

console.log("🎀 Settings Test Suite\n══════════════════════");

// Test readConfig with missing file
console.log("\n── readConfig ──");
const empty = readConfig(testConfigPath);
assert(Object.keys(empty).length === 0, "returns empty config when file missing");

// Test writeConfig + readConfig roundtrip
console.log("\n── writeConfig + readConfig ──");
const testConfig: SidebarConfig = {
  openclawUrl: "https://localhost:3001",
  anthropicApiKey: "sk-ant-api03-abcdef1234567890",
  openaiApiKey: "sk-proj-abcdef1234567890abcdef",
  localEndpoints: [{ name: "LM Studio", baseUrl: "http://10.0.0.167:1234/v1" }],
  defaultModel: "openai:gpt-4o",
};
writeConfig(testConfig, testConfigPath);

const readBack = readConfig(testConfigPath);
assert(readBack.openclawUrl === "https://localhost:3001", "preserves openclawUrl");
assert(readBack.anthropicApiKey === "sk-ant-api03-abcdef1234567890", "preserves anthropicApiKey");
assert(readBack.openaiApiKey === "sk-proj-abcdef1234567890abcdef", "preserves openaiApiKey");
assert(readBack.localEndpoints?.length === 1, "preserves localEndpoints");
assert(readBack.localEndpoints?.[0].name === "LM Studio", "preserves endpoint name");
assert(readBack.defaultModel === "openai:gpt-4o", "preserves defaultModel");

// Test file permissions (unix only)
if (process.platform !== "win32") {
  const stats = fs.statSync(testConfigPath);
  const mode = (stats.mode & 0o777).toString(8);
  assert(mode === "600", `file permissions are 600 (got ${mode})`);
}

// Test maskConfig
console.log("\n── maskConfig ──");
const masked = maskConfig(testConfig);
assert(masked.openclawUrl === "https://localhost:3001", "non-sensitive fields unmasked");
assert(masked.anthropicApiKey !== testConfig.anthropicApiKey, "anthropicApiKey is masked");
assert(masked.anthropicApiKey!.includes("****"), "masked key contains ****");
assert(masked.anthropicApiKey!.endsWith("7890"), "masked key shows last 4 chars");
assert(masked.openaiApiKey!.includes("****"), "openaiApiKey is masked");
assert(masked.localEndpoints?.length === 1, "localEndpoints preserved in masked output");

// Test masking short keys
const shortKeyConfig: SidebarConfig = { anthropicApiKey: "abc" };
const shortMasked = maskConfig(shortKeyConfig);
assert(shortMasked.anthropicApiKey === "****", "short keys fully masked");

// Test masking undefined keys
const noKeyConfig: SidebarConfig = {};
const noKeyMasked = maskConfig(noKeyConfig);
assert(noKeyMasked.anthropicApiKey === undefined, "undefined keys stay undefined");

// Test mergeConfig
console.log("\n── mergeConfig ──");
const base: SidebarConfig = {
  openclawUrl: "https://localhost:3001",
  anthropicApiKey: "sk-real-key-12345678",
  openaiApiKey: "sk-proj-real-key-abcd",
};

// Update only openclawUrl
const m1 = mergeConfig(base, { openclawUrl: "https://newhost:3002" });
assert(m1.openclawUrl === "https://newhost:3002", "updates provided field");
assert(m1.anthropicApiKey === "sk-real-key-12345678", "preserves unchanged fields");

// Skip masked values
const m2 = mergeConfig(base, { anthropicApiKey: "sk-****5678" });
assert(m2.anthropicApiKey === "sk-real-key-12345678", "skips masked values (doesn't overwrite real key)");

// Update with real new key
const m3 = mergeConfig(base, { anthropicApiKey: "sk-brand-new-key-9999" });
assert(m3.anthropicApiKey === "sk-brand-new-key-9999", "accepts real new key");

// Add local endpoints
const m4 = mergeConfig(base, { localEndpoints: [{ name: "New", baseUrl: "http://localhost:5000/v1" }] });
assert(m4.localEndpoints?.length === 1, "adds localEndpoints");

// Test Express handlers (lightweight)
console.log("\n── Route Handlers ──");
import { handleGetSettings, handlePostSettings } from "./settings";

// Mock Express req/res
function mockRes() {
  let statusCode = 200;
  let body: any = null;
  return {
    status(code: number) { statusCode = code; return this; },
    json(data: any) { body = data; },
    get statusCode() { return statusCode; },
    get body() { return body; },
  };
}

const getHandler = handleGetSettings(testConfigPath);
const res1 = mockRes();
getHandler({}, res1);
assert(res1.body.ok === true, "GET returns ok:true");
assert(res1.body.data.anthropicApiKey?.includes("****"), "GET masks sensitive fields");
assert(res1.body.data.openclawUrl === "https://localhost:3001", "GET returns non-sensitive fields");

const postHandler = handlePostSettings(testConfigPath);
const res2 = mockRes();
postHandler({ body: { defaultModel: "anthropic:claude-sonnet-4-20250514" } }, res2);
assert(res2.body.ok === true, "POST returns ok:true");
assert(res2.body.data.defaultModel === "anthropic:claude-sonnet-4-20250514", "POST updates field");

// Verify it persisted
const afterPost = readConfig(testConfigPath);
assert(afterPost.defaultModel === "anthropic:claude-sonnet-4-20250514", "POST persists to disk");
assert(afterPost.anthropicApiKey === "sk-ant-api03-abcdef1234567890", "POST preserves existing keys");

// Bad request
const res3 = mockRes();
postHandler({ body: null }, res3);
assert(res3.statusCode === 400, "POST rejects null body");

cleanup();

console.log(`\n══════════════════════`);
console.log(`Results: ${passed} passed, ${failed} failed`);
process.exit(failed > 0 ? 1 : 0);
