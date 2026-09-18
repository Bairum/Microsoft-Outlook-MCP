import { strict as assert } from "node:assert";
import {
  TOKEN_CACHE_ACCOUNT,
  TOKEN_CACHE_SERVICE,
  VERIFY_ACCOUNT,
  VERIFY_SERVICE,
  isTinyTokenCache,
  MIN_PERSISTED_CACHE_CHARS,
} from "./auth.js";

/**
 * Unit tests for Linux token-cache persistence helpers.
 * Run with: node dist/auth.test.js
 */

function testVerifyKeysDifferFromLiveCache() {
  assert.notEqual(
    `${VERIFY_SERVICE}/${VERIFY_ACCOUNT}`,
    `${TOKEN_CACHE_SERVICE}/${TOKEN_CACHE_ACCOUNT}`,
    "verify probe must not reuse the live libsecret service+account pair",
  );
  assert.equal(TOKEN_CACHE_SERVICE, "microsoft-outlook-mcp");
  assert.equal(TOKEN_CACHE_ACCOUNT, "token-cache");
  assert.equal(VERIFY_SERVICE, "microsoft-outlook-mcp-verify");
  assert.equal(VERIFY_ACCOUNT, "verify-probe");
}

function testTinyCacheDetection() {
  assert.equal(isTinyTokenCache(null), true);
  assert.equal(isTinyTokenCache(""), true);
  assert.equal(isTinyTokenCache("   "), true);
  const emptyMsal =
    '{"Account":{},"IdToken":{},"AccessToken":{},"RefreshToken":{},"AppMetadata":{}}';
  assert.ok(emptyMsal.length < MIN_PERSISTED_CACHE_CHARS);
  assert.equal(isTinyTokenCache(emptyMsal), true);
  assert.equal(isTinyTokenCache("x".repeat(MIN_PERSISTED_CACHE_CHARS)), false);
}

function runTests() {
  testVerifyKeysDifferFromLiveCache();
  testTinyCacheDetection();
  console.log("✅ auth persistence helper tests passed");
}

if (import.meta.url === `file://${process.argv[1]}`) {
  runTests();
}
