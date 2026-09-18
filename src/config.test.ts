import { strict as assert } from "node:assert";
import { parseEnvAssignment } from "./config.js";

/**
 * Unit tests for .env / .keyring-env line parsing.
 * Run with: node dist/config.test.js
 */

function testParsesExportQuotedLines() {
  const parsed = parseEnvAssignment(
    "export DBUS_SESSION_BUS_ADDRESS='unix:path=/run/user/1000/bus'",
  );
  assert.ok(parsed);
  assert.equal(parsed.key, "DBUS_SESSION_BUS_ADDRESS");
  assert.equal(parsed.value, "unix:path=/run/user/1000/bus");
}

function testParsesBareAssignment() {
  const parsed = parseEnvAssignment('OUTLOOK_CLIENT_ID="abc-123"');
  assert.ok(parsed);
  assert.equal(parsed.key, "OUTLOOK_CLIENT_ID");
  assert.equal(parsed.value, "abc-123");
}

function testSkipsCommentsAndJunk() {
  assert.equal(parseEnvAssignment("# export FOO='bar'"), null);
  assert.equal(parseEnvAssignment(""), null);
  assert.equal(parseEnvAssignment("not-an-assignment"), null);
  assert.equal(parseEnvAssignment("export 1BAD=nope"), null);
}

function runTests() {
  testParsesExportQuotedLines();
  testParsesBareAssignment();
  testSkipsCommentsAndJunk();
  console.log("✅ config.env assignment tests passed");
}

if (import.meta.url === `file://${process.argv[1]}`) {
  runTests();
}
