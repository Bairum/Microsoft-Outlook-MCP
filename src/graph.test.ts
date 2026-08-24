import { strict as assert } from "node:assert";
import { GraphClient, GraphError } from "./graph.js";
import type { AuthProvider } from "./auth.js";

/**
 * Unit tests for GraphClient security hardening (absolute URL handling).
 * Run with: node --test dist/graph.test.js
 */

// Mock AuthProvider for testing
class MockAuthProvider {
  async getAccessToken(): Promise<string> {
    return "mock-token";
  }

  async isSignedIn(): Promise<boolean> {
    return true;
  }

  async signOut(): Promise<void> {}

  async deviceCodeLogin(): Promise<string> {
    return "mock-token";
  }
}

// Test: GraphClient.request accepts valid graph.microsoft.com absolute URLs
async function testAcceptsValidAbsoluteUrl() {
  const mockAuth = new MockAuthProvider() as unknown as AuthProvider;
  const client = new GraphClient(mockAuth);

  // Mock fetch to verify the request
  const originalFetch = global.fetch;
  let fetchedUrl: string | undefined;

  try {
    global.fetch = async (input: string | URL | Request, init?: RequestInit) => {
      fetchedUrl = typeof input === "string" ? input : input.toString();
      
      // Verify the Authorization header is present
      assert.ok(init?.headers, "Expected headers in fetch call");
      const headers = init.headers as Record<string, string>;
      assert.strictEqual(headers.Authorization, "Bearer mock-token");
      
      return new Response(JSON.stringify({ value: [] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    };

    // Test with a valid graph.microsoft.com deltaLink
    await client.request({
      path: "https://graph.microsoft.com/v1.0/me/messages/delta?$deltatoken=abc123",
    });

    assert.strictEqual(
      fetchedUrl,
      "https://graph.microsoft.com/v1.0/me/messages/delta?$deltatoken=abc123",
      "Expected fetch to use the absolute URL directly"
    );

    console.log("✅ testAcceptsValidAbsoluteUrl passed");
  } finally {
    global.fetch = originalFetch;
  }
}

// Test: GraphClient.request rejects non-Graph absolute URLs
async function testRejectsNonGraphAbsoluteUrl() {
  const mockAuth = new MockAuthProvider() as unknown as AuthProvider;
  const client = new GraphClient(mockAuth);

  try {
    await client.request({
      path: "https://evil.com/api/messages?token=steal-me",
    });
    assert.fail("Expected request to throw for non-Graph origin");
  } catch (err) {
    assert.ok(err instanceof GraphError);
    assert.match(
      err.message,
      /origin must be https:\/\/graph\.microsoft\.com/i,
      "Expected origin validation error"
    );
  }

  console.log("✅ testRejectsNonGraphAbsoluteUrl passed");
}

// Test: GraphClient.request handles relative paths normally
async function testHandlesRelativePaths() {
  const mockAuth = new MockAuthProvider() as unknown as AuthProvider;
  const client = new GraphClient(mockAuth);

  const originalFetch = global.fetch;
  let fetchedUrl: string | undefined;

  try {
    global.fetch = async (input: string | URL | Request, init?: RequestInit) => {
      fetchedUrl = typeof input === "string" ? input : input.toString();
      
      return new Response(JSON.stringify({ id: "test-message" }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    };

    // Test with a relative path
    await client.request({
      path: "/me/messages",
      query: { $top: 10 },
    });

    assert.ok(
      fetchedUrl?.startsWith("https://graph.microsoft.com/v1.0/me/messages"),
      "Expected fetch to concatenate GRAPH_BASE with relative path"
    );
    assert.ok(
      fetchedUrl?.includes("%24top=10") || fetchedUrl?.includes("$top=10"),
      `Expected query parameters to be added, got: ${fetchedUrl}`
    );

    console.log("✅ testHandlesRelativePaths passed");
  } finally {
    global.fetch = originalFetch;
  }
}

// Test: GraphClient.request ignores query params for absolute URLs
async function testIgnoresQueryForAbsoluteUrls() {
  const mockAuth = new MockAuthProvider() as unknown as AuthProvider;
  const client = new GraphClient(mockAuth);

  const originalFetch = global.fetch;
  let fetchedUrl: string | undefined;

  try {
    global.fetch = async (input: string | URL | Request, init?: RequestInit) => {
      fetchedUrl = typeof input === "string" ? input : input.toString();
      
      return new Response(JSON.stringify({ value: [] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    };

    // Test with an absolute URL and query params
    // Query params should be ignored since the absolute URL already has them
    await client.request({
      path: "https://graph.microsoft.com/v1.0/me/messages/delta?$deltatoken=abc123",
      query: { $top: 999 }, // This should be ignored
    });

    assert.strictEqual(
      fetchedUrl,
      "https://graph.microsoft.com/v1.0/me/messages/delta?$deltatoken=abc123",
      "Expected fetch to use the absolute URL without adding query params"
    );
    assert.ok(
      !fetchedUrl?.includes("$top=999"),
      "Expected query params to be ignored for absolute URLs"
    );

    console.log("✅ testIgnoresQueryForAbsoluteUrls passed");
  } finally {
    global.fetch = originalFetch;
  }
}

// Run all tests
async function runTests() {
  console.log("Running GraphClient tests...\n");

  try {
    await testAcceptsValidAbsoluteUrl();
    await testRejectsNonGraphAbsoluteUrl();
    await testHandlesRelativePaths();
    await testIgnoresQueryForAbsoluteUrls();

    console.log("\n✅ All tests passed!");
  } catch (err) {
    console.error("\n❌ Test failed:", err);
    process.exit(1);
  }
}

// Run tests if this file is executed directly
if (import.meta.url === `file://${process.argv[1]}`) {
  runTests();
}
