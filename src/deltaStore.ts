import { PersistenceCreator, DataProtectionScope } from "@azure/msal-node-extensions";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

function log(msg: string): void {
  process.stderr.write(`[delta-store] ${msg}\n`);
}

/**
 * Secure storage for Graph API delta cursors. Uses the same OS-native encrypted
 * backend as the token cache (Windows DPAPI / macOS Keychain / Linux libsecret).
 * Fails closed if native storage is unavailable — no plaintext fallback.
 */
export class DeltaStore {
  private persistence: any;
  private data: Map<string, string> = new Map();
  private loaded = false;

  private constructor(persistence: any) {
    this.persistence = persistence;
  }

  /**
   * Create and verify a DeltaStore with OS-native encrypted storage.
   * Fails if the native backend is unavailable.
   */
  static async create(cachePath: string): Promise<DeltaStore> {
    const testDir = mkdtempSync(join(tmpdir(), "mcp-outlook-delta-verify-"));
    const testCachePath = join(testDir, "test-verify.cache");

    try {
      const testPersistence = await PersistenceCreator.createPersistence({
        cachePath: testCachePath,
        dataProtectionScope: DataProtectionScope.CurrentUser,
        serviceName: "microsoft-outlook-mcp",
        accountName: "delta-cursors",
        usePlaintextFileOnLinux: false,
      });

      const okToUse = await testPersistence.verifyPersistence();
      if (!okToUse) {
        throw new Error(
          "verifyPersistence() returned false — OS-native encrypted storage is not available"
        );
      }

      rmSync(testDir, { recursive: true, force: true });

      const persistence = await PersistenceCreator.createPersistence({
        cachePath,
        dataProtectionScope: DataProtectionScope.CurrentUser,
        serviceName: "microsoft-outlook-mcp",
        accountName: "delta-cursors",
        usePlaintextFileOnLinux: false,
      });

      log("delta cursor store: OS-native encrypted storage (msal-node-extensions) — verified");
      return new DeltaStore(persistence);
    } catch (err) {
      rmSync(testDir, { recursive: true, force: true });
      throw new Error(
        `OS-native encrypted delta cursor storage is required but unavailable: ${(err as Error).message}. ` +
        `This hardened fork refuses to fall back to plaintext. ` +
        `See README troubleshooting for instructions to enable encrypted storage on your OS.`
      );
    }
  }

  /** Load cursors from encrypted storage. */
  private async load(): Promise<void> {
    if (this.loaded) return;
    try {
      const raw = await this.persistence.load();
      if (raw) {
        const parsed = JSON.parse(raw);
        this.data = new Map(Object.entries(parsed));
      }
    } catch {
      // Empty or corrupt store — start fresh.
      this.data = new Map();
    }
    this.loaded = true;
  }

  /** Save cursors to encrypted storage. */
  private async save(): Promise<void> {
    const obj = Object.fromEntries(this.data);
    await this.persistence.save(JSON.stringify(obj, null, 2));
  }

  /**
   * Get a deltaLink cursor by key. Returns undefined if not found.
   * Security: validates that the stored URL points to graph.microsoft.com.
   */
  async get(key: string): Promise<string | undefined> {
    await this.load();
    const cursor = this.data.get(key);
    if (!cursor) return undefined;

    // Validate stored cursor before returning it.
    try {
      const url = new URL(cursor);
      if (url.origin !== "https://graph.microsoft.com") {
        log(`stored cursor for ${key} has invalid origin: ${url.origin}; discarding`);
        this.data.delete(key);
        await this.save();
        return undefined;
      }
    } catch {
      log(`stored cursor for ${key} is not a valid URL; discarding`);
      this.data.delete(key);
      await this.save();
      return undefined;
    }

    return cursor;
  }

  /**
   * Set a deltaLink cursor by key.
   * Security: validates that the URL points to graph.microsoft.com before storing.
   */
  async set(key: string, cursor: string): Promise<void> {
    await this.load();

    // Validate cursor URL before storing it.
    try {
      const url = new URL(cursor);
      if (url.origin !== "https://graph.microsoft.com") {
        throw new Error(
          `deltaLink origin must be https://graph.microsoft.com, got ${url.origin}`
        );
      }
    } catch (err) {
      throw new Error(`Invalid deltaLink URL: ${(err as Error).message}`);
    }

    this.data.set(key, cursor);
    await this.save();
  }

  /** Clear a cursor by key. */
  async delete(key: string): Promise<void> {
    await this.load();
    this.data.delete(key);
    await this.save();
  }

  /** Clear all cursors. */
  async clear(): Promise<void> {
    await this.load();
    this.data.clear();
    await this.save();
  }
}
