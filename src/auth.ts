import {
  PublicClientApplication,
  LogLevel,
  type Configuration,
  type ICachePlugin,
} from "@azure/msal-node";
import {
  PersistenceCreator,
  PersistenceCachePlugin,
  DataProtectionScope,
  type IPersistence,
} from "@azure/msal-node-extensions";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { Config } from "./config.js";

/**
 * IMPORTANT: everything here logs to stderr. In an MCP stdio server, stdout is
 * reserved exclusively for the JSON-RPC protocol stream.
 */
function log(msg: string): void {
  process.stderr.write(`[auth] ${msg}\n`);
}

/**
 * Live libsecret / Keychain keys for the MSAL token cache.
 * On Linux, libsecret keys by serviceName + accountName — not by cachePath.
 */
export const TOKEN_CACHE_SERVICE = "microsoft-outlook-mcp";
export const TOKEN_CACHE_ACCOUNT = "token-cache";

/**
 * Startup verify-probe keys. MUST differ from the live cache pair above.
 * Reusing microsoft-outlook-mcp / token-cache lets a probe write clobber the
 * real token on Linux (same Secret Service key, different cachePath is ignored).
 */
export const VERIFY_SERVICE = "microsoft-outlook-mcp-verify";
export const VERIFY_ACCOUNT = "verify-probe";

/** Empty MSAL cache JSON is ~86 chars; a real session is much larger. */
export const MIN_PERSISTED_CACHE_CHARS = 128;

export function isTinyTokenCache(serialized: string | null | undefined): boolean {
  return !serialized || serialized.trim().length < MIN_PERSISTED_CACHE_CHARS;
}

/**
 * Enforced cache: OS-native encrypted storage via @azure/msal-node-extensions.
 *   - Windows -> DPAPI (encrypted, tied to the current Windows user)
 *   - macOS   -> Keychain
 *   - Linux   -> libsecret / GNOME Keyring
 * Fails closed if the native backend is unavailable. No plaintext fallback.
 */
async function buildCachePlugin(
  config: Config,
): Promise<{ plugin: ICachePlugin; persistence: IPersistence }> {
  // Verification must not leave dummy files in the project root. Use a temp
  // directory for the persistence verify test and clean it up after.
  const testDir = mkdtempSync(join(tmpdir(), "mcp-outlook-verify-"));
  const testCachePath = join(testDir, "test-verify.cache");

  try {
    const testPersistence = await PersistenceCreator.createPersistence({
      cachePath: testCachePath,
      dataProtectionScope: DataProtectionScope.CurrentUser,
      serviceName: VERIFY_SERVICE,
      accountName: VERIFY_ACCOUNT,
      usePlaintextFileOnLinux: false,
    });

    const okToUse = await testPersistence.verifyPersistence();
    if (!okToUse) {
      throw new Error(
        "verifyPersistence() returned false — OS-native encrypted storage is not available",
      );
    }

    // Verification passed. Clean up test dir and create the real persistence.
    rmSync(testDir, { recursive: true, force: true });

    const persistence = await PersistenceCreator.createPersistence({
      cachePath: config.tokenCachePath,
      dataProtectionScope: DataProtectionScope.CurrentUser,
      serviceName: TOKEN_CACHE_SERVICE,
      accountName: TOKEN_CACHE_ACCOUNT,
      usePlaintextFileOnLinux: false,
    });

    log("token cache: OS-native encrypted storage (msal-node-extensions) — verified");
    return { plugin: new PersistenceCachePlugin(persistence), persistence };
  } catch (err) {
    rmSync(testDir, { recursive: true, force: true });
    throw new Error(
      `OS-native encrypted token storage is required but unavailable: ${(err as Error).message}. ` +
        `This hardened fork refuses to fall back to plaintext. ` +
        `See README troubleshooting for instructions to enable encrypted storage on your OS.`,
    );
  }
}

export class AuthProvider {
  private readonly pca: PublicClientApplication;
  private readonly scopes: string[];
  private readonly expectedUsername?: string;
  /** Live IPersistence — used to force-persist after login when the cache plugin skips a write. */
  private readonly persistence: IPersistence;

  private constructor(
    pca: PublicClientApplication,
    scopes: string[],
    persistence: IPersistence,
    expectedUsername?: string,
  ) {
    this.pca = pca;
    this.scopes = scopes;
    this.persistence = persistence;
    this.expectedUsername = expectedUsername;
  }

  /**
   * Async factory — building the encrypted persistence backend is async, so
   * construction must be awaited. Use this instead of `new AuthProvider(...)`.
   */
  static async create(config: Config): Promise<AuthProvider> {
    const { plugin, persistence } = await buildCachePlugin(config);

    const msalConfig: Configuration = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
      cache: { cachePlugin: plugin },
      system: {
        loggerOptions: {
          loggerCallback: (level, message) => {
            if (level <= LogLevel.Warning) log(message);
          },
          piiLoggingEnabled: false,
          logLevel: LogLevel.Warning,
        },
      },
    };

    return new AuthProvider(
      new PublicClientApplication(msalConfig),
      config.scopes,
      persistence,
      config.expectedUsername,
    );
  }

  /**
   * Returns a valid access token, refreshing silently from the cache when
   * possible. Only falls back to interactive device-code sign-in when there is
   * no usable cached account. When `interactive` is false (the default for
   * request-time acquisition) a missing/expired session throws instead of
   * blocking on user input — that keeps tool calls from hanging.
   *
   * Security: when expectedUsername is configured, only use an account that
   * matches it. Never silently pick the first cached account.
   */
  async getAccessToken(interactive = false): Promise<string> {
    const cache = this.pca.getTokenCache();
    const accounts = await cache.getAllAccounts();

    if (accounts.length > 0) {
      let targetAccount = accounts[0];

      // If a specific username is expected, find the matching account.
      if (this.expectedUsername) {
        const normalized = this.expectedUsername.toLowerCase();
        const match = accounts.find(
          (acc) => acc.username?.toLowerCase() === normalized,
        );
        if (!match) {
          const cached = accounts.map((a) => a.username).join(", ");
          throw new Error(
            `Expected account ${this.expectedUsername} is not in the token cache. ` +
              `Cached accounts: ${cached}. Run \`npm run login\` as ${this.expectedUsername}.`,
          );
        }
        targetAccount = match;
      }

      try {
        const result = await this.pca.acquireTokenSilent({
          account: targetAccount,
          scopes: this.scopes,
        });
        if (result?.accessToken) {
          await this.persistTokenCacheBestEffort("silent refresh");
          return result.accessToken;
        }
      } catch (err) {
        log(`silent token acquisition failed: ${(err as Error).message}`);
      }
    }

    if (!interactive) {
      throw new Error(
        "Not signed in (or the session expired). Run the login step first: " +
          "`npm run login`.",
      );
    }

    return this.deviceCodeLogin();
  }

  /** Interactive device-code sign-in. Prompts on stderr. */
  async deviceCodeLogin(): Promise<string> {
    const result = await this.pca.acquireTokenByDeviceCode({
      scopes: this.scopes,
      deviceCodeCallback: (info) => {
        // info.message already contains the URL + code + instructions.
        process.stderr.write(`\n${info.message}\n\n`);
      },
    });
    if (!result?.accessToken) {
      throw new Error("Device code sign-in returned no access token.");
    }
    log(`signed in as ${result.account?.username ?? "unknown account"}`);
    // PersistenceCachePlugin sometimes skips afterCacheAccess after device-code
    // login even when in-memory accounts exist. Force-write and round-trip.
    await this.forcePersistTokenCache("device-code login");
    return result.accessToken;
  }

  async isSignedIn(): Promise<boolean> {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    return accounts.length > 0;
  }

  async signOut(): Promise<void> {
    const cache = this.pca.getTokenCache();
    for (const account of await cache.getAllAccounts()) {
      await cache.removeAccount(account);
    }
    log("signed out; cached accounts cleared.");
  }

  /**
   * Write the in-memory MSAL cache through the live IPersistence and confirm
   * the Secret Service / DPAPI / Keychain round-trip is non-empty.
   */
  private async forcePersistTokenCache(reason: string): Promise<void> {
    const serialized = this.pca.getTokenCache().serialize();
    if (isTinyTokenCache(serialized)) {
      throw new Error(
        `Token cache serialize after ${reason} was empty/tiny (${serialized?.length ?? 0} chars). ` +
          `Login succeeded in memory but nothing durable was available to write.`,
      );
    }
    await this.persistence.save(serialized);
    const loaded = await this.persistence.load();
    if (isTinyTokenCache(loaded)) {
      throw new Error(
        `Token cache failed to persist after ${reason}: reload was empty/tiny (${loaded?.length ?? 0} chars). ` +
          `On Linux, confirm Secret Service is running and the keyring is unlocked (see README troubleshooting).`,
      );
    }
    log(`force-persisted token cache after ${reason} (${loaded!.length} chars)`);
  }

  private async persistTokenCacheBestEffort(reason: string): Promise<void> {
    try {
      await this.forcePersistTokenCache(reason);
    } catch (err) {
      log(`best-effort persist after ${reason} skipped: ${(err as Error).message}`);
    }
  }
}
