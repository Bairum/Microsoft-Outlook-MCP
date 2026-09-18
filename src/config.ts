import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { existsSync, readFileSync } from "node:fs";

/**
 * Parse one `.env` / `.keyring-env` assignment. Accepts `KEY=value`,
 * `export KEY='value'`, and double-quoted values. Comments and blank lines
 * return null.
 */
export function parseEnvAssignment(
  rawLine: string,
): { key: string; value: string } | null {
  const line = rawLine.trim();
  if (!line || line.startsWith("#")) return null;
  const assignment = line.startsWith("export ") ? line.slice(7).trim() : line;
  const eq = assignment.indexOf("=");
  if (eq === -1) return null;
  const key = assignment.slice(0, eq).trim();
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) return null;
  let value = assignment.slice(eq + 1).trim();
  if (
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"))
  ) {
    value = value.slice(1, -1);
  }
  return { key, value };
}

function applyEnvFile(fileName: string): void {
  const here = dirname(fileURLToPath(import.meta.url));
  // dist/config.js -> project root is one level up from dist
  const root = join(here, "..");
  const envPath = join(root, fileName);
  if (!existsSync(envPath)) return;

  const text = readFileSync(envPath, "utf8");
  for (const rawLine of text.split(/\r?\n/)) {
    const parsed = parseEnvAssignment(rawLine);
    if (!parsed) continue;
    if (process.env[parsed.key] === undefined) {
      process.env[parsed.key] = parsed.value;
    }
  }
}

/**
 * Minimal .env loader (no dependency). Reads KEY=VALUE lines from a .env file
 * in the project root and populates process.env for any keys not already set.
 */
function loadDotEnv(): void {
  applyEnvFile(".env");
}

/**
 * Headless Linux sessions often need D-Bus / keyring variables that a desktop
 * session would have exported already. `.keyring-env` accepts `export KEY='value'`
 * lines (git-ignored). Existing process.env keys are not overwritten.
 */
function loadKeyringEnv(): void {
  applyEnvFile(".keyring-env");
}

loadDotEnv();
loadKeyringEnv();

function projectRoot(): string {
  const here = dirname(fileURLToPath(import.meta.url));
  return join(here, "..");
}

export interface Config {
  clientId: string;
  tenantId: string;
  scopes: string[];
  tokenCachePath: string;
  expectedUsername?: string;
  allowWrites: boolean;
}

/**
 * Hard-coded allowlist of Graph scopes for a hardened work-tenant fork.
 * MSAL may add offline_access automatically.
 */
const ALLOWED_SCOPES = new Set([
  "User.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "Calendars.ReadWrite",
  "Contacts.ReadWrite",
  "MailboxSettings.ReadWrite",
  "offline_access",
]);

const DEFAULT_SCOPES =
  "User.Read Mail.ReadWrite Mail.Send Calendars.ReadWrite Contacts.ReadWrite MailboxSettings.ReadWrite";

export function loadConfig(): Config {
  const clientId = process.env.OUTLOOK_CLIENT_ID?.trim();
  if (!clientId) {
    throw new Error(
      "OUTLOOK_CLIENT_ID is not set. Copy .env.example to .env and set your " +
        "Azure AD application (client) ID.",
    );
  }

  const tenantId = process.env.OUTLOOK_TENANT_ID?.trim();
  if (!tenantId) {
    throw new Error(
      "OUTLOOK_TENANT_ID is required. This hardened fork does not accept the default 'common' tenant. " +
        "Set OUTLOOK_TENANT_ID to your organization's tenant GUID in .env.",
    );
  }

  // Reject well-known tenant identifiers that allow any account type.
  const forbidden = ["common", "organizations", "consumers"];
  if (forbidden.includes(tenantId.toLowerCase())) {
    throw new Error(
      `OUTLOOK_TENANT_ID="${tenantId}" is not allowed in this hardened work-tenant fork. ` +
        `Use your organization's tenant GUID (e.g. a UUID like 12345678-1234-1234-1234-123456789abc).`,
    );
  }

  // Basic validation that it looks like a GUID.
  const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!guidRegex.test(tenantId)) {
    throw new Error(
      `OUTLOOK_TENANT_ID="${tenantId}" does not look like a valid tenant GUID. ` +
        `Expected format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`,
    );
  }

  const scopesString = process.env.OUTLOOK_SCOPES?.trim() || DEFAULT_SCOPES;
  const scopes = scopesString.split(/\s+/).filter(Boolean);

  // Validate scopes against allowlist.
  for (const scope of scopes) {
    if (!ALLOWED_SCOPES.has(scope)) {
      throw new Error(
        `Scope "${scope}" is not allowed in this hardened fork. ` +
          `Allowed scopes: ${Array.from(ALLOWED_SCOPES).filter((s) => s !== "offline_access").join(", ")}. ` +
          `Remove or replace disallowed scopes in OUTLOOK_SCOPES.`,
      );
    }
  }

  const tokenCachePath =
    process.env.OUTLOOK_TOKEN_CACHE_PATH?.trim() ||
    join(projectRoot(), ".token-cache.json");

  const expectedUsername = process.env.OUTLOOK_EXPECTED_USERNAME?.trim();

  const allowWrites =
    process.env.OUTLOOK_ALLOW_WRITES?.trim().toLowerCase() === "true";

  return { clientId, tenantId, scopes, tokenCachePath, expectedUsername, allowWrites };
}
