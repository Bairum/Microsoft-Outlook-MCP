import type { AuthProvider } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Sanitize a Graph API path segment (message ID, folder ID, contact ID, event
 * ID, or folder display name). Rejects `/` and `..` to prevent path traversal.
 * Encodes the segment so interpolation into a path is safe.
 */
export function sanitizePathSegment(segment: string, label = "id"): string {
  if (!segment || typeof segment !== "string") {
    throw new Error(`${label} must be a non-empty string`);
  }
  if (segment.includes("/") || segment.includes("..")) {
    throw new Error(
      `${label} contains forbidden path characters (/ or ..) that could ` +
        `escape /me/ scope: ${segment}`,
    );
  }
  return encodeURIComponent(segment);
}

export interface GraphRequestOptions {
  method?: "GET" | "POST" | "PATCH" | "PUT" | "DELETE";
  /** Path relative to the Graph v1.0 base, e.g. "/me/messages". */
  path: string;
  /** Query string params. */
  query?: Record<string, string | number | boolean | undefined>;
  /** JSON body for POST/PATCH/PUT. */
  body?: unknown;
  /** Extra headers (e.g. Prefer for immutable IDs or paging). */
  headers?: Record<string, string>;
}

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly body: unknown,
  ) {
    super(message);
    this.name = "GraphError";
  }
}

/**
 * Thin fetch-based Microsoft Graph client. Keeps the dependency surface small
 * (no full graph SDK) while handling auth, query building, JSON, and errors.
 */
export class GraphClient {
  constructor(private readonly auth: AuthProvider) {}

  async request<T = unknown>(opts: GraphRequestOptions): Promise<T> {
    const token = await this.auth.getAccessToken();

    const url = new URL(GRAPH_BASE + opts.path);
    if (opts.query) {
      for (const [k, v] of Object.entries(opts.query)) {
        if (v !== undefined) url.searchParams.set(k, String(v));
      }
    }

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      ...opts.headers,
    };

    let body: string | undefined;
    if (opts.body !== undefined) {
      headers["Content-Type"] = "application/json";
      body = JSON.stringify(opts.body);
    }

    const res = await fetch(url, {
      method: opts.method ?? "GET",
      headers,
      body,
    });

    // 204 No Content (common for DELETE / some PATCH).
    if (res.status === 204) return undefined as T;

    const text = await res.text();
    const parsed = text ? safeJson(text) : undefined;

    if (!res.ok) {
      const gErr = (parsed as any)?.error;
      const msg =
        gErr?.message ?? `Graph request failed (HTTP ${res.status})`;
      throw new GraphError(msg, res.status, parsed ?? text);
    }

    return parsed as T;
  }

  /** GET that transparently follows @odata.nextLink up to `maxPages`. */
  async getAllPages<T = unknown>(
    opts: GraphRequestOptions & { maxPages?: number },
  ): Promise<T[]> {
    const maxPages = opts.maxPages ?? 5;
    const items: T[] = [];
    let page = await this.request<{ value?: T[]; "@odata.nextLink"?: string }>(
      opts,
    );
    if (page.value) items.push(...page.value);

    let count = 1;
    while (page["@odata.nextLink"] && count < maxPages) {
      const nextLink = page["@odata.nextLink"];
      
      // Security: only follow nextLink URLs that point to the official Graph API
      // host. Never send the Bearer token to any other origin.
      let nextUrl: URL;
      try {
        nextUrl = new URL(nextLink);
      } catch {
        throw new GraphError(
          "Invalid @odata.nextLink URL",
          400,
          { nextLink },
        );
      }
      if (nextUrl.origin !== "https://graph.microsoft.com") {
        throw new GraphError(
          `@odata.nextLink origin must be https://graph.microsoft.com, got ${nextUrl.origin}`,
          400,
          { nextLink },
        );
      }

      const token = await this.auth.getAccessToken();
      const res = await fetch(nextLink, {
        headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
      });
      const text = await res.text();
      const parsed = safeJson(text) as {
        value?: T[];
        "@odata.nextLink"?: string;
      };
      if (!res.ok) {
        throw new GraphError(
          (parsed as any)?.error?.message ?? "Graph paging failed",
          res.status,
          parsed,
        );
      }
      if (parsed.value) items.push(...parsed.value);
      page = parsed;
      count++;
    }
    return items;
  }
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}
