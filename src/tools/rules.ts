import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphClient } from "../graph.js";
import { sanitizePathSegment } from "../graph.js";
import type { Config } from "../config.js";
import { ok, fail, type ToolResult } from "./util.js";

const recipient = (addr: string) => ({ emailAddress: { address: addr } });

// ---------------------------------------------------------------------------
// Shared schema fragments for message-rule conditions and actions. Exposed as
// a practical subset of the full Graph messageRule schema.
// ---------------------------------------------------------------------------

const conditionFields = {
  fromAddresses: z
    .array(z.string())
    .optional()
    .describe("Match when the sender is one of these email addresses."),
  senderContains: z
    .array(z.string())
    .optional()
    .describe("Match when the sender text contains any of these strings."),
  subjectContains: z
    .array(z.string())
    .optional()
    .describe("Match when the subject contains any of these strings."),
  bodyContains: z
    .array(z.string())
    .optional()
    .describe("Match when the body contains any of these strings."),
  importance: z.enum(["low", "normal", "high"]).optional(),
  hasAttachments: z.boolean().optional(),
};

const actionFields = {
  moveToFolder: z
    .string()
    .optional()
    .describe("Folder id or well-known name to move matching messages to."),
  copyToFolder: z.string().optional(),
  markAsRead: z.boolean().optional(),
  markImportance: z.enum(["low", "normal", "high"]).optional(),
  forwardTo: z
    .array(z.string())
    .optional()
    .describe("Forward matching messages to these addresses."),
  delete: z
    .boolean()
    .optional()
    .describe("Move matching messages to Deleted Items."),
  stopProcessingRules: z.boolean().optional(),
};

function buildConditions(a: Record<string, unknown>): Record<string, unknown> {
  const c: Record<string, unknown> = {};
  if (a.fromAddresses)
    c.fromAddresses = (a.fromAddresses as string[]).map(recipient);
  if (a.senderContains) c.senderContains = a.senderContains;
  if (a.subjectContains) c.subjectContains = a.subjectContains;
  if (a.bodyContains) c.bodyContains = a.bodyContains;
  if (a.importance) c.importance = a.importance;
  if (a.hasAttachments !== undefined) c.hasAttachments = a.hasAttachments;
  return c;
}

function buildActions(a: Record<string, unknown>): Record<string, unknown> {
  const act: Record<string, unknown> = {};
  if (a.moveToFolder) act.moveToFolder = a.moveToFolder;
  if (a.copyToFolder) act.copyToFolder = a.copyToFolder;
  if (a.markAsRead !== undefined) act.markAsRead = a.markAsRead;
  if (a.markImportance) act.markImportance = a.markImportance;
  if (a.forwardTo)
    act.forwardTo = (a.forwardTo as string[]).map(recipient);
  if (a.delete !== undefined) act.delete = a.delete;
  if (a.stopProcessingRules !== undefined)
    act.stopProcessingRules = a.stopProcessingRules;
  return act;
}

function summarizeRule(r: any) {
  return {
    id: r.id,
    displayName: r.displayName,
    sequence: r.sequence,
    isEnabled: r.isEnabled,
    hasError: r.hasError,
    conditions: r.conditions,
    actions: r.actions,
  };
}

export function registerRulesAndFolderTools(
  server: McpServer,
  graph: GraphClient,
  config: Config,
): void {
  // -------------------------------------------------------------------------
  // Mail folder management (list lives in mail.ts as list_mail_folders).
  // -------------------------------------------------------------------------

  server.registerTool(
    "create_mail_folder",
    {
      title: "Create mail folder",
      description:
        "Create a mail folder, optionally nested under a parent folder.",
      inputSchema: {
        displayName: z.string(),
        parentFolder: z
          .string()
          .optional()
          .describe(
            "Parent folder id or well-known name (e.g. inbox). Omit for a " +
              "top-level folder.",
          ),
      },
    },
    async ({ displayName, parentFolder }): Promise<ToolResult> => {
      try {
        const path = parentFolder
          ? `/me/mailFolders/${sanitizePathSegment(parentFolder, "parentFolder")}/childFolders`
          : "/me/mailFolders";
        const f = await graph.request<any>({
          method: "POST",
          path,
          body: { displayName },
        });
        return ok({ id: f.id, name: f.displayName, parentFolderId: f.parentFolderId });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "rename_mail_folder",
    {
      title: "Rename mail folder",
      description: "Rename an existing mail folder.",
      inputSchema: {
        id: z.string().describe("Folder id (not a well-known name)."),
        displayName: z.string(),
      },
    },
    async ({ id, displayName }): Promise<ToolResult> => {
      try {
        const f = await graph.request<any>({
          method: "PATCH",
          path: `/me/mailFolders/${sanitizePathSegment(id, "folder id")}`,
          body: { displayName },
        });
        return ok({ id: f.id, name: f.displayName });
      } catch (e) {
        return fail(e);
      }
    },
  );

  // delete_mail_folder is gated behind OUTLOOK_ALLOW_WRITES.
  if (config.allowWrites) {
    server.registerTool(
      "delete_mail_folder",
      {
        title: "Delete mail folder",
        description:
          "Delete a mail folder and its contents. This is not easily " +
          "recoverable — use with care. (Write-gated: requires OUTLOOK_ALLOW_WRITES=true)",
        inputSchema: { id: z.string().describe("Folder id (not a well-known name).") },
      },
      async ({ id }): Promise<ToolResult> => {
        try {
          await graph.request({ method: "DELETE", path: `/me/mailFolders/${sanitizePathSegment(id, "folder id")}` });
          return ok({ deleted: true, id });
        } catch (e) {
          return fail(e);
        }
      },
    );
  }

  // -------------------------------------------------------------------------
  // Inbox message rules. Requires MailboxSettings.ReadWrite.
  // -------------------------------------------------------------------------

  server.registerTool(
    "list_message_rules",
    {
      title: "List inbox rules",
      description: "List message (inbox) rules with their conditions/actions.",
      inputSchema: {},
    },
    async (): Promise<ToolResult> => {
      try {
        const data = await graph.request<{ value: any[] }>({
          path: "/me/mailFolders/inbox/messageRules",
        });
        return ok(data.value.map(summarizeRule));
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "get_message_rule",
    {
      title: "Get inbox rule",
      description: "Fetch a single inbox rule by id.",
      inputSchema: { id: z.string() },
    },
    async ({ id }): Promise<ToolResult> => {
      try {
        const r = await graph.request<any>({
          path: `/me/mailFolders/inbox/messageRules/${sanitizePathSegment(id, "rule id")}`,
        });
        return ok(summarizeRule(r));
      } catch (e) {
        return fail(e);
      }
    },
  );

  // Inbox rule write operations are DISABLED in this hardened fork.
  // Creating or updating rules with forwardTo is a durable mail-forwarding
  // vector that is too dangerous for a local Graph proxy. List and get are
  // still available for inspection.
  //
  // If you need to manage inbox rules, set OUTLOOK_ALLOW_WRITES=true and
  // carefully review any rule actions before creating them.
}
