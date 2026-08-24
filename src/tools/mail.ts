import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphClient } from "../graph.js";
import { sanitizePathSegment } from "../graph.js";
import type { Config } from "../config.js";
import { ok, fail, type ToolResult } from "./util.js";

const recipient = (addr: string) => ({ emailAddress: { address: addr } });

interface MessageLite {
  id: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  toRecipients?: { emailAddress?: { address?: string } }[];
  receivedDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  bodyPreview?: string;
  webLink?: string;
}

function summarizeMessage(m: MessageLite) {
  return {
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address,
    fromName: m.from?.emailAddress?.name,
    received: m.receivedDateTime,
    isRead: m.isRead,
    hasAttachments: m.hasAttachments,
    preview: m.bodyPreview,
    webLink: m.webLink,
  };
}

const SELECT =
  "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink";

export function registerMailTools(server: McpServer, graph: GraphClient, config: Config): void {
  server.registerTool(
    "list_messages",
    {
      title: "List messages",
      description:
        "List messages from a mail folder (default: inbox). Supports OData " +
        "$filter and $search. Returns compact summaries, not full bodies.",
      inputSchema: {
        folder: z
          .string()
          .optional()
          .describe(
            "Well-known folder name (inbox, drafts, sentitems, deleteditems, " +
              "archive, junkemail) or a folder id. Defaults to inbox.",
          ),
        top: z.number().int().min(1).max(100).default(25),
        search: z
          .string()
          .optional()
          .describe('Full-text search across the mailbox, e.g. "invoice".'),
        filter: z
          .string()
          .optional()
          .describe('OData $filter, e.g. "isRead eq false".'),
        unreadOnly: z.boolean().optional(),
      },
    },
    async ({ folder, top, search, filter, unreadOnly }): Promise<ToolResult> => {
      try {
        const seg = folder ? `/me/mailFolders/${sanitizePathSegment(folder, "folder")}/messages` : "/me/messages";
        const query: Record<string, string | number> = {
          $top: top,
          $select: SELECT,
          $orderby: "receivedDateTime desc",
        };
        const filters: string[] = [];
        if (filter) filters.push(filter);
        if (unreadOnly) filters.push("isRead eq false");
        if (filters.length) query.$filter = filters.join(" and ");

        const headers: Record<string, string> = {};
        if (search) {
          query.$search = `"${search}"`;
          // $search cannot combine with $orderby.
          delete query.$orderby;
        }

        const data = await graph.request<{ value: MessageLite[] }>({
          path: seg,
          query,
          headers,
        });
        return ok(data.value.map(summarizeMessage));
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "get_message",
    {
      title: "Get message",
      description:
        "Fetch the full content of a single message by id, including body.",
      inputSchema: {
        id: z.string(),
        bodyType: z
          .enum(["text", "html"])
          .default("text")
          .describe("Return body as plain text or HTML."),
      },
    },
    async ({ id, bodyType }): Promise<ToolResult> => {
      try {
        const data = await graph.request<any>({
          path: `/me/messages/${sanitizePathSegment(id, "message id")}`,
          headers: { Prefer: `outlook.body-content-type="${bodyType}"` },
        });
        return ok({
          id: data.id,
          subject: data.subject,
          from: data.from?.emailAddress,
          to: (data.toRecipients ?? []).map(
            (r: any) => r.emailAddress?.address,
          ),
          cc: (data.ccRecipients ?? []).map((r: any) => r.emailAddress?.address),
          received: data.receivedDateTime,
          isRead: data.isRead,
          hasAttachments: data.hasAttachments,
          body: data.body?.content,
          webLink: data.webLink,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  // send_mail and reply_to_message are gated behind OUTLOOK_ALLOW_WRITES.
  if (config.allowWrites) {
    server.registerTool(
      "send_mail",
      {
        title: "Send mail",
        description: "Compose and send a new email immediately. (Write-gated: requires OUTLOOK_ALLOW_WRITES=true)",
        inputSchema: {
          to: z.array(z.string()).min(1).describe("Recipient email addresses."),
          subject: z.string(),
          body: z.string(),
          bodyType: z.enum(["text", "html"]).default("text"),
          cc: z.array(z.string()).optional(),
          bcc: z.array(z.string()).optional(),
          saveToSentItems: z.boolean().default(true),
        },
      },
      async ({ to, subject, body, bodyType, cc, bcc, saveToSentItems }): Promise<ToolResult> => {
        try {
          await graph.request({
            method: "POST",
            path: "/me/sendMail",
            body: {
              message: {
                subject,
                body: { contentType: bodyType, content: body },
                toRecipients: to.map(recipient),
                ccRecipients: (cc ?? []).map(recipient),
                bccRecipients: (bcc ?? []).map(recipient),
              },
              saveToSentItems,
            },
          });
          return ok({ sent: true, to, subject });
        } catch (e) {
          return fail(e);
        }
      },
    );

    server.registerTool(
      "reply_to_message",
      {
        title: "Reply to message",
        description:
          "Reply to an existing message. Use replyAll to include everyone. (Write-gated: requires OUTLOOK_ALLOW_WRITES=true)",
        inputSchema: {
          id: z.string(),
          comment: z.string().describe("The reply text."),
          replyAll: z.boolean().default(false),
        },
      },
      async ({ id, comment, replyAll }): Promise<ToolResult> => {
        try {
          await graph.request({
            method: "POST",
            path: `/me/messages/${sanitizePathSegment(id, "message id")}/${replyAll ? "replyAll" : "reply"}`,
            body: { comment },
          });
          return ok({ replied: true, id, replyAll });
        } catch (e) {
          return fail(e);
        }
      },
    );
  }

  server.registerTool(
    "update_message",
    {
      title: "Update message flags",
      description: "Mark a message read/unread or set its flag status.",
      inputSchema: {
        id: z.string(),
        isRead: z.boolean().optional(),
        flag: z.enum(["flagged", "complete", "notFlagged"]).optional(),
      },
    },
    async ({ id, isRead, flag }): Promise<ToolResult> => {
      try {
        const body: Record<string, unknown> = {};
        if (isRead !== undefined) body.isRead = isRead;
        if (flag) body.flag = { flagStatus: flag };
        if (Object.keys(body).length === 0) {
          return fail("Provide at least one of isRead or flag.");
        }
        const data = await graph.request<any>({
          method: "PATCH",
          path: `/me/messages/${sanitizePathSegment(id, "message id")}`,
          body,
        });
        return ok({ id: data.id, isRead: data.isRead, flag: data.flag });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "move_message",
    {
      title: "Move message",
      description: "Move a message to another mail folder.",
      inputSchema: {
        id: z.string(),
        destinationFolder: z
          .string()
          .describe(
            "Well-known folder name (e.g. archive, deleteditems, junkemail) " +
              "or a folder id.",
          ),
      },
    },
    async ({ id, destinationFolder }): Promise<ToolResult> => {
      try {
        const data = await graph.request<any>({
          method: "POST",
          path: `/me/messages/${sanitizePathSegment(id, "message id")}/move`,
          body: { destinationId: destinationFolder },
        });
        return ok({ moved: true, newId: data.id, parentFolderId: data.parentFolderId });
      } catch (e) {
        return fail(e);
      }
    },
  );

  // delete_message with permanent=true is gated behind OUTLOOK_ALLOW_WRITES.
  // Soft delete (move to trash) is always allowed.
  server.registerTool(
    "delete_message",
    {
      title: "Delete message",
      description:
        "Delete a message. By default moves it to Deleted Items (recoverable). " +
        "Permanent deletion requires OUTLOOK_ALLOW_WRITES=true.",
      inputSchema: {
        id: z.string(),
        permanent: z
          .boolean()
          .default(false)
          .describe("If true, permanently delete instead of moving to trash (requires OUTLOOK_ALLOW_WRITES=true)."),
      },
    },
    async ({ id, permanent }): Promise<ToolResult> => {
      try {
        const safeId = sanitizePathSegment(id, "message id");
        if (permanent) {
          if (!config.allowWrites) {
            return fail(
              "Permanent message deletion requires OUTLOOK_ALLOW_WRITES=true. " +
              "Set this environment variable to enable dangerous writes, or use " +
              "permanent=false to move to trash instead.",
            );
          }
          await graph.request({ method: "DELETE", path: `/me/messages/${safeId}` });
        } else {
          await graph.request({
            method: "POST",
            path: `/me/messages/${safeId}/move`,
            body: { destinationId: "deleteditems" },
          });
        }
        return ok({ deleted: true, id, permanent });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "create_draft",
    {
      title: "Create draft",
      description:
        "Create a new draft message without sending. The draft is saved in the " +
        "Drafts folder and can be edited, sent later, or discarded. Does not " +
        "require OUTLOOK_ALLOW_WRITES.",
      inputSchema: {
        subject: z.string().optional(),
        body: z.string().optional(),
        bodyType: z.enum(["text", "html"]).default("text"),
        to: z
          .array(z.string())
          .optional()
          .describe("Recipient email addresses."),
        cc: z.array(z.string()).optional(),
        bcc: z.array(z.string()).optional(),
      },
    },
    async ({ subject, body, bodyType, to, cc, bcc }): Promise<ToolResult> => {
      try {
        const message: Record<string, unknown> = {};
        if (subject !== undefined) message.subject = subject;
        if (body !== undefined) {
          message.body = { contentType: bodyType, content: body };
        }
        if (to && to.length > 0) {
          message.toRecipients = to.map(recipient);
        }
        if (cc && cc.length > 0) {
          message.ccRecipients = cc.map(recipient);
        }
        if (bcc && bcc.length > 0) {
          message.bccRecipients = bcc.map(recipient);
        }

        const draft = await graph.request<any>({
          method: "POST",
          path: "/me/messages",
          body: message,
        });

        return ok({
          id: draft.id,
          subject: draft.subject,
          isDraft: draft.isDraft,
          webLink: draft.webLink,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "create_reply_draft",
    {
      title: "Create reply draft",
      description:
        "Create a reply draft for an existing message without sending. The draft " +
        "is created as a reply to the original sender only. Does not require " +
        "OUTLOOK_ALLOW_WRITES.",
      inputSchema: {
        id: z.string().describe("The message id to reply to."),
        comment: z
          .string()
          .optional()
          .describe("Optional reply text to include in the draft body."),
      },
    },
    async ({ id, comment }): Promise<ToolResult> => {
      try {
        const body: Record<string, unknown> = {};
        if (comment !== undefined) body.comment = comment;

        const draft = await graph.request<any>({
          method: "POST",
          path: `/me/messages/${sanitizePathSegment(id, "message id")}/createReply`,
          body: Object.keys(body).length > 0 ? body : undefined,
        });

        return ok({
          id: draft.id,
          subject: draft.subject,
          isDraft: draft.isDraft,
          inReplyTo: id,
          webLink: draft.webLink,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "create_reply_all_draft",
    {
      title: "Create reply-all draft",
      description:
        "Create a reply-all draft for an existing message without sending. The " +
        "draft is created as a reply to all recipients of the original message. " +
        "Does not require OUTLOOK_ALLOW_WRITES.",
      inputSchema: {
        id: z.string().describe("The message id to reply to."),
        comment: z
          .string()
          .optional()
          .describe("Optional reply text to include in the draft body."),
      },
    },
    async ({ id, comment }): Promise<ToolResult> => {
      try {
        const body: Record<string, unknown> = {};
        if (comment !== undefined) body.comment = comment;

        const draft = await graph.request<any>({
          method: "POST",
          path: `/me/messages/${sanitizePathSegment(id, "message id")}/createReplyAll`,
          body: Object.keys(body).length > 0 ? body : undefined,
        });

        return ok({
          id: draft.id,
          subject: draft.subject,
          isDraft: draft.isDraft,
          inReplyTo: id,
          replyAll: true,
          webLink: draft.webLink,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "create_forward_draft",
    {
      title: "Create forward draft",
      description:
        "Create a forward draft for an existing message without sending. Does not " +
        "require OUTLOOK_ALLOW_WRITES.",
      inputSchema: {
        id: z.string().describe("The message id to forward."),
        comment: z
          .string()
          .optional()
          .describe("Optional text to include in the draft body."),
        to: z
          .array(z.string())
          .optional()
          .describe("Recipient email addresses for the forward."),
      },
    },
    async ({ id, comment, to }): Promise<ToolResult> => {
      try {
        const body: Record<string, unknown> = {};
        if (comment !== undefined) body.comment = comment;
        if (to && to.length > 0) {
          body.toRecipients = to.map(recipient);
        }

        const draft = await graph.request<any>({
          method: "POST",
          path: `/me/messages/${sanitizePathSegment(id, "message id")}/createForward`,
          body: Object.keys(body).length > 0 ? body : undefined,
        });

        return ok({
          id: draft.id,
          subject: draft.subject,
          isDraft: draft.isDraft,
          forwarding: id,
          webLink: draft.webLink,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "update_draft",
    {
      title: "Update draft",
      description:
        "Update an existing draft message. Can update subject, body, and " +
        "recipients. Only works on draft messages (not sent messages). Does not " +
        "require OUTLOOK_ALLOW_WRITES.",
      inputSchema: {
        id: z.string().describe("The draft message id to update."),
        subject: z.string().optional(),
        body: z.string().optional(),
        bodyType: z
          .enum(["text", "html"])
          .default("text")
          .describe("Content type for the body field."),
        to: z.array(z.string()).optional(),
        cc: z.array(z.string()).optional(),
        bcc: z.array(z.string()).optional(),
      },
    },
    async ({ id, subject, body, bodyType, to, cc, bcc }): Promise<ToolResult> => {
      try {
        const update: Record<string, unknown> = {};
        if (subject !== undefined) update.subject = subject;
        if (body !== undefined) {
          update.body = { contentType: bodyType, content: body };
        }
        if (to !== undefined) {
          update.toRecipients = to.map(recipient);
        }
        if (cc !== undefined) {
          update.ccRecipients = cc.map(recipient);
        }
        if (bcc !== undefined) {
          update.bccRecipients = bcc.map(recipient);
        }

        if (Object.keys(update).length === 0) {
          return fail("Provide at least one field to update.");
        }

        const draft = await graph.request<any>({
          method: "PATCH",
          path: `/me/messages/${sanitizePathSegment(id, "message id")}`,
          body: update,
        });

        return ok({
          id: draft.id,
          subject: draft.subject,
          isDraft: draft.isDraft,
          webLink: draft.webLink,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "list_mail_folders",
    {
      title: "List mail folders",
      description: "List mail folders in the mailbox with unread/total counts.",
      inputSchema: {
        top: z.number().int().min(1).max(100).default(50),
      },
    },
    async ({ top }): Promise<ToolResult> => {
      try {
        const data = await graph.request<{ value: any[] }>({
          path: "/me/mailFolders",
          query: {
            $top: top,
            $select: "id,displayName,unreadItemCount,totalItemCount",
          },
        });
        return ok(
          data.value.map((f) => ({
            id: f.id,
            name: f.displayName,
            unread: f.unreadItemCount,
            total: f.totalItemCount,
          })),
        );
      } catch (e) {
        return fail(e);
      }
    },
  );
}
