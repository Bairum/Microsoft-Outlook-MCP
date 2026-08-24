import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphClient } from "../graph.js";
import { sanitizePathSegment } from "../graph.js";
import type { Config } from "../config.js";
import { ok, fail, type ToolResult } from "./util.js";

function summarizeContact(c: any) {
  return {
    id: c.id,
    displayName: c.displayName,
    givenName: c.givenName,
    surname: c.surname,
    emails: (c.emailAddresses ?? []).map((e: any) => e.address),
    mobilePhone: c.mobilePhone,
    businessPhones: c.businessPhones,
    companyName: c.companyName,
    jobTitle: c.jobTitle,
  };
}

const CONTACT_SELECT =
  "id,displayName,givenName,surname,emailAddresses,mobilePhone,businessPhones,companyName,jobTitle";

export function registerContactTools(
  server: McpServer,
  graph: GraphClient,
  config: Config,
): void {
  server.registerTool(
    "list_contacts",
    {
      title: "List contacts",
      description: "List contacts, optionally filtered by a search term.",
      inputSchema: {
        top: z.number().int().min(1).max(100).default(50),
        search: z
          .string()
          .optional()
          .describe('Search by name/email, e.g. "smith".'),
      },
    },
    async ({ top, search }): Promise<ToolResult> => {
      try {
        const query: Record<string, string | number> = {
          $top: top,
          $select: CONTACT_SELECT,
        };
        const headers: Record<string, string> = {};
        if (search) {
          query.$search = `"${search}"`;
          headers.ConsistencyLevel = "eventual";
        } else {
          query.$orderby = "displayName";
        }
        const data = await graph.request<{ value: any[] }>({
          path: "/me/contacts",
          query,
          headers,
        });
        return ok(data.value.map(summarizeContact));
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "get_contact",
    {
      title: "Get contact",
      description: "Fetch a single contact by id.",
      inputSchema: { id: z.string() },
    },
    async ({ id }): Promise<ToolResult> => {
      try {
        const c = await graph.request<any>({ path: `/me/contacts/${sanitizePathSegment(id, "contact id")}` });
        return ok(summarizeContact(c));
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "create_contact",
    {
      title: "Create contact",
      description: "Create a new contact.",
      inputSchema: {
        givenName: z.string().optional(),
        surname: z.string().optional(),
        displayName: z.string().optional(),
        emails: z.array(z.string()).optional().describe("Email addresses."),
        mobilePhone: z.string().optional(),
        companyName: z.string().optional(),
        jobTitle: z.string().optional(),
      },
    },
    async (args): Promise<ToolResult> => {
      try {
        const body: Record<string, unknown> = {};
        if (args.givenName) body.givenName = args.givenName;
        if (args.surname) body.surname = args.surname;
        if (args.displayName) body.displayName = args.displayName;
        if (args.mobilePhone) body.mobilePhone = args.mobilePhone;
        if (args.companyName) body.companyName = args.companyName;
        if (args.jobTitle) body.jobTitle = args.jobTitle;
        if (args.emails?.length) {
          body.emailAddresses = args.emails.map((address) => ({ address }));
        }
        const c = await graph.request<any>({
          method: "POST",
          path: "/me/contacts",
          body,
        });
        return ok(summarizeContact(c));
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "update_contact",
    {
      title: "Update contact",
      description: "Update fields on an existing contact. Only pass what changes.",
      inputSchema: {
        id: z.string(),
        givenName: z.string().optional(),
        surname: z.string().optional(),
        displayName: z.string().optional(),
        emails: z.array(z.string()).optional(),
        mobilePhone: z.string().optional(),
        companyName: z.string().optional(),
        jobTitle: z.string().optional(),
      },
    },
    async ({ id, ...args }): Promise<ToolResult> => {
      try {
        const body: Record<string, unknown> = {};
        if (args.givenName !== undefined) body.givenName = args.givenName;
        if (args.surname !== undefined) body.surname = args.surname;
        if (args.displayName !== undefined) body.displayName = args.displayName;
        if (args.mobilePhone !== undefined) body.mobilePhone = args.mobilePhone;
        if (args.companyName !== undefined) body.companyName = args.companyName;
        if (args.jobTitle !== undefined) body.jobTitle = args.jobTitle;
        if (args.emails !== undefined) {
          body.emailAddresses = args.emails.map((address) => ({ address }));
        }
        if (Object.keys(body).length === 0) {
          return fail("Provide at least one field to update.");
        }
        const c = await graph.request<any>({
          method: "PATCH",
          path: `/me/contacts/${sanitizePathSegment(id, "contact id")}`,
          body,
        });
        return ok(summarizeContact(c));
      } catch (e) {
        return fail(e);
      }
    },
  );

  // delete_contact is gated behind OUTLOOK_ALLOW_WRITES.
  if (config.allowWrites) {
    server.registerTool(
      "delete_contact",
      {
        title: "Delete contact",
        description: "Delete a contact by id. (Write-gated: requires OUTLOOK_ALLOW_WRITES=true)",
        inputSchema: { id: z.string() },
      },
      async ({ id }): Promise<ToolResult> => {
        try {
          await graph.request({ method: "DELETE", path: `/me/contacts/${sanitizePathSegment(id, "contact id")}` });
          return ok({ deleted: true, id });
        } catch (e) {
          return fail(e);
        }
      },
    );
  }
}
