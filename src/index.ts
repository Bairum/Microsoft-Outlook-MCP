#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { loadConfig } from "./config.js";
import { AuthProvider } from "./auth.js";
import { GraphClient } from "./graph.js";
import { registerMailTools } from "./tools/mail.js";
import { registerCalendarTools } from "./tools/calendar.js";
import { registerContactTools } from "./tools/contacts.js";
import { registerRulesAndFolderTools } from "./tools/rules.js";
import { ok, fail, type ToolResult } from "./tools/util.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const auth = await AuthProvider.create(config);
  const graph = new GraphClient(auth);

  const server = new McpServer({
    name: "microsoft-outlook-mcp",
    version: "0.1.0",
  });

  // Account / session tools.
  server.registerTool(
    "whoami",
    {
      title: "Who am I",
      description:
        "Return the signed-in user's profile. Fails if not signed in.",
      inputSchema: {},
    },
    async (): Promise<ToolResult> => {
      try {
        const me = await graph.request<any>({
          path: "/me",
          query: {
            $select: "displayName,userPrincipalName,mail,jobTitle,id",
          },
        });
        return ok(me);
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "auth_status",
    {
      title: "Auth status",
      description:
        "Check whether the server currently has a cached sign-in. Use the " +
        "`npm run login` step (outside the MCP session) to sign in.",
      inputSchema: {},
    },
    async (): Promise<ToolResult> => {
      const signedIn = await auth.isSignedIn();
      return ok({
        signedIn,
        scopes: config.scopes,
        tenant: config.tenantId,
        expectedUsername: config.expectedUsername || "(any cached account)",
        allowWrites: config.allowWrites,
        hint: signedIn
          ? "Session cached. Tools should work."
          : "Not signed in. Run `npm run login` in a terminal, then retry.",
      });
    },
  );

  registerMailTools(server, graph, config);
  registerCalendarTools(server, graph, config);
  registerContactTools(server, graph, config);
  registerRulesAndFolderTools(server, graph, config);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write(
    "[outlook-mcp] server started on stdio; tenant=" +
      config.tenantId +
      "\n",
  );
}

main().catch((err) => {
  process.stderr.write(`[outlook-mcp] fatal: ${(err as Error).message}\n`);
  process.exit(1);
});
