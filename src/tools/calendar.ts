import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphClient } from "../graph.js";
import { sanitizePathSegment } from "../graph.js";
import type { Config } from "../config.js";
import { ok, fail, type ToolResult } from "./util.js";

const attendee = (addr: string) => ({
  emailAddress: { address: addr },
  type: "required",
});

function summarizeEvent(e: any) {
  return {
    id: e.id,
    subject: e.subject,
    start: e.start?.dateTime,
    end: e.end?.dateTime,
    timeZone: e.start?.timeZone,
    location: e.location?.displayName,
    organizer: e.organizer?.emailAddress?.address,
    isAllDay: e.isAllDay,
    isOnlineMeeting: e.isOnlineMeeting,
    onlineMeetingUrl: e.onlineMeeting?.joinUrl,
    webLink: e.webLink,
  };
}

const EVENT_SELECT =
  "id,subject,start,end,location,organizer,isAllDay,isOnlineMeeting,onlineMeeting,webLink";

export function registerCalendarTools(
  server: McpServer,
  graph: GraphClient,
  config: Config,
): void {
  server.registerTool(
    "list_events",
    {
      title: "List calendar events",
      description:
        "List events within a date/time window using calendarView. " +
        "Datetimes are ISO 8601; if no timezone offset is given they are " +
        "interpreted in the requested timeZone.",
      inputSchema: {
        start: z
          .string()
          .describe("Window start, ISO 8601, e.g. 2026-07-14T00:00:00."),
        end: z
          .string()
          .describe("Window end, ISO 8601, e.g. 2026-07-21T00:00:00."),
        timeZone: z
          .string()
          .default("UTC")
          .describe('IANA/Windows time zone, e.g. "Pacific Standard Time".'),
        top: z.number().int().min(1).max(100).default(50),
      },
    },
    async ({ start, end, timeZone, top }): Promise<ToolResult> => {
      try {
        const events = await graph.getAllPages<any>({
          path: "/me/calendarView",
          query: {
            startDateTime: start,
            endDateTime: end,
            $top: top,
            $select: EVENT_SELECT,
            $orderby: "start/dateTime",
          },
          headers: { Prefer: `outlook.timezone="${timeZone}"` },
          maxPages: 3,
        });
        return ok(events.map(summarizeEvent));
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "get_event",
    {
      title: "Get calendar event",
      description: "Fetch full details of one event by id, including body.",
      inputSchema: {
        id: z.string(),
        timeZone: z.string().default("UTC"),
      },
    },
    async ({ id, timeZone }): Promise<ToolResult> => {
      try {
        const e = await graph.request<any>({
          path: `/me/events/${sanitizePathSegment(id, "event id")}`,
          headers: { Prefer: `outlook.timezone="${timeZone}"` },
        });
        return ok({
          ...summarizeEvent(e),
          body: e.body?.content,
          attendees: (e.attendees ?? []).map((a: any) => ({
            address: a.emailAddress?.address,
            name: a.emailAddress?.name,
            type: a.type,
            response: a.status?.response,
          })),
        });
      } catch (err) {
        return fail(err);
      }
    },
  );

  // create_event with attendees is gated behind OUTLOOK_ALLOW_WRITES (counts as sending).
  // Creating events without attendees is always allowed.
  server.registerTool(
    "create_event",
    {
      title: "Create calendar event",
      description: "Create a new calendar event / meeting. Creating events with attendees requires OUTLOOK_ALLOW_WRITES=true.",
      inputSchema: {
        subject: z.string(),
        start: z.string().describe("ISO 8601 start datetime."),
        end: z.string().describe("ISO 8601 end datetime."),
        timeZone: z.string().default("UTC"),
        body: z.string().optional().describe("Event description / agenda."),
        location: z.string().optional(),
        attendees: z
          .array(z.string())
          .optional()
          .describe("Attendee email addresses (requires OUTLOOK_ALLOW_WRITES=true)."),
        isOnlineMeeting: z
          .boolean()
          .default(false)
          .describe("Create a Teams online meeting link."),
      },
    },
    async ({
      subject,
      start,
      end,
      timeZone,
      body,
      location,
      attendees,
      isOnlineMeeting,
    }): Promise<ToolResult> => {
      try {
        if (attendees?.length && !config.allowWrites) {
          return fail(
            "Creating events with attendees (sending meeting invites) requires " +
            "OUTLOOK_ALLOW_WRITES=true. Set this environment variable to enable, or " +
            "create the event without attendees.",
          );
        }
        const payload: Record<string, unknown> = {
          subject,
          start: { dateTime: start, timeZone },
          end: { dateTime: end, timeZone },
        };
        if (body) payload.body = { contentType: "text", content: body };
        if (location) payload.location = { displayName: location };
        if (attendees?.length) payload.attendees = attendees.map(attendee);
        if (isOnlineMeeting) {
          payload.isOnlineMeeting = true;
          payload.onlineMeetingProvider = "teamsForBusiness";
        }
        const e = await graph.request<any>({
          method: "POST",
          path: "/me/events",
          body: payload,
        });
        return ok(summarizeEvent(e));
      } catch (err) {
        return fail(err);
      }
    },
  );

  server.registerTool(
    "update_event",
    {
      title: "Update calendar event",
      description: "Update fields on an existing event. Only pass what changes.",
      inputSchema: {
        id: z.string(),
        subject: z.string().optional(),
        start: z.string().optional(),
        end: z.string().optional(),
        timeZone: z.string().default("UTC"),
        body: z.string().optional(),
        location: z.string().optional(),
      },
    },
    async ({ id, subject, start, end, timeZone, body, location }): Promise<ToolResult> => {
      try {
        const payload: Record<string, unknown> = {};
        if (subject !== undefined) payload.subject = subject;
        if (start !== undefined) payload.start = { dateTime: start, timeZone };
        if (end !== undefined) payload.end = { dateTime: end, timeZone };
        if (body !== undefined) payload.body = { contentType: "text", content: body };
        if (location !== undefined) payload.location = { displayName: location };
        if (Object.keys(payload).length === 0) {
          return fail("Provide at least one field to update.");
        }
        const e = await graph.request<any>({
          method: "PATCH",
          path: `/me/events/${sanitizePathSegment(id, "event id")}`,
          body: payload,
        });
        return ok(summarizeEvent(e));
      } catch (err) {
        return fail(err);
      }
    },
  );

  // delete_event is gated behind OUTLOOK_ALLOW_WRITES.
  if (config.allowWrites) {
    server.registerTool(
      "delete_event",
      {
        title: "Delete calendar event",
        description: "Delete/cancel an event by id. (Write-gated: requires OUTLOOK_ALLOW_WRITES=true)",
        inputSchema: { id: z.string() },
      },
      async ({ id }): Promise<ToolResult> => {
        try {
          await graph.request({ method: "DELETE", path: `/me/events/${sanitizePathSegment(id, "event id")}` });
          return ok({ deleted: true, id });
        } catch (err) {
          return fail(err);
        }
      },
    );
  }

  server.registerTool(
    "respond_to_event",
    {
      title: "Respond to event invite",
      description: "Accept, tentatively accept, or decline a meeting invite.",
      inputSchema: {
        id: z.string(),
        response: z.enum(["accept", "tentativelyAccept", "decline"]),
        comment: z.string().optional(),
        sendResponse: z.boolean().default(true),
      },
    },
    async ({ id, response, comment, sendResponse }): Promise<ToolResult> => {
      try {
        await graph.request({
          method: "POST",
          path: `/me/events/${sanitizePathSegment(id, "event id")}/${response}`,
          body: { comment: comment ?? "", sendResponse },
        });
        return ok({ responded: true, id, response });
      } catch (err) {
        return fail(err);
      }
    },
  );
}
