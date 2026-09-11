import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphClient } from "../graph.js";
import { GraphError, sanitizePathSegment } from "../graph.js";
import type { Config } from "../config.js";
import type { DeltaStore } from "../deltaStore.js";
import { ok, fail, type ToolResult } from "./util.js";

/**
 * Compact message summary for delta/snapshot results.
 */
interface CompactMessage {
  id: string;
  subject?: string;
  from?: string;
  received?: string;
  isRead?: boolean;
  isDraft?: boolean;
  folder?: string;
  preview?: string;
}

/**
 * Compact event summary for delta/snapshot results.
 */
export interface CompactEvent {
  id: string;
  subject?: string;
  organizer?: string;
  start?: string;
  isAllDay?: boolean;
  attendeeCount?: number;
  hasNoAttendees?: boolean;
  isCancelled?: boolean;
}

/**
 * True when Graph marks the event cancelled, or Outlook prefixes the subject
 * with Canceled/Cancelled (common after a series instance is cancelled).
 */
export function isCancelledEvent(e: {
  isCancelled?: boolean;
  subject?: string;
}): boolean {
  if (e.isCancelled === true) return true;
  const subject = (e.subject ?? "").trim();
  return /^cancell?ed(\s|:)/i.test(subject);
}

function attendeeList(e: { attendees?: unknown }): unknown[] {
  return Array.isArray(e.attendees) ? e.attendees : [];
}

/**
 * Summarize a Graph message into a compact format.
 */
function summarizeMessage(m: any): CompactMessage {
  return {
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address,
    received: m.receivedDateTime,
    isRead: m.isRead,
    isDraft: m.isDraft,
    folder: m.parentFolderId,
    preview: m.bodyPreview,
  };
}

/**
 * Summarize a Graph event into a compact format, flagging no-attendee events.
 *
 * `hasNoAttendees` is a candidate flag only: calendar delta (and occasionally
 * calendarView) can omit or empty `attendees` on series instances even when
 * invitees exist. Callers must run {@link verifyZeroAttendeeEvents} before
 * returning results so the flag means "Graph Required list is empty after a
 * full read", not "this payload happened to lack attendees".
 *
 * Cancelled and deleted (`@removed`) events are never flagged as actionable
 * empty holds.
 */
export function summarizeEvent(e: any): CompactEvent {
  const attendees = attendeeList(e);
  const cancelled = isCancelledEvent(e);
  const removed = Boolean(e["@removed"]);
  return {
    id: e.id,
    subject: e.subject,
    organizer: e.organizer?.emailAddress?.address,
    start: e.start?.dateTime,
    isAllDay: e.isAllDay,
    attendeeCount: attendees.length,
    ...(cancelled ? { isCancelled: true } : {}),
    // Only empty, non-cancelled, non-deleted events are candidates.
    hasNoAttendees: !cancelled && !removed && attendees.length === 0,
  };
}

/**
 * Re-read only events that look like zero-attendee holds.
 *
 * Graph calendar delta often strips `attendees` from series instances. A
 * single `GET /me/events/{id}` corrects false positives. Verified empty
 * (or a missing event) keeps `hasNoAttendees: true`. Pending RSVP
 * (`notResponded`) is still an attendee and will clear the flag.
 */
export async function verifyZeroAttendeeEvents(
  graph: GraphClient,
  events: CompactEvent[]
): Promise<void> {
  for (const event of events) {
    if (!event.hasNoAttendees || event.isCancelled || !event.id) continue;

    try {
      const verified = await graph.request<{
        attendees?: unknown;
        isCancelled?: boolean;
      }>({
        path: `/me/events/${sanitizePathSegment(event.id, "event id")}`,
        query: { $select: "id,attendees,isCancelled" },
      });

      if (verified.isCancelled || isCancelledEvent(verified)) {
        event.isCancelled = true;
        event.hasNoAttendees = false;
        const count = attendeeList(verified).length;
        event.attendeeCount = count;
        continue;
      }

      const count = attendeeList(verified).length;
      event.attendeeCount = count;
      event.hasNoAttendees = count === 0;
    } catch (err) {
      // Missing/gone event: treat as still empty (keep the candidate flag).
      if (err instanceof GraphError && (err.status === 404 || err.status === 410)) {
        continue;
      }
      throw err;
    }
  }
}

/**
 * Fetch a mail folder by display name. Returns undefined if not found.
 */
async function findFolderByName(
  graph: GraphClient,
  name: string
): Promise<string | undefined> {
  try {
    const data = await graph.request<{ value: any[] }>({
      path: "/me/mailFolders",
      query: { $filter: `displayName eq '${name}'`, $select: "id,displayName" },
    });
    return data.value[0]?.id;
  } catch {
    return undefined;
  }
}

/**
 * Fetch delta changes for a single mail folder.
 */
async function fetchMailFolderDelta(
  graph: GraphClient,
  folderId: string,
  deltaLink?: string
): Promise<{ messages: CompactMessage[]; nextDeltaLink: string }> {
  const messages: CompactMessage[] = [];

  // If we have a deltaLink, use it to get changes since last sync.
  // Otherwise, do a full sync starting from the folder's messages endpoint.
  // Security: sanitize folder ID to prevent path traversal.
  const path = deltaLink ?? `/me/mailFolders/${sanitizePathSegment(folderId, "folder id")}/messages/delta`;
  const query = deltaLink
    ? undefined
    : {
        $select:
          "id,subject,from,receivedDateTime,isRead,isDraft,parentFolderId,bodyPreview",
      };

  // Fetch the first page. GraphClient.request now handles absolute URLs (deltaLink)
  // and validates their origin.
  let page = await graph.request<{ value: any[]; "@odata.deltaLink"?: string; "@odata.nextLink"?: string }>({
    path,
    query,
  });

  if (page.value) messages.push(...page.value.map(summarizeMessage));

  // Follow nextLink pagination until we reach the deltaLink.
  // GraphClient.request handles absolute nextLink URLs and validates origin.
  while (page["@odata.nextLink"]) {
    page = await graph.request<{ value: any[]; "@odata.deltaLink"?: string; "@odata.nextLink"?: string }>({
      path: page["@odata.nextLink"],
    });

    if (page.value) messages.push(...page.value.map(summarizeMessage));
  }

  if (!page["@odata.deltaLink"]) {
    throw new Error("No @odata.deltaLink returned from mail delta endpoint");
  }

  return { messages, nextDeltaLink: page["@odata.deltaLink"] };
}

/**
 * Fetch delta changes for calendar events.
 */
async function fetchCalendarDelta(
  graph: GraphClient,
  deltaLink?: string
): Promise<{ events: CompactEvent[]; nextDeltaLink: string }> {
  const events: CompactEvent[] = [];

  // If we have a deltaLink, use it. Otherwise, start with a full sync covering
  // past 7 days and next 30 days. startDateTime/endDateTime are only for the
  // first relative call; subsequent deltaLink calls already have them baked in.
  const path = deltaLink ?? "/me/calendarView/delta";
  const query = deltaLink
    ? undefined
    : {
        startDateTime: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString(),
        endDateTime: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString(),
        $select: "id,subject,organizer,start,isAllDay,attendees,isCancelled",
      };

  // GraphClient.request now handles absolute URLs (deltaLink) and validates origin.
  let page = await graph.request<{ value: any[]; "@odata.deltaLink"?: string; "@odata.nextLink"?: string }>({
    path,
    query,
  });

  if (page.value) events.push(...page.value.map(summarizeEvent));

  // Follow nextLink pagination. GraphClient.request validates origin.
  while (page["@odata.nextLink"]) {
    page = await graph.request<{ value: any[]; "@odata.deltaLink"?: string; "@odata.nextLink"?: string }>({
      path: page["@odata.nextLink"],
    });

    if (page.value) events.push(...page.value.map(summarizeEvent));
  }

  if (!page["@odata.deltaLink"]) {
    throw new Error("No @odata.deltaLink returned from calendar delta endpoint");
  }

  // Delta often omits attendees on series instances — verify only empty candidates.
  await verifyZeroAttendeeEvents(graph, events);

  return { events, nextDeltaLink: page["@odata.deltaLink"] };
}

export function registerDeltaTools(
  server: McpServer,
  graph: GraphClient,
  config: Config,
  deltaStore: DeltaStore
): void {
  server.registerTool(
    "list_changes",
    {
      title: "List changes (delta sync)",
      description:
        "Incremental delta sync for mail (Inbox, Action, Staging, Drafts) and calendar. " +
        "Returns only what changed since the last call. Use reset=true to start a fresh sync. " +
        "Delta cursors are persisted securely in OS-native encrypted storage. " +
        "Read-only: does not send mail, move messages, or modify events.",
      inputSchema: {
        reset: z
          .boolean()
          .default(false)
          .describe("If true, discard stored cursors and perform a full sync."),
      },
    },
    async ({ reset }): Promise<ToolResult> => {
      try {
        if (reset) {
          await deltaStore.clear();
        }

        const result: {
          mail: Record<string, CompactMessage[]>;
          calendar: CompactEvent[];
          skippedFolders: string[];
        } = {
          mail: {},
          calendar: [],
          skippedFolders: [],
        };

        // Well-known folder names to sync. If a folder doesn't exist, skip it.
        const folderNames = ["Inbox", "Action", "Staging", "Drafts"];

        for (const name of folderNames) {
          const folderId = await findFolderByName(graph, name);
          if (!folderId) {
            result.skippedFolders.push(name);
            continue;
          }

          const cursorKey = `mail-${name.toLowerCase()}`;
          const storedCursor = await deltaStore.get(cursorKey);

          const { messages, nextDeltaLink } = await fetchMailFolderDelta(
            graph,
            folderId,
            storedCursor
          );

          result.mail[name] = messages;
          await deltaStore.set(cursorKey, nextDeltaLink);
        }

        // Calendar delta.
        const calendarCursor = await deltaStore.get("calendar");
        const { events, nextDeltaLink } = await fetchCalendarDelta(
          graph,
          calendarCursor
        );

        result.calendar = events;
        await deltaStore.set("calendar", nextDeltaLink);

        return ok(result);
      } catch (e) {
        return fail(e);
      }
    }
  );

  server.registerTool(
    "get_ops_snapshot",
    {
      title: "Get operational snapshot",
      description:
        "One-shot snapshot of recent mailbox activity: last 48 hours of Inbox, " +
        "leftover Staging and Action (if folders exist), unsent Drafts (recent), " +
        "and calendar today + tomorrow in a specified IANA timezone (default UTC). " +
        "Flags zero-attendee calendar events. Payloads are compact summaries. " +
        "Read-only: does not send mail, move messages, or modify events.",
      inputSchema: {
        timezone: z
          .string()
          .default("UTC")
          .describe(
            "IANA timezone for calendar window, e.g. 'America/New_York' or 'Europe/London'. Defaults to UTC."
          ),
      },
    },
    async ({ timezone }): Promise<ToolResult> => {
      try {
        const result: {
          inbox48h: CompactMessage[];
          staging: CompactMessage[];
          action: CompactMessage[];
          drafts: CompactMessage[];
          calendarTodayTomorrow: CompactEvent[];
          skippedFolders: string[];
        } = {
          inbox48h: [],
          staging: [],
          action: [],
          drafts: [],
          calendarTodayTomorrow: [],
          skippedFolders: [],
        };

        const now = new Date();
        const fortyEightHoursAgo = new Date(now.getTime() - 48 * 60 * 60 * 1000);

        // Inbox: last 48 hours.
        const inboxId = await findFolderByName(graph, "Inbox");
        if (inboxId) {
          const inboxData = await graph.request<{ value: any[] }>({
            path: `/me/mailFolders/${sanitizePathSegment(inboxId, "inbox folder id")}/messages`,
            query: {
              $filter: `receivedDateTime ge ${fortyEightHoursAgo.toISOString()}`,
              $select:
                "id,subject,from,receivedDateTime,isRead,isDraft,parentFolderId,bodyPreview",
              $orderby: "receivedDateTime desc",
              $top: 100,
            },
          });
          result.inbox48h = inboxData.value.map(summarizeMessage);
        } else {
          result.skippedFolders.push("Inbox");
        }

        // Staging: all messages.
        const stagingId = await findFolderByName(graph, "Staging");
        if (stagingId) {
          const stagingData = await graph.request<{ value: any[] }>({
            path: `/me/mailFolders/${sanitizePathSegment(stagingId, "staging folder id")}/messages`,
            query: {
              $select:
                "id,subject,from,receivedDateTime,isRead,isDraft,parentFolderId,bodyPreview",
              $orderby: "receivedDateTime desc",
              $top: 100,
            },
          });
          result.staging = stagingData.value.map(summarizeMessage);
        } else {
          result.skippedFolders.push("Staging");
        }

        // Action: all messages.
        const actionId = await findFolderByName(graph, "Action");
        if (actionId) {
          const actionData = await graph.request<{ value: any[] }>({
            path: `/me/mailFolders/${sanitizePathSegment(actionId, "action folder id")}/messages`,
            query: {
              $select:
                "id,subject,from,receivedDateTime,isRead,isDraft,parentFolderId,bodyPreview",
              $orderby: "receivedDateTime desc",
              $top: 100,
            },
          });
          result.action = actionData.value.map(summarizeMessage);
        } else {
          result.skippedFolders.push("Action");
        }

        // Drafts: unsent drafts.
        const draftsId = await findFolderByName(graph, "Drafts");
        if (draftsId) {
          const draftsData = await graph.request<{ value: any[] }>({
            path: `/me/mailFolders/${sanitizePathSegment(draftsId, "drafts folder id")}/messages`,
            query: {
              $filter: "isDraft eq true",
              $select:
                "id,subject,from,receivedDateTime,isRead,isDraft,parentFolderId,bodyPreview",
              $orderby: "lastModifiedDateTime desc",
              $top: 50,
            },
          });
          result.drafts = draftsData.value.map(summarizeMessage);
        } else {
          result.skippedFolders.push("Drafts");
        }

        // Calendar: today + tomorrow in the specified timezone.
        // Construct date boundaries reliably without relying on Date parsing.
        // Get current date parts in the target timezone.
        const dateFormatter = new Intl.DateTimeFormat("en-US", {
          timeZone: timezone,
          year: "numeric",
          month: "2-digit",
          day: "2-digit",
        });
        const parts = dateFormatter.formatToParts(now);
        const year = parts.find((p) => p.type === "year")!.value;
        const month = parts.find((p) => p.type === "month")!.value;
        const day = parts.find((p) => p.type === "day")!.value;

        // Today start at 00:00 in the target timezone.
        const todayStart = `${year}-${month}-${day}T00:00:00`;
        
        // Tomorrow end at 23:59:59 (2 days from today start).
        const todayDate = new Date(`${year}-${month}-${day}T00:00:00Z`);
        todayDate.setUTCDate(todayDate.getUTCDate() + 2);
        const tomorrowEnd = todayDate.toISOString().split("T")[0] + "T23:59:59";

        const calendarData = await graph.request<{ value: any[] }>({
          path: "/me/calendarView",
          query: {
            startDateTime: todayStart,
            endDateTime: tomorrowEnd,
            $select: "id,subject,organizer,start,isAllDay,attendees,isCancelled",
            $orderby: "start/dateTime",
            $top: 100,
          },
          headers: { Prefer: `outlook.timezone="${timezone}"` },
        });

        const calendarEvents = calendarData.value.map(summarizeEvent);
        await verifyZeroAttendeeEvents(graph, calendarEvents);
        result.calendarTodayTomorrow = calendarEvents;

        return ok(result);
      } catch (e) {
        return fail(e);
      }
    }
  );
}
