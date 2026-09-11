import { strict as assert } from "node:assert";
import { GraphError } from "../graph.js";
import type { GraphClient } from "../graph.js";
import {
  isCancelledEvent,
  summarizeEvent,
  verifyZeroAttendeeEvents,
  type CompactEvent,
} from "./delta.js";

/**
 * Unit tests for calendar hasNoAttendees candidate + verification helpers.
 * Run with: node dist/tools/delta.test.js
 */

class MockGraph {
  readonly requests: { path: string; query?: Record<string, unknown> }[] = [];
  private readonly byId = new Map<
    string,
    { attendees?: unknown; isCancelled?: boolean } | GraphError
  >();

  setEvent(
    id: string,
    body: { attendees?: unknown; isCancelled?: boolean }
  ): void {
    this.byId.set(id, body);
  }

  setError(id: string, err: GraphError): void {
    this.byId.set(id, err);
  }

  async request<T = unknown>(opts: {
    path: string;
    query?: Record<string, unknown>;
  }): Promise<T> {
    this.requests.push({ path: opts.path, query: opts.query });
    const match = opts.path.match(/^\/me\/events\/([^/?]+)$/);
    if (!match) {
      throw new Error(`Unexpected path: ${opts.path}`);
    }
    const id = decodeURIComponent(match[1]);
    const stored = this.byId.get(id);
    if (stored instanceof GraphError) throw stored;
    if (!stored) {
      throw new GraphError("Not found", 404, { id });
    }
    return stored as T;
  }
}

function asGraph(mock: MockGraph): GraphClient {
  return mock as unknown as GraphClient;
}

function testSummarizeEventFlagsEmptyAttendees() {
  const omitted = summarizeEvent({
    id: "evt-omit",
    subject: "Weekly series instance",
    organizer: { emailAddress: { address: "organizer@example.com" } },
    start: { dateTime: "2026-09-11T15:00:00" },
    isAllDay: false,
  });
  assert.strictEqual(omitted.hasNoAttendees, true);
  assert.strictEqual(omitted.attendeeCount, 0);

  const empty = summarizeEvent({
    id: "evt-empty",
    subject: "Teams hold",
    attendees: [],
  });
  assert.strictEqual(empty.hasNoAttendees, true);
  assert.strictEqual(empty.attendeeCount, 0);

  const withInvitees = summarizeEvent({
    id: "evt-full",
    subject: "Staff meeting",
    attendees: [
      {
        emailAddress: { address: "invitee@example.com" },
        status: { response: "notResponded" },
      },
    ],
  });
  assert.strictEqual(
    withInvitees.hasNoAttendees,
    false,
    "Pending RSVP (notResponded) is still an attendee"
  );
  assert.strictEqual(withInvitees.attendeeCount, 1);

  console.log("✅ testSummarizeEventFlagsEmptyAttendees passed");
}

function testSummarizeEventSkipsCancelledAndRemoved() {
  assert.strictEqual(isCancelledEvent({ isCancelled: true }), true);
  assert.strictEqual(
    isCancelledEvent({ subject: "Canceled: Weekly sync" }),
    true
  );
  assert.strictEqual(
    isCancelledEvent({ subject: "Cancelled: Weekly sync" }),
    true
  );
  assert.strictEqual(isCancelledEvent({ subject: "Weekly sync" }), false);

  const cancelled = summarizeEvent({
    id: "evt-cancel",
    subject: "Canceled: Weekly sync",
    isCancelled: true,
    attendees: [],
  });
  assert.strictEqual(cancelled.hasNoAttendees, false);
  assert.strictEqual(cancelled.isCancelled, true);

  const prefixOnly = summarizeEvent({
    id: "evt-prefix",
    subject: "Canceled: leftover hold",
    attendees: [],
  });
  assert.strictEqual(prefixOnly.hasNoAttendees, false);
  assert.strictEqual(prefixOnly.isCancelled, true);

  const removed = summarizeEvent({
    id: "evt-deleted",
    "@removed": { reason: "deleted" },
  });
  assert.strictEqual(removed.hasNoAttendees, false);

  console.log("✅ testSummarizeEventSkipsCancelledAndRemoved passed");
}

async function testVerifyCorrectsDeltaFalsePositive() {
  const mock = new MockGraph();
  mock.setEvent("series-instance", {
    attendees: [
      {
        emailAddress: { address: "invitee@example.com" },
        status: { response: "notResponded" },
      },
      { emailAddress: { address: "second@example.com" } },
    ],
  });

  const events: CompactEvent[] = [
    summarizeEvent({
      id: "series-instance",
      subject: "Weekly series",
      // Delta omitted attendees even though invitees exist.
    }),
    summarizeEvent({
      id: "already-populated",
      subject: "1:1",
      attendees: [{ emailAddress: { address: "peer@example.com" } }],
    }),
  ];

  assert.strictEqual(events[0].hasNoAttendees, true);

  await verifyZeroAttendeeEvents(asGraph(mock), events);

  assert.strictEqual(events[0].hasNoAttendees, false);
  assert.strictEqual(events[0].attendeeCount, 2);
  assert.strictEqual(mock.requests.length, 1);
  assert.strictEqual(mock.requests[0].path, "/me/events/series-instance");
  assert.deepStrictEqual(mock.requests[0].query, {
    $select: "id,attendees,isCancelled",
  });

  console.log("✅ testVerifyCorrectsDeltaFalsePositive passed");
}

async function testVerifyKeepsTrueEmptyHoldsAndMissingEvents() {
  const mock = new MockGraph();
  mock.setEvent("real-empty-hold", { attendees: [] });
  mock.setError(
    "gone-event",
    new GraphError("Not Found", 404, { error: { code: "ErrorItemNotFound" } })
  );

  const events: CompactEvent[] = [
    summarizeEvent({ id: "real-empty-hold", subject: "Focus block", attendees: [] }),
    summarizeEvent({ id: "gone-event", subject: "Stale delta row" }),
  ];

  await verifyZeroAttendeeEvents(asGraph(mock), events);

  assert.strictEqual(events[0].hasNoAttendees, true);
  assert.strictEqual(events[0].attendeeCount, 0);
  assert.strictEqual(events[1].hasNoAttendees, true);
  assert.strictEqual(mock.requests.length, 2);

  console.log("✅ testVerifyKeepsTrueEmptyHoldsAndMissingEvents passed");
}

async function testVerifyClearsCancelledOnFullReadAndSkipsKnownCancelled() {
  const mock = new MockGraph();
  mock.setEvent("hidden-cancel", { attendees: [], isCancelled: true });

  const events: CompactEvent[] = [
    summarizeEvent({
      id: "hidden-cancel",
      subject: "Old hold",
      attendees: [],
    }),
    summarizeEvent({
      id: "already-canceled",
      subject: "Canceled: skip me",
      attendees: [],
    }),
  ];

  await verifyZeroAttendeeEvents(asGraph(mock), events);

  assert.strictEqual(events[0].hasNoAttendees, false);
  assert.strictEqual(events[0].isCancelled, true);
  assert.strictEqual(events[1].hasNoAttendees, false);
  assert.strictEqual(
    mock.requests.length,
    1,
    "Must not re-fetch events already recognized as cancelled"
  );

  console.log("✅ testVerifyClearsCancelledOnFullReadAndSkipsKnownCancelled passed");
}

async function testVerifyDoesNotFetchPopulatedEvents() {
  const mock = new MockGraph();
  const events: CompactEvent[] = [
    summarizeEvent({
      id: "meeting",
      subject: "Standup",
      attendees: [{ emailAddress: { address: "a@example.com" } }],
    }),
  ];

  await verifyZeroAttendeeEvents(asGraph(mock), events);

  assert.strictEqual(mock.requests.length, 0);
  assert.strictEqual(events[0].hasNoAttendees, false);

  console.log("✅ testVerifyDoesNotFetchPopulatedEvents passed");
}

async function runTests() {
  console.log("Running delta hasNoAttendees tests...\n");

  try {
    testSummarizeEventFlagsEmptyAttendees();
    testSummarizeEventSkipsCancelledAndRemoved();
    await testVerifyCorrectsDeltaFalsePositive();
    await testVerifyKeepsTrueEmptyHoldsAndMissingEvents();
    await testVerifyClearsCancelledOnFullReadAndSkipsKnownCancelled();
    await testVerifyDoesNotFetchPopulatedEvents();

    console.log("\n✅ All tests passed!");
  } catch (err) {
    console.error("\n❌ Test failed:", err);
    process.exit(1);
  }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  runTests();
}
