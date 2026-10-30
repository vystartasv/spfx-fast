import * as React from "react";
import { Button, Card, CardHeader, MessageBar, MessageBarBody, Spinner, Text } from "@fluentui/react-components";
import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;

function displayTime(value?: { dateTime: string; timeZone: string }, allDay?: boolean): string {
  if (!value) return "Time unavailable";
  if (allDay) return "All day";
  const hasOffset = /(?:Z|[+-]\d\d:\d\d)$/i.test(value.dateTime);
  const date = new Date(hasOffset ? value.dateTime : `${value.dateTime}Z`);
  return Number.isFinite(date.getTime()) ? date.toLocaleString() : "Time unavailable";
}

/** A cache-free, bounded agenda for the signed-in user's default Outlook calendar. */
export function CalendarAgenda({ client, days = 7 }: { client: Client; days?: number }): React.ReactElement {
  const [events, setEvents] = React.useState<Awaited<ReturnType<Client["getCalendarView"]>>["items"]>([]);
  const [cursor, setCursor] = React.useState<string>();
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState(false);

  const load = React.useCallback(async (signal?: AbortSignal, nextCursor?: string) => {
    setLoading(true);
    setError(false);
    try {
      const start = new Date();
      const end = new Date(start.getTime() + days * 86_400_000);
      const page = await client.getCalendarView(start.toISOString(), end.toISOString(), { maxItems: 100, cursor: nextCursor, signal });
      setEvents((previous) => nextCursor ? [...previous, ...page.items] : page.items);
      setCursor(page.nextCursor);
    } catch {
      if (!signal?.aborted) setError(true);
    } finally {
      if (!signal?.aborted) setLoading(false);
    }
  }, [client, days]);

  React.useEffect(() => {
    const controller = new AbortController();
    setEvents([]);
    load(controller.signal).catch(() => undefined);
    return () => controller.abort();
  }, [load]);

  return <section aria-label="Upcoming calendar events" style={{ display: "grid", gap: 12, maxWidth: 720 }}>
    <h2 style={{ margin: 0 }}>Upcoming events</h2>
    {loading && events.length === 0 ? <Spinner label="Loading events" /> : null}
    {error ? <MessageBar intent="error"><MessageBarBody>Calendar events could not be loaded. Check the Graph permission and try again.</MessageBarBody></MessageBar> : null}
    {!loading && !error && events.length === 0 ? <p>No events in this date range.</p> : null}
    <ol style={{ display: "grid", gap: 8, listStyle: "none", margin: 0, padding: 0 }}>
      {events.map((event) => <li key={event.id}>
        <Card>
          <CardHeader header={<Text weight="semibold">{event.subject || "Untitled event"}</Text>} description={<Text>{displayTime(event.start, event.isAllDay)}{event.end ? ` – ${displayTime(event.end, event.isAllDay)}` : ""}</Text>} />
          {event.location?.displayName ? <Text>{event.location.displayName}</Text> : null}
        </Card>
      </li>)}
    </ol>
    {cursor ? <Button disabled={loading} onClick={() => load(undefined, cursor).catch(() => undefined)}>{loading ? "Loading…" : "Load more events"}</Button> : null}
    <Button appearance="secondary" disabled={loading} onClick={() => load().catch(() => undefined)}>{loading ? "Loading…" : "Refresh"}</Button>
  </section>;
}
