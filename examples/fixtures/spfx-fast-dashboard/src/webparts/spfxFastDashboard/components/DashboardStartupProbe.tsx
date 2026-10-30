import * as React from "react";

export interface DashboardReadState {
  loading: boolean;
  source?: "cache" | "network" | "local";
}

interface ReadTiming {
  firstResultMs: number;
  firstSource?: DashboardReadState["source"];
  networkRefreshMs?: number;
}

/** Shows mount-to-first-result and mount-to-network-refresh timings for each list. */
export function DashboardStartupProbe({ startedAt, tasks, projects }: {
  startedAt: number;
  tasks: DashboardReadState;
  projects: DashboardReadState;
}): React.ReactElement {
  const [timings, setTimings] = React.useState<{ tasks?: ReadTiming; projects?: ReadTiming }>({});

  React.useEffect(() => {
    const elapsed = performance.now() - startedAt;
    const update = (current: ReadTiming | undefined, read: DashboardReadState): ReadTiming | undefined => {
      if (!read.source) return current;
      const result = current ?? { firstResultMs: elapsed, firstSource: read.source };
      return read.source === "network" && result.networkRefreshMs === undefined
        ? { ...result, networkRefreshMs: elapsed }
        : result;
    };
    setTimings((current) => ({
      tasks: update(current.tasks, tasks),
      projects: update(current.projects, projects),
    }));
  }, [projects, startedAt, tasks]);

  return <section aria-label="Dashboard startup timing">
    <h2>Startup timing</h2>
    <p>Measured in this browser from dashboard mount. First result shows its source; network refresh is measured separately.</p>
    <ul>
      <li>Tasks first result: {formatTiming(timings.tasks?.firstResultMs, timings.tasks?.firstSource)}</li>
      <li>Tasks network refresh: {formatTiming(timings.tasks?.networkRefreshMs)}</li>
      <li>Projects first result: {formatTiming(timings.projects?.firstResultMs, timings.projects?.firstSource)}</li>
      <li>Projects network refresh: {formatTiming(timings.projects?.networkRefreshMs)}</li>
    </ul>
  </section>;
}

function formatTiming(milliseconds?: number, source?: DashboardReadState["source"]): string {
  if (milliseconds === undefined) return "waiting";
  return `${milliseconds.toFixed(1)} ms${source ? ` (${source})` : ""}`;
}
