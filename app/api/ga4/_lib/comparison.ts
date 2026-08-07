// app/api/ga4/_lib/comparison.ts
//
// Computes two non-overlapping, equal-length date windows for
// period-over-period comparisons: the currently selected range, and the
// same-length range immediately preceding it (7d vs prior 7d, 30d vs prior
// 30d, 90d vs prior 90d). Uses explicit ISO dates rather than GA4's relative
// date shorthand ("7daysAgo" etc.) to avoid off-by-one ambiguity at the
// boundary between the two windows.

export type RangeKey = "7d" | "30d" | "90d";

const RANGE_DAYS: Record<RangeKey, number> = { "7d": 7, "30d": 30, "90d": 90 };

function toISODate(d: Date): string {
  return d.toISOString().slice(0, 10);
}

export function getComparisonWindows(range: string) {
  const days = RANGE_DAYS[range as RangeKey] ?? 30;

  const today = new Date();
  today.setUTCHours(0, 0, 0, 0);

  // Current window: the most recent `days` days, ending today (inclusive).
  const currentEnd = new Date(today);
  const currentStart = new Date(today);
  currentStart.setUTCDate(currentStart.getUTCDate() - (days - 1));

  // Previous window: the `days` days immediately before the current window
  // starts — no overlap, no gap.
  const previousEnd = new Date(currentStart);
  previousEnd.setUTCDate(previousEnd.getUTCDate() - 1);
  const previousStart = new Date(previousEnd);
  previousStart.setUTCDate(previousStart.getUTCDate() - (days - 1));

  return {
    current: { startDate: toISODate(currentStart), endDate: toISODate(currentEnd) },
    previous: { startDate: toISODate(previousStart), endDate: toISODate(previousEnd) },
  };
}

// Percent change from `previous` to `current`. Returns null when the
// previous period was zero and current isn't — "% change from zero" has no
// defined value, and returning null lets the UI show "—" instead of a
// misleading Infinity or 0.
export function percentDelta(current: number, previous: number): number | null {
  if (previous === 0) {
    return current === 0 ? 0 : null;
  }
  return ((current - previous) / previous) * 100;
}
