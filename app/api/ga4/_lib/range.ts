// Place this file at: app/api/ga4/_lib/range.ts
// The `_lib` folder (underscore prefix) is excluded from Next.js routing,
// so it's safe to import from the three route.ts files as a shared helper.

export type Ga4Range = "7d" | "30d" | "90d";

/** Reads ?range=7d|30d|90d off the request URL. Defaults to 7d for any
 *  missing or unrecognized value, so existing calls without the param
 *  keep working exactly as before. */
export function parseRange(requestUrl: string): Ga4Range {
  const { searchParams } = new URL(requestUrl);
  const raw = searchParams.get("range");
  if (raw === "30d" || raw === "90d") return raw;
  return "7d";
}

/** Converts a range into the GA4 relative-date string used in dateRanges. */
export function rangeToStartDate(range: Ga4Range): string {
  if (range === "30d") return "30daysAgo";
  if (range === "90d") return "90daysAgo";
  return "7daysAgo";
}
