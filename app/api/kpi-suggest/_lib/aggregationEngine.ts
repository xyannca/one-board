// app/api/kpi-suggest/_lib/aggregationEngine.ts
//
// Pure code, no AI call. Given the FULL parsed dataset (never a sample) and
// one ChartSpec, computes the real number(s) it describes. This is the
// layer that keeps every card on the Custom Data tab traceable to an
// actual computed value — the AI only ever suggested WHAT to compute and
// HOW (see /api/kpi-suggest); it never sees or produces the numbers
// themselves.
//
// Importable from both server code and client code, same as
// app/api/ai-narrative/_lib/anomalies.ts.

import type { ChartSpec, ColumnProfile } from "../../../../types/kpi-spec";

export interface SeriesPoint {
  label: string;
  value: number;
}

export interface ChartResult {
  spec: ChartSpec;
  // Populated when spec.groupByColumn is null — a single KPI number (or,
  // for aggregation "latest" on a non-numeric column, its raw value).
  value: number | string | null;
  // Populated when spec.groupByColumn is set — one point per real group,
  // never a fabricated bucket.
  series: SeriesPoint[];
  isEmpty: boolean;
  // How many decimal places every number this result produced (the single
  // `value`, and every `series` point) should be displayed with — 2 only
  // for a sum/avg over a column that genuinely carries decimals, 0
  // otherwise. Computed once here, from the spec + source column, so every
  // surface that renders this result (KPI card, bar label, ...) agrees —
  // never re-derived per surface from the numeric result itself, which is
  // exactly what let e.g. 0.50 print as "0.5" in one place and not another.
  decimals: 0 | 2;
}

// Strips common formatting ($, commas, %) before parsing — real
// spreadsheet exports routinely carry these on otherwise-numeric columns.
function parseNumber(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
  const stripped = String(raw).trim().replace(/^[$€£]/, "").replace(/,/g, "").replace(/%$/, "");
  if (stripped === "") return null;
  const n = Number(stripped);
  return Number.isFinite(n) ? n : null;
}

function groupKey(raw: unknown): string | null {
  if (raw === null || raw === undefined) return null;
  const str = String(raw).trim();
  return str === "" ? null : str;
}

function parseDateMs(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;
  const ms = Date.parse(String(raw).trim());
  return Number.isNaN(ms) ? null : ms;
}

// True only when spec.metricColumn is a "number" column whose real values
// were ALL integers (see ColumnProfile.isInteger) — looked up once per
// call, from the column's own profile, never guessed from the computed
// result. Unknown (no profile, or column not found there) defaults to
// false: 2-decimal precision is the safer default because it never hides
// real fractional data, whereas defaulting to "integer" would silently
// round away digits that might be real.
function isMetricColumnInteger(spec: ChartSpec, profile?: ColumnProfile[]): boolean {
  if (!spec.metricColumn || !profile) return false;
  return profile.find((p) => p.name === spec.metricColumn)?.isInteger === true;
}

// True when spec.metricColumn was profiled as a "date" column — looked up
// once per call, same pattern as isMetricColumnInteger. Drives min/max:
// a date-typed column's earliest/latest value is a real date, not a
// number, and must be compared/formatted as one.
function isMetricColumnDate(spec: ChartSpec, profile?: ColumnProfile[]): boolean {
  if (!spec.metricColumn || !profile) return false;
  return profile.find((p) => p.name === spec.metricColumn)?.inferredType === "date";
}

// Fixed locale (not the server/browser's ambient locale) so the same date
// always renders the same way regardless of where this runs — deterministic
// output, same reasoning as everything else this file computes.
const READABLE_DATE_FORMAT = new Intl.DateTimeFormat("en-US", {
  year: "numeric",
  month: "short",
  day: "numeric",
});

function formatReadableDate(ms: number): string {
  return READABLE_DATE_FORMAT.format(new Date(ms));
}

// Only sum/avg can introduce fractional output from a non-integer column
// (division, or float accumulation) — count/count_distinct are always
// whole numbers, and min/max/latest return a value already present
// verbatim in the data, whatever precision it naturally has. So 2-decimal
// display only ever applies to a sum/avg over a column that isn't
// integer-typed.
function decimalsForSpec(spec: ChartSpec, metricIsInteger: boolean): 0 | 2 {
  const isRoundedAggregation = spec.aggregation === "sum" || spec.aggregation === "avg";
  return isRoundedAggregation && !metricIsInteger ? 2 : 0;
}

// Rounds a sum/avg result to the precision implied by its SOURCE column:
// a whole number for integer-typed columns (e.g. salary), 2 decimal places
// otherwise (e.g. a rate like 0.23) — never a hardcoded "always round to
// integer" regardless of what the underlying data actually looks like.
function roundByColumnPrecision(value: number, isInteger: boolean): number {
  return isInteger ? Math.round(value) : Math.round(value * 100) / 100;
}

// Collapses a real date string down to "YYYY-MM" — used to bucket a line
// chart's date groups to a readable grain instead of one point per exact
// raw date value. Prefers a direct ISO regex match (correct regardless of
// timezone) and only falls back to Date.parse for non-ISO formats.
function toMonthBucket(raw: string): string {
  const iso = raw.match(/^(\d{4})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}`;
  const ms = parseDateMs(raw);
  if (ms === null) return raw; // unreachable — callers only bucket already-parseable dates
  const d = new Date(ms);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

// Same idea as toMonthBucket, one grain coarser — used when month buckets
// alone would still produce more than MAX_LINE_BUCKETS points.
function toYearBucket(raw: string): string {
  const iso = raw.match(/^(\d{4})-/);
  if (iso) return iso[1];
  const ms = parseDateMs(raw);
  if (ms === null) return raw;
  return String(new Date(ms).getFullYear());
}

// Reduces one group of rows to a single number (or, for "latest" on a
// non-numeric column, a raw value) per the spec's aggregation +
// metricColumn. Returns null when there's nothing real to report — e.g.
// every value in the group was non-numeric — never a guessed 0.
function aggregateRows(
  rows: Record<string, unknown>[],
  spec: ChartSpec,
  metricIsInteger: boolean,
  metricIsDate: boolean
): number | string | null {
  const { aggregation, metricColumn } = spec;

  if (aggregation === "count") {
    return rows.length;
  }

  if (aggregation === "count_distinct") {
    // With no metricColumn, "distinct" has nothing to count over — falls
    // back to a plain row count (e.g. "rows per category"), same case the
    // /api/kpi-suggest system prompt calls out as valid.
    if (!metricColumn) return rows.length;
    const distinct = new Set(
      rows.map((r) => groupKey(r[metricColumn])).filter((v): v is string => v !== null)
    );
    return distinct.size;
  }

  if (!metricColumn) return null; // sum/avg/min/max/latest all need a real column

  if (aggregation === "latest") {
    // ChartSpec carries no explicit date-column hint, so "latest" scans
    // every field of every row in this group for the most recent
    // date-parseable value and reads metricColumn off that row. Falls back
    // to the last row in file order when no date field is found anywhere.
    let bestRow: Record<string, unknown> | null = null;
    let bestMs = -Infinity;
    let anyDateFound = false;
    for (const row of rows) {
      for (const key of Object.keys(row)) {
        const ms = parseDateMs(row[key]);
        if (ms !== null) {
          anyDateFound = true;
          if (ms >= bestMs) {
            bestMs = ms;
            bestRow = row;
          }
        }
      }
    }
    const chosen = anyDateFound ? bestRow : (rows[rows.length - 1] ?? null);
    if (!chosen) return null;
    const raw = chosen[metricColumn];
    const num = parseNumber(raw);
    return num !== null ? num : groupKey(raw);
  }

  if (aggregation === "min" || aggregation === "max") {
    // A date-typed column's earliest/latest value is a real date — compare
    // it as one (parseDateMs), not as a number (a date string like
    // "07/11/2022" isn't numeric at all, which is exactly why this used to
    // return null for every date column). The result is formatted into a
    // readable string, never the raw millisecond timestamp.
    if (metricIsDate) {
      const dateValues = rows.map((r) => parseDateMs(r[metricColumn])).filter((v): v is number => v !== null);
      if (dateValues.length === 0) return null;
      const ms = aggregation === "min" ? Math.min(...dateValues) : Math.max(...dateValues);
      return formatReadableDate(ms);
    }
    const values = rows.map((r) => parseNumber(r[metricColumn])).filter((v): v is number => v !== null);
    if (values.length === 0) return null;
    // Returns a value already present verbatim in the data — no rounding.
    return aggregation === "min" ? Math.min(...values) : Math.max(...values);
  }

  const values = rows.map((r) => parseNumber(r[metricColumn])).filter((v): v is number => v !== null);
  if (values.length === 0) return null;

  // sum/avg are the two aggregations that can introduce fractional output
  // even from an integer-typed column (division, or floating-point
  // accumulation) — precision is decided by the SOURCE column, not by
  // whether this particular result happens to land on a whole number.
  if (aggregation === "sum") return roundByColumnPrecision(values.reduce((a, b) => a + b, 0), metricIsInteger);
  if (aggregation === "avg") return roundByColumnPrecision(values.reduce((a, b) => a + b, 0) / values.length, metricIsInteger);
  return null;
}

const MAX_SERIES_POINTS = 12;
// Line charts bucket dates to "YYYY-MM"; if that alone would still produce
// more than this many points, they fall back one grain coarser to "YYYY".
const MAX_LINE_BUCKETS = 24;

export function computeChartSpec(
  rows: Record<string, unknown>[],
  spec: ChartSpec,
  profile?: ColumnProfile[]
): ChartResult {
  const metricIsInteger = isMetricColumnInteger(spec, profile);
  const metricIsDate = isMetricColumnDate(spec, profile);
  const decimals = decimalsForSpec(spec, metricIsInteger);

  if (!spec.groupByColumn) {
    const value = aggregateRows(rows, spec, metricIsInteger, metricIsDate);
    return { spec, value, series: [], isEmpty: value === null, decimals };
  }
  const groupByColumn = spec.groupByColumn;

  // Determine whether the groupBy column is date-valued BEFORE building
  // groups, so a line chart over a date column can bucket to a coarser
  // grain (month, or year if that's still too many points) instead of
  // grouping on each row's exact raw date value — which, when every row
  // has a distinct timestamp, degenerates into one real-but-useless point
  // per row (a flat "count 1" line the whole way across).
  const rawKeys = rows
    .map((r) => groupKey(r[groupByColumn]))
    .filter((k): k is string => k !== null);
  const isDateColumn = rawKeys.length > 0 && rawKeys.every((k) => parseDateMs(k) !== null);

  let bucket: ((raw: string) => string) | null = null;
  if (spec.chartType === "line" && isDateColumn) {
    const monthBucketCount = new Set(rawKeys.map(toMonthBucket)).size;
    bucket = monthBucketCount > MAX_LINE_BUCKETS ? toYearBucket : toMonthBucket;
  }

  // Real groups only — a row with no value in the group column contributes
  // to no bucket, never a guessed "Unknown" catch-all.
  const groups = new Map<string, Record<string, unknown>[]>();
  for (const row of rows) {
    const raw = groupKey(row[groupByColumn]);
    if (raw === null) continue;
    const key = bucket ? bucket(raw) : raw;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  const groupLabels = Array.from(groups.keys());
  const isDateGroup = bucket !== null || (groupLabels.length > 0 && groupLabels.every((k) => parseDateMs(k) !== null));

  let series: SeriesPoint[] = Array.from(groups.entries())
    .map(([label, groupRows]) => {
      const result = aggregateRows(groupRows, spec, metricIsInteger, metricIsDate);
      return { label, value: typeof result === "number" ? result : NaN };
    })
    .filter((p) => Number.isFinite(p.value));

  series = isDateGroup
    ? series.sort((a, b) => (parseDateMs(a.label) ?? 0) - (parseDateMs(b.label) ?? 0))
    : series.sort((a, b) => b.value - a.value);

  // Display cap for categorical breakdowns — every point shown is still a
  // real computed group; this just leaves off the longest tail of small
  // categories rather than rendering an unreadable chart. Chronological
  // (line) series are never capped this way.
  if (!isDateGroup && series.length > MAX_SERIES_POINTS) {
    series = series.slice(0, MAX_SERIES_POINTS);
  }

  return { spec, value: null, series, isEmpty: series.length === 0, decimals };
}
