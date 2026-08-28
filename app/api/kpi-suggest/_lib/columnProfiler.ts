// app/api/kpi-suggest/_lib/columnProfiler.ts
//
// Pure code, no AI call. Turns a fully-parsed raw dataset (any shape — not
// the old closed department/metricKey schema) into a per-column profile:
// inferred type, a few sample values, and a distinct-value count. This is
// what gets sent to /api/kpi-suggest instead of the rows themselves — the
// model never sees (and never needs) the full file to suggest what's worth
// computing.
//
// Importable from both server code and client code (page.tsx), same as
// app/api/ai-narrative/_lib/anomalies.ts — pure JS/TS, no server-only APIs.

import type { ColumnProfile } from "../../../../types/kpi-spec";

// Strips common formatting ($, commas, %, whitespace) before parsing — real
// spreadsheet exports routinely carry these on otherwise-numeric columns.
// Returns null (never NaN) so callers can tell "not a number" apart from
// "parsed to zero".
function toNumber(raw: string): number | null {
  const stripped = raw.trim().replace(/^[$€£]/, "").replace(/,/g, "").replace(/%$/, "");
  if (stripped === "") return null;
  const n = Number(stripped);
  return Number.isFinite(n) ? n : null;
}

function isNumericLike(raw: string): boolean {
  return toNumber(raw) !== null;
}

// Used only for columns already classified "number" — whether every one of
// its real values has no fractional part (e.g. salary, headcount) versus
// genuinely carrying decimals (e.g. a rate). Drives the aggregation
// engine's display precision for sum/avg — see ColumnProfile.isInteger.
function isIntegerLike(raw: string): boolean {
  const n = toNumber(raw);
  return n !== null && Number.isInteger(n);
}

// Deliberately stricter than a bare Date.parse: a lone numeric string like
// "2024" or "12" parses as a valid Date in most engines, which would
// misclassify plain numeric columns as dates. Requiring a date-shaped
// separator or a month name first avoids that — numeric columns are also
// checked before this function ever runs (see inferColumnType).
function isDateLike(raw: string): boolean {
  const value = raw.trim();
  const hasDateSeparator = /[-/]/.test(value);
  const hasMonthName = /^[A-Za-z]{3,9}\.?\s+\d{1,2}/.test(value);
  if (!hasDateSeparator && !hasMonthName) return false;
  return !Number.isNaN(Date.parse(value));
}

function toComparableString(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  const str = String(value).trim();
  return str === "" ? null : str;
}

// Order matters: number and date are checked first because a highly-unique
// numeric column (e.g. salary, where 30 random values are all likely
// distinct) must still read as "number", never "id" — "id" is reserved for
// non-numeric near-unique columns (employee codes, emails, names).
function inferColumnType(values: string[], totalRows: number): ColumnProfile["inferredType"] {
  if (values.length === 0) return "text";

  const numericCount = values.filter(isNumericLike).length;
  if (numericCount / values.length >= 0.9) return "number";

  const dateCount = values.filter(isDateLike).length;
  if (dateCount / values.length >= 0.9) return "date";

  const distinct = new Set(values.map((v) => v.toLowerCase()));
  if (distinct.size / values.length >= 0.9 && totalRows >= 8) return "id";

  if (distinct.size <= 30 || distinct.size / values.length <= 0.5) return "category";

  return "text";
}

export function profileColumns(rows: Record<string, unknown>[]): ColumnProfile[] {
  if (rows.length === 0) return [];

  // Union of keys across every row, not just the first — a sparse file
  // (some rows missing a trailing column) shouldn't silently drop that
  // column from the profile.
  const columnNames = new Set<string>();
  for (const row of rows) {
    for (const key of Object.keys(row)) columnNames.add(key);
  }

  return Array.from(columnNames).map((name) => {
    const values = rows
      .map((row) => toComparableString(row[name]))
      .filter((v): v is string => v !== null);

    const inferredType = inferColumnType(values, rows.length);
    const distinctValues = Array.from(new Set(values));

    const profile: ColumnProfile = {
      name,
      inferredType,
      sampleValues: distinctValues.slice(0, 5),
      distinctCount: distinctValues.length,
    };

    // Computed once, over the WHOLE column — not re-derived per aggregation
    // result — so every card that sums/averages this column agrees on
    // precision (an avg that happens to land on a whole number by
    // coincidence shouldn't get different rounding than one that doesn't).
    if (inferredType === "number") {
      profile.isInteger = values.length > 0 && values.every(isIntegerLike);
    }

    return profile;
  });
}

// Random sample of ~n rows (not the full file) to give the model a feel for
// real row shapes alongside the column profile. Order doesn't matter for
// this use — it's context for a suggestion call, never the data the
// suggested KPIs are actually computed from (that's the aggregation
// engine's job, over the full dataset).
export function sampleRows<T>(rows: T[], n = 25): T[] {
  if (rows.length <= n) return rows;
  const indices = new Set<number>();
  while (indices.size < n) {
    indices.add(Math.floor(Math.random() * rows.length));
  }
  return Array.from(indices)
    .sort((a, b) => a - b)
    .map((i) => rows[i]);
}
