export type ChartType = "kpi" | "bar" | "line" | "pie";
export type Aggregation = "sum" | "avg" | "count" | "count_distinct" | "min" | "max" | "latest";

export interface ChartSpec {
  title: string;
  metricColumn: string | null; // null only valid when aggregation is "count"
  aggregation: Aggregation;
  groupByColumn: string | null; // null = single KPI number, no grouping
  chartType: ChartType;
}

export interface ColumnProfile {
  name: string;
  inferredType: "number" | "date" | "category" | "text" | "id";
  sampleValues: string[];
  distinctCount?: number;
  // Only meaningful when inferredType is "number": true when every sampled
  // value in this column has no fractional part (e.g. salary, headcount).
  // The aggregation engine uses this to decide display precision for
  // sum/avg results — round to a whole number for integer-typed columns,
  // keep 2 decimal places for columns that are genuinely fractional (e.g.
  // a rate like 0.23) — instead of guessing precision from the computed
  // result alone.
  isInteger?: boolean;
}

export interface KpiSuggestResponse {
  available: boolean;
  specs?: ChartSpec[];
  error?: string;
}
