// Shared shapes for the AI Narrative feature — used by the anomaly
// detection engine (anomalies.ts), the API route (route.ts), and the
// dashboard UI (page.tsx). Importing from one place means a field added
// or renamed on one side is a compile error on the others, instead of a
// silent runtime mismatch (this is what caused the earlier targets/context
// bug — the two sides drifted apart with no compiler warning).

export interface AiNarrativeFact {
  id: string;
  metric: string;
  value: number | string;
  unit?: string;
  // Matches BriefingFact["delta"] in components/ExecutiveBriefingPDF.tsx —
  // label is fully formatted by the caller (e.g. "0.7pp (< 8% Target)")
  // rather than raw value/period fields, so the two types stay assignable
  // to each other without an adapter.
  delta?: { direction: "up" | "down"; label: string };
  context?: string;
}

export interface AiAnomaly {
  id: string;
  type: "spike" | "drop" | "sustained_trend" | "target_breach" | "technical_issue";
  label: string;
  detail: string;
  severity: "info" | "warning";
}

export interface AiNarrativeResult {
  overview: string;
  keyObservations: string[];
  alerts: { label: string; detail: string }[];
  recommendations: string[];
}