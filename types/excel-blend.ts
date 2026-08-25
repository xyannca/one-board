// Shared shapes for the Excel/CSV blend classifier — used by the API route
// (app/api/excel-classify/route.ts) and the dashboard UI (page.tsx). Same
// reasoning as types/ai-narrative.ts: importing from one place means a
// field added or renamed on one side is a compile error on the other,
// not a silent runtime mismatch.

export type ExcelDepartment = "marketing" | "operations" | "hr" | "executive";

// The only metricKeys the app can actually act on — the classifier must
// map to one of these or return null, never invent a new key. Marketing
// rows don't get a metricKey: they're matched by channel *name* against
// live GA4 data (see matchExcelChannelTargets in page.tsx), not by a
// fixed key, so metricKey is always null for department "marketing".
export type ExcelMetricKey =
  | "bounceRateTarget"
  | "avgEngagementDurationTargetSeconds"
  | "attritionTargetMax"
  | "complianceTargetMin"
  | "activeUsersTarget"
  | "sessionsTarget"
  | "engagementRateTarget";

export interface ExcelRowClassification {
  rowIndex: number;
  department: ExcelDepartment | null; // null = AI found no department this row belongs to
  metricKey: ExcelMetricKey | null;
}

export interface ExcelClassifyResponse {
  available: boolean; // false = no ANTHROPIC_API_KEY configured — mirrors /api/ai/summary
  classifications?: ExcelRowClassification[];
  error?: string;
}
