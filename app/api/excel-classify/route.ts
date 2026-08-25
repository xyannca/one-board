// app/api/excel-classify/route.ts
//
// Classifies parsed Excel/CSV rows into the OneBoard department they
// belong to (Marketing/Operations/HR/Executive), or null if a row doesn't
// match anything the dashboard tracks. Runs ONCE at upload time — the
// result is persisted client-side (see ExcelState in page.tsx) and reused
// for every later render/export, never re-derived per report.
//
// Same pattern as app/api/ai-narrative/route.ts: raw fetch to the
// Anthropic API, strict-JSON system prompt, defensive markdown-fence
// stripping, honest `{ available: false }` fallback when no API key is
// configured (matches app/api/ai/summary/route.ts) rather than silently
// guessing or leaving the caller to infer failure from a generic error.

import { NextResponse } from "next/server";
import type {
  ExcelDepartment,
  ExcelMetricKey,
  ExcelRowClassification,
} from "../../../types/excel-blend";

export const runtime = "nodejs";

const DEPARTMENTS: ExcelDepartment[] = ["marketing", "operations", "hr", "executive"];
const METRIC_KEYS: ExcelMetricKey[] = [
  "bounceRateTarget",
  "avgEngagementDurationTargetSeconds",
  "attritionTargetMax",
  "complianceTargetMin",
  "activeUsersTarget",
  "sessionsTarget",
  "engagementRateTarget",
];

// Each metricKey only ever belongs to one department (see SYSTEM_PROMPT).
// Enforced here too, not just described to the model — a client-side
// lookup by metricKey alone (see findMetricTargetFromExcel in page.tsx)
// depends on this pairing actually holding, so a model mistake that
// pairs e.g. department:"executive" with metricKey:"attritionTargetMax"
// must be caught and discarded here, not trusted downstream.
const METRIC_KEY_DEPARTMENT: Record<ExcelMetricKey, ExcelDepartment> = {
  bounceRateTarget: "operations",
  avgEngagementDurationTargetSeconds: "operations",
  attritionTargetMax: "hr",
  complianceTargetMin: "hr",
  activeUsersTarget: "executive",
  sessionsTarget: "executive",
  engagementRateTarget: "executive",
};

// Rows beyond this are classified null locally, no model call — these are
// small baseline/target config files by design, not event logs.
const MAX_ROWS = 500;

const SYSTEM_PROMPT = `You classify rows from an uploaded Excel/CSV "baseline target" file for OneBoard, a dashboard with four tabs, each backed by its own real dataset:

- marketing: GA4 channel/traffic data. Its only known schema is columns "channel" + "target_sessions" — a per-channel session target (e.g. channel="Direct", target_sessions=10). A row with this shape is ALWAYS department "marketing", metricKey null (marketing rows are matched by channel name later, not by a metric key).
- operations: bounce rate and average engagement duration. Tracked metric keys:
  - "bounceRateTarget" — the row's metric is a bounce rate target (e.g. metric="bounce_rate")
  - "avgEngagementDurationTargetSeconds" — average engagement/session duration target, in seconds (e.g. metric="avg_engagement_duration_seconds")
- hr: workforce attrition and training compliance. Tracked metric keys:
  - "attritionTargetMax" — maximum acceptable attrition rate (e.g. metric="attrition_target")
  - "complianceTargetMin" — minimum acceptable training compliance rate (e.g. metric="compliance_target")
- executive: cross-department traffic summary. Tracked metric keys:
  - "activeUsersTarget" — active users target (e.g. metric="active_users")
  - "sessionsTarget" — sessions target (e.g. metric="sessions")
  - "engagementRateTarget" — engagement rate target (e.g. metric="engagement_rate")

Operations/HR/Executive rows typically use columns "metric" + "target_value" (or similarly named target/value columns) — the "metric" column's text is what tells you which department and which metricKey a row belongs to.

You will be given a JSON object: { "columns": string[], "rows": (object)[] }. Rows are 0-indexed in the order given.

Classify EACH row independently. A file may mix rows from multiple departments, contain rows that don't belong to any tracked metric, or be entirely unrelated to all four departments — handle all of these the same way, row by row, based only on that row's own content (column names and values). Never use any context outside the given data.

Rules:
1. Every row you classify into a department (other than marketing) MUST get a metricKey from the closed list above. If a row's metric text doesn't clearly and specifically match one of the described tracked fields, its department AND metricKey must both be null. Never guess, never invent a new metricKey, never force a weak match.
2. A row shaped like the marketing schema (channel + target_sessions style columns) is always department "marketing" with metricKey null.
3. Output ONLY a JSON array, one entry per input row, in the SAME ORDER as the input rows — no markdown fences, no preamble, no explanation. Each entry: { "department": "marketing"|"operations"|"hr"|"executive"|null, "metricKey": <one of the closed metricKey strings> | null }.
4. The output array length MUST exactly equal the number of input rows.`;

function isValidDepartment(v: unknown): v is ExcelDepartment {
  return typeof v === "string" && (DEPARTMENTS as string[]).includes(v);
}
function isValidMetricKey(v: unknown): v is ExcelMetricKey {
  return typeof v === "string" && (METRIC_KEYS as string[]).includes(v);
}

export async function POST(request: Request) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return NextResponse.json({ available: false });
  }

  let body: { columns?: unknown; rows?: unknown };
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid request body" }, { status: 400 });
  }

  const columns = Array.isArray(body.columns) ? body.columns : null;
  const rows = Array.isArray(body.rows) ? body.rows : null;
  if (!columns || !rows) {
    return NextResponse.json(
      { error: "Request must include columns[] and rows[]" },
      { status: 400 }
    );
  }
  if (rows.length === 0) {
    return NextResponse.json({ available: true, classifications: [] });
  }

  // Rows beyond MAX_ROWS are classified null without a model call — never
  // silently truncated data pretending to be a full classification.
  const modelRows = rows.slice(0, MAX_ROWS);

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 4000,
        system: SYSTEM_PROMPT,
        messages: [{ role: "user", content: JSON.stringify({ columns, rows: modelRows }) }],
      }),
    });

    if (!response.ok) {
      const detail = await response.text();
      console.error("Anthropic API error:", response.status, detail);
      return NextResponse.json(
        { error: "Excel classification request failed", detail },
        { status: 502 }
      );
    }

    const data = await response.json();
    const rawText: string =
      (data?.content || []).find((block: any) => block.type === "text")?.text?.trim() ?? "";
    const cleanedText = rawText
      .replace(/^```(?:json)?\s*/i, "")
      .replace(/\s*```$/i, "")
      .trim();

    let parsed: unknown;
    try {
      parsed = JSON.parse(cleanedText);
    } catch {
      console.error("Excel classification returned non-JSON:", rawText);
      return NextResponse.json(
        { error: "AI response was not valid JSON" },
        { status: 502 }
      );
    }

    // Strict validation: length must match, every entry must use only the
    // closed department/metricKey enums, and every department (other than
    // marketing) must carry a metricKey — a partially-broken mapping is
    // treated as a full failure, never accepted as partial classification.
    if (!Array.isArray(parsed) || parsed.length !== modelRows.length) {
      console.error("Excel classification array length mismatch:", parsed);
      return NextResponse.json(
        { error: "AI classification did not match the number of rows sent" },
        { status: 502 }
      );
    }

    const classifications: ExcelRowClassification[] = [];
    for (let i = 0; i < parsed.length; i += 1) {
      const entry = parsed[i] as { department?: unknown; metricKey?: unknown };
      const department = entry?.department === null ? null : entry?.department;
      const metricKey = entry?.metricKey === null ? null : entry?.metricKey;

      if (department !== null && !isValidDepartment(department)) {
        return NextResponse.json(
          { error: `AI returned an unrecognized department at row ${i}` },
          { status: 502 }
        );
      }
      if (metricKey !== null && !isValidMetricKey(metricKey)) {
        return NextResponse.json(
          { error: `AI returned an unrecognized metricKey at row ${i}` },
          { status: 502 }
        );
      }
      if (department !== null && department !== "marketing" && metricKey === null) {
        // A non-marketing department without a metricKey is meaningless
        // downstream — fall back to fully unclassified rather than keep a
        // half-formed classification.
        classifications.push({ rowIndex: i, department: null, metricKey: null });
        continue;
      }
      if (metricKey !== null && METRIC_KEY_DEPARTMENT[metricKey] !== department) {
        // Mismatched pairing (model error) — discard rather than trust a
        // metricKey that doesn't belong to the department it was paired
        // with.
        classifications.push({ rowIndex: i, department: null, metricKey: null });
        continue;
      }

      classifications.push({ rowIndex: i, department, metricKey });
    }

    // Rows beyond MAX_ROWS (never sent to the model) are explicitly
    // unclassified, not silently omitted.
    for (let i = modelRows.length; i < rows.length; i += 1) {
      classifications.push({ rowIndex: i, department: null, metricKey: null });
    }

    return NextResponse.json({ available: true, classifications });
  } catch (error: any) {
    console.error("Excel classification error:", error);
    return NextResponse.json(
      { error: "Failed to classify Excel/CSV data", detail: error?.message },
      { status: 500 }
    );
  }
}
