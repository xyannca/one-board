import { NextResponse } from "next/server";
import type { ChartType, Aggregation } from "../../../types/kpi-spec";

export const runtime = "nodejs";

const CHART_TYPES: ChartType[] = ["kpi", "bar", "line", "pie"];
const AGGREGATIONS: Aggregation[] = ["sum", "avg", "count", "count_distinct", "min", "max", "latest"];

const SYSTEM_PROMPT = `You are looking at the structure of an uploaded spreadsheet and suggesting which numbers would make good KPI cards or charts for a business dashboard. You do NOT see the full dataset — only column names, an inferred type per column, a few sample values, and a small sample of rows. You are suggesting WHAT to compute and HOW; the actual numbers will be computed separately from the full dataset, never by you.

Suggest between 4 and 10 chart specs. Each is one of:
- A single KPI number: aggregation over one numeric column, no grouping (chartType "kpi")
- A breakdown by category: aggregation grouped by one category column (chartType "bar" or "pie")
- A trend over time: aggregation grouped by one date column (chartType "line")

Only reference column names actually given to you. When no column represents a measurable amount, aggregation "count" or "count_distinct" with metricColumn null is valid and often the right choice (e.g. counting rows per category).

Respond with ONLY a JSON array, no markdown, no preamble. Each entry:
{ "title": string, "metricColumn": string | null, "aggregation": "sum"|"avg"|"count"|"count_distinct"|"min"|"max"|"latest", "groupByColumn": string | null, "chartType": "kpi"|"bar"|"line"|"pie" }`;

export async function POST(request: Request) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return NextResponse.json({ available: false });

  let body: { columns?: unknown; profile?: unknown; sampleRows?: unknown };
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid request body" }, { status: 400 });
  }

  const { columns, profile, sampleRows } = body;
  if (!Array.isArray(columns) || !Array.isArray(profile) || !Array.isArray(sampleRows)) {
    return NextResponse.json(
      { error: "Request must include columns[], profile[], sampleRows[]" },
      { status: 400 }
    );
  }

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
        thinking: { type: "disabled" },
        system: SYSTEM_PROMPT,
        messages: [{ role: "user", content: JSON.stringify({ columns, profile, sampleRows }) }],
      }),
    });

    if (!response.ok) {
      const detail = await response.text();
      console.error("Anthropic API error:", response.status, detail);
      return NextResponse.json({ error: "KPI suggestion request failed", detail }, { status: 502 });
    }

    const data = await response.json();
    const rawText: string =
      (data?.content || []).find((b: any) => b.type === "text")?.text?.trim() ?? "";
    const firstBracket = rawText.indexOf("[");
    const lastBracket = rawText.lastIndexOf("]");
    const cleanedText =
      firstBracket !== -1 && lastBracket > firstBracket
        ? rawText.slice(firstBracket, lastBracket + 1)
        : rawText.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "").trim();

    let parsed: unknown;
    try {
      parsed = JSON.parse(cleanedText);
    } catch {
      console.error("KPI suggestion returned non-JSON:", rawText);
      return NextResponse.json({ error: "AI response was not valid JSON" }, { status: 502 });
    }
    if (!Array.isArray(parsed)) {
      return NextResponse.json({ error: "AI did not return an array" }, { status: 502 });
    }

    const columnSet = new Set(columns as string[]);
    const specs = (parsed as any[])
      .filter((s) => s && typeof s.title === "string")
      .filter((s) => s.metricColumn === null || columnSet.has(s.metricColumn))
      .filter((s) => s.groupByColumn === null || columnSet.has(s.groupByColumn))
      .filter((s) => AGGREGATIONS.includes(s.aggregation))
      .filter((s) => CHART_TYPES.includes(s.chartType))
      .slice(0, 10);

    return NextResponse.json({ available: true, specs });
  } catch (error: any) {
    console.error("KPI suggestion error:", error);
    return NextResponse.json({ error: "Failed to suggest KPIs", detail: error?.message }, { status: 500 });
  }
}
