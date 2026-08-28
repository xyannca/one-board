// app/api/ai-narrative/route.ts
//
// Structured version. Input now includes optional `anomalies` (pre-computed
// by anomalies.ts — rule-based, not AI) and optional `targets` (only present
// when the tab has real internal target values, e.g. HR's companyTargets).
//
// The model returns a JSON object with four sections. Hard constraints:
// - `alerts` may ONLY restate items from the `anomalies` array it was given.
//   If `anomalies` is empty, `alerts` must be an empty array — never invent one.
// - `recommendations` may ONLY appear if `targets` was provided AND the
//   recommendation shows its arithmetic using only given numbers. If no
//   targets were given, `recommendations` must be an empty array.
// - `keyObservations` may reference facts/trend numbers but must never
//   state or imply causation between two facts unless anomalies links them.

import { NextResponse } from "next/server";
import type { AiNarrativeFact, AiAnomaly } from "../../../types/ai-narrative";

export const runtime = "nodejs";

const SYSTEM_PROMPT = `You are OneBoard's data narrative assistant. You turn pre-computed data into a structured executive summary. Output ONLY valid JSON matching this exact shape, nothing else — no markdown fences, no preamble:

{
  "overview": string,
  "keyObservations": string[],
  "alerts": [{ "label": string, "detail": string }],
  "recommendations": string[]
}

Strict rules — violating any of these is a failure:
1. Only reference numbers and fields present in the input. Never invent, estimate, or round to a number not present.
2. "overview": 1-2 plain declarative sentences framing the period. No hedging words ("may", "might", "could be due to").
3. "keyObservations": 2-4 sentences, each stating a fact or a stat from the input. Never claim two facts are related or that one caused another unless the input's anomalies array explicitly links them.
4. "alerts": populate this ONLY by rephrasing entries from the input's "anomalies" array, one alert per anomaly, using the anomaly's own "detail" field as the source of truth — do not add severity judgments or causes beyond what "detail" states. If the input has no anomalies (empty or missing array), "alerts" MUST be an empty array. Never invent an alert.
5. "recommendations": populate this under EITHER of two conditions, never otherwise:
   (a) the input includes a "context" object with usable numeric or ranking data — each such recommendation must show its arithmetic or ranking reasoning using only numbers present in the input (e.g. a gap against a target, a capacity-vs-demand calculation, or an ordering by magnitude).
   (b) the input's "anomalies" array contains an entry with "type": "technical_issue" — for these ONLY, you may offer brief, standard, well-established troubleshooting guidance for that category of technical problem (e.g. typical causes and checks for a broken route: mistyped URL, deleted page, missing redirect, stale external link). Stay generic — never claim to know the specific cause on THIS site, never invent a number not present in the input, and never extend this allowance to business/financial/HR judgment calls.
   If neither (a) nor (b) applies, "recommendations" MUST be an empty array — do not write generic business advice.
6. Never introduce outside information: no industry benchmarks, no assumptions about the business or sector.
7. Output must be parseable JSON. Do not wrap it in markdown code fences.`;

export async function POST(request: Request) {
  try {
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      return NextResponse.json({ error: "Missing ANTHROPIC_API_KEY" }, { status: 500 });
    }

    const body = await request.json();
    const facts: AiNarrativeFact[] = body?.facts;
    const anomalies: AiAnomaly[] = body?.anomalies ?? [];
    const context: Record<string, number | string> | undefined = body?.context;

    if (!Array.isArray(facts) || facts.length === 0) {
      return NextResponse.json({ error: "Request must include a non-empty facts array" }, { status: 400 });
    }

    const t0 = Date.now();

    const userPayload = { facts, anomalies, context: context ?? null };

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
        messages: [{ role: "user", content: JSON.stringify(userPayload) }],
      }),
    });

    if (!response.ok) {
      const detail = await response.text();
      console.error("Anthropic API error:", response.status, detail);
      return NextResponse.json(
        { error: "AI narrative generation failed", detail },
        { status: response.status }
      );
    }

    const data = await response.json();
    const rawText: string = data.content?.find((block: any) => block.type === "text")?.text?.trim() ?? "";

    // Defensive cleanup: the model is instructed not to wrap output in markdown
    // code fences, but occasionally does anyway. Strip them before parsing
    // rather than trusting the instruction alone.
    const cleanedText = rawText
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();

    let parsed;
    try {
    parsed = JSON.parse(cleanedText);
    } catch {


      console.error("AI narrative returned non-JSON:", rawText);
      return NextResponse.json({ error: "AI response was not valid JSON", raw: cleanedText}, { status: 502 });
    }

    // Defensive guard, mirroring the system prompt's own rules: if the
    // caller supplied no anomalies, force alerts empty regardless of what
    // the model returned. Same for recommendations without targets.
    if (anomalies.length === 0) parsed.alerts = [];
    const hasTechnicalIssue = anomalies.some((a) => a.type === "technical_issue");
    if (!context && !hasTechnicalIssue) parsed.recommendations = [];

    const queryMs = Date.now() - t0;

    return NextResponse.json({ ...parsed, queryMs });
  } catch (error: any) {
    console.error("AI narrative route error:", error);
    return NextResponse.json(
      { error: "Failed to generate AI narrative", detail: error?.message },
      { status: 500 }
    );
  }
}
