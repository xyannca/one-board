import { NextResponse } from "next/server";

export const runtime = "nodejs";

// Place this file at: app/api/ai/summary/route.ts

type SummaryInput = {
  range: string;
  activeUsers: number;
  sessions: number;
  engagementRate: number; // 0–1
  topChannel: string;
  topChannelSharePct: number | null;
  topCountry: string | null;
  topCountrySharePct: number | null;
  notFoundCount: number;
};

export async function POST(request: Request) {
  const apiKey = process.env.ANTHROPIC_API_KEY;

  // Honest fallback: no key configured, so nothing is generated —
  // the frontend shows a clearly-labeled "connect an API key" state instead
  // of fabricated bullets.
  if (!apiKey) {
    return NextResponse.json({ available: false });
  }

  let input: SummaryInput;
  try {
    input = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid request body" }, { status: 400 });
  }

  const prompt = `You are writing exactly 3 short executive-summary bullet points for a live web analytics dashboard, based only on the real numbers below. Be concrete and specific, reference the actual figures given, and keep each bullet under 25 words. Do not invent numbers not listed here. Respond with ONLY a JSON array of exactly 3 strings — no markdown, no code fences, no preamble, no explanation.

Data (${input.range}):
- Active users: ${input.activeUsers}
- Sessions: ${input.sessions}
- Engagement rate: ${(input.engagementRate * 100).toFixed(0)}%
- Top channel: ${input.topChannel}${
    input.topChannelSharePct !== null ? ` (${input.topChannelSharePct.toFixed(0)}% of sessions)` : ""
  }
${
  input.topCountry
    ? `- Top country: ${input.topCountry}${
        input.topCountrySharePct !== null ? ` (${input.topCountrySharePct.toFixed(0)}% of active users)` : ""
      }`
    : ""
}
- Views hitting a missing page (404): ${input.notFoundCount}`;

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
        max_tokens: 2000,
        thinking: { type: "disabled" },
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const detail = await response.text();
      console.error("Anthropic API error:", response.status, detail);
      return NextResponse.json({ error: "AI summary request failed", detail }, { status: 502 });
    }

    const data = await response.json();
    const rawText: string =
      (data?.content || []).find((block: any) => block.type === "text")?.text ?? "";
    const cleaned = rawText.replace(/```json|```/g, "").trim();

    let bullets: string[];
    try {
      const parsed = JSON.parse(cleaned);
      bullets = Array.isArray(parsed) ? parsed.filter((b) => typeof b === "string") : [];
    } catch {
      // Fallback if the model didn't return clean JSON: split into lines.
      bullets = cleaned
        .split("\n")
        .map((line: string) => line.replace(/^[-•*\d.]+\s*/, "").trim())
        .filter(Boolean);
    }

    if (bullets.length === 0) {
      return NextResponse.json({ error: "AI summary returned no content" }, { status: 502 });
    }

    return NextResponse.json({ available: true, bullets });
  } catch (error: any) {
    console.error("AI summary generation error:", error);
    return NextResponse.json(
      { error: "Failed to generate summary", detail: error?.message },
      { status: 500 }
    );
  }
}
