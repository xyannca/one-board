import { NextResponse } from "next/server";
import { readFile } from "fs/promises";
import path from "path";
import { computeHRFacts, type HRDataset } from "./_lib/hrFacts";

export const runtime = "nodejs";

// Unlike the GA4 routes, this has no external API call — the "data source"
// is a static, versioned JSON file checked into the repo
// (public/sample-hr-dataset.json), generated once by a seeded script so the
// numbers are reproducible across page loads and deploys. queryMs is still
// reported for UI consistency with the GA4 routes, even though it will
// always be near-zero — this is client-side/static, not a live query, and
// the response says so explicitly via `source`.
export async function GET() {
  try {
    const t0 = Date.now();

    const filePath = path.join(process.cwd(), "public", "sample-hr-dataset.json");
    const raw = await readFile(filePath, "utf-8");
    const dataset: HRDataset = JSON.parse(raw);

    const facts = computeHRFacts(dataset);
    const queryMs = Date.now() - t0;

    return NextResponse.json({
      source: "sample" as const, // distinguishes from GA4 routes' implicit "live" source
      meta: dataset.meta,
      facts,
      monthlySnapshot: dataset.monthlySnapshot,
      companyTargets: dataset.companyTargets,
      queryMs,
    });
  } catch (error: any) {
    console.error("HR sample dataset error:", error);
    return NextResponse.json(
      { error: "Failed to load HR dataset", detail: error?.message },
      { status: 500 }
    );
  }
}
