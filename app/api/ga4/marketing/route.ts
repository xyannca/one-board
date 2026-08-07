import { NextResponse } from "next/server";
import { BetaAnalyticsDataClient } from "@google-analytics/data";
import { parseRange, rangeToStartDate } from "../_lib/range";
import { getComparisonWindows, percentDelta } from "../_lib/comparison";

export const runtime = "nodejs";

function getClient() {
  return new BetaAnalyticsDataClient({
    credentials: {
      client_email: process.env.GA4_CLIENT_EMAIL,
      private_key: process.env.GA4_PRIVATE_KEY?.replace(/\\n/g, "\n"),
    },
  });
}

export async function GET(request: Request) {
  try {
    const propertyId = process.env.GA4_PROPERTY_ID;
    if (!propertyId) {
      return NextResponse.json({ error: "Missing GA4_PROPERTY_ID" }, { status: 500 });
    }

    const range = parseRange(request.url);
    const startDate = rangeToStartDate(range);
    const client = getClient();
    const t0 = Date.now();

    // Traffic source / medium breakdown
    const [sourceReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "sessionSourceMedium" }],
      metrics: [{ name: "sessions" }],
      orderBys: [{ metric: { metricName: "sessions" }, desc: true }],
      limit: 10,
    });

    // Top pages by views
    const [pagesReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "pageTitle" }],
      metrics: [{ name: "screenPageViews" }],
      orderBys: [{ metric: { metricName: "screenPageViews" }, desc: true }],
      limit: 10,
    });

    // Landing pages
    const [landingReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "landingPage" }],
      metrics: [{ name: "sessions" }],
      orderBys: [{ metric: { metricName: "sessions" }, desc: true }],
      limit: 10,
    });

    // Channel group breakdown (for the channel radar chart) — same dimension
    // the Executive route uses, kept here too since each tab route is self-contained.
    const [channelReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "sessionDefaultChannelGroup" }],
      metrics: [{ name: "sessions" }],
      orderBys: [{ metric: { metricName: "sessions" }, desc: true }],
    });

    // Funnel step 1 & 2: sessions -> engaged sessions (always available)
    const [funnelReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      metrics: [{ name: "sessions" }, { name: "engagedSessions" }],
    });

    // Period-over-period comparison for the Total Traffic Sessions KPI —
    // same-length window immediately preceding the current one.
    const windows = getComparisonWindows(range);
    const [previousSessionsReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate: windows.previous.startDate, endDate: windows.previous.endDate }],
      metrics: [{ name: "sessions" }],
    });

    // Funnel step 3: key events (formerly "conversions"). Not every property
    // has these configured, so this query is isolated — if it fails for any
    // reason, the funnel just stays at 2 steps instead of breaking the route.
    let keyEvents: number | null = null;
    try {
      const [keyEventsReport] = await client.runReport({
        property: `properties/${propertyId}`,
        dateRanges: [{ startDate, endDate: "today" }],
        metrics: [{ name: "keyEvents" }],
      });
      keyEvents = Number(keyEventsReport.rows?.[0]?.metricValues?.[0]?.value || 0);
    } catch (keyEventsError) {
      console.warn("Key events metric unavailable, skipping funnel step 3:", keyEventsError);
      keyEvents = null;
    }

    const sources = (sourceReport.rows || []).map((row) => ({
      source: row.dimensionValues?.[0]?.value,
      sessions: Number(row.metricValues?.[0]?.value || 0),
    }));

    const pages = (pagesReport.rows || []).map((row) => ({
      title: row.dimensionValues?.[0]?.value,
      views: Number(row.metricValues?.[0]?.value || 0),
    }));

    const landingPages = (landingReport.rows || []).map((row) => ({
      page: row.dimensionValues?.[0]?.value,
      sessions: Number(row.metricValues?.[0]?.value || 0),
    }));

    const channels = (channelReport.rows || []).map((row) => ({
      channel: row.dimensionValues?.[0]?.value,
      sessions: Number(row.metricValues?.[0]?.value || 0),
    }));

    const funnelRow = funnelReport.rows?.[0];
    const funnel = {
      sessions: Number(funnelRow?.metricValues?.[0]?.value || 0),
      engagedSessions: Number(funnelRow?.metricValues?.[1]?.value || 0),
      keyEvents,
    };

    const previousSessions = Number(previousSessionsReport.rows?.[0]?.metricValues?.[0]?.value || 0);
    const deltas = {
      sessions: percentDelta(funnel.sessions, previousSessions),
    };

    const queryMs = Date.now() - t0;

    return NextResponse.json({ sources, pages, landingPages, channels, funnel, deltas, queryMs, range });
  } catch (error: any) {
    console.error("GA4 marketing report error:", error);
    return NextResponse.json(
      { error: "Failed to fetch GA4 data", detail: error?.message },
      { status: 500 }
    );
  }
}
