import { NextResponse } from "next/server";
import { BetaAnalyticsDataClient } from "@google-analytics/data";
import { parseRange, rangeToStartDate } from "../_lib/range";

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

    // Site health summary (also requests real quota usage from GA4 —
    // genuine tokensPerHour/tokensPerDay figures, not invented ones)
    const [summaryReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      metrics: [{ name: "bounceRate" }, { name: "averageSessionDuration" }],
      returnPropertyQuota: true,
    });

    // Pages with "404" in the title
    const [notFoundReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "pageTitle" }],
      metrics: [{ name: "screenPageViews" }],
      dimensionFilter: {
        filter: {
          fieldName: "pageTitle",
          stringFilter: { matchType: "CONTAINS", value: "404" },
        },
      },
      orderBys: [{ metric: { metricName: "screenPageViews" }, desc: true }],
    });

    // Unfiltered top pages by views — this is the same "Top Content Pages"
    // list that used to live on the Marketing tab. It's shown here instead
    // since page-level views (including any 404 entries) are more of an
    // operational signal than a marketing one.
    const [pagesReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "pageTitle" }],
      metrics: [{ name: "screenPageViews" }],
      orderBys: [{ metric: { metricName: "screenPageViews" }, desc: true }],
      limit: 10,
    });

    // Device category (desktop / mobile / tablet)
    const [deviceReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "deviceCategory" }],
      metrics: [{ name: "sessions" }],
      orderBys: [{ metric: { metricName: "sessions" }, desc: true }],
    });

    // Browser
    const [browserReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "browser" }],
      metrics: [{ name: "sessions" }],
      orderBys: [{ metric: { metricName: "sessions" }, desc: true }],
      limit: 8,
    });

    // Operating system
    const [osReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "operatingSystem" }],
      metrics: [{ name: "sessions" }],
      orderBys: [{ metric: { metricName: "sessions" }, desc: true }],
      limit: 8,
    });

    const summaryRow = summaryReport.rows?.[0];
    const summary = {
      bounceRate: Number(summaryRow?.metricValues?.[0]?.value || 0),
      avgSessionDuration: Number(summaryRow?.metricValues?.[1]?.value || 0),
    };

    const notFoundPages = (notFoundReport.rows || []).map((row) => ({
      title: row.dimensionValues?.[0]?.value,
      views: Number(row.metricValues?.[0]?.value || 0),
    }));

    const pages = (pagesReport.rows || []).map((row) => ({
      title: row.dimensionValues?.[0]?.value,
      views: Number(row.metricValues?.[0]?.value || 0),
    }));

    const devices = (deviceReport.rows || []).map((row) => ({
      device: row.dimensionValues?.[0]?.value,
      sessions: Number(row.metricValues?.[0]?.value || 0),
    }));

    const browsers = (browserReport.rows || []).map((row) => ({
      browser: row.dimensionValues?.[0]?.value,
      sessions: Number(row.metricValues?.[0]?.value || 0),
    }));

    const operatingSystems = (osReport.rows || []).map((row) => ({
      os: row.dimensionValues?.[0]?.value,
      sessions: Number(row.metricValues?.[0]?.value || 0),
    }));

    // Real GA4 API quota usage (not a fabricated "SLA" number)
    const tokensPerHour = summaryReport.propertyQuota?.tokensPerHour;
    const quota = tokensPerHour
      ? {
          consumed: Number(tokensPerHour.consumed || 0),
          remaining: Number(tokensPerHour.remaining || 0),
        }
      : null;

    const queryMs = Date.now() - t0;

    return NextResponse.json({
      summary,
      notFoundPages,
      pages,
      devices,
      browsers,
      operatingSystems,
      quota,
      queryMs,
      range,
    });
  } catch (error: any) {
    console.error("GA4 operations report error:", error);
    return NextResponse.json(
      { error: "Failed to fetch GA4 data", detail: error?.message },
      { status: 500 }
    );
  }
}
