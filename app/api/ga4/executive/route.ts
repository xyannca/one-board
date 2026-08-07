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

    // Trend: active users + sessions, by date
    const [trendReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "date" }],
      metrics: [{ name: "activeUsers" }, { name: "sessions" }],
      orderBys: [{ dimension: { dimensionName: "date" } }],
    });

    // Traffic by channel
    const [channelReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "sessionDefaultChannelGroup" }],
      metrics: [{ name: "sessions" }],
      orderBys: [{ metric: { metricName: "sessions" }, desc: true }],
    });

    // Headline totals for the current period
    const [summaryReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      metrics: [
        { name: "activeUsers" },
        { name: "sessions" },
        { name: "engagementRate" },
      ],
    });

    // Period-over-period comparison: same-length window immediately before
    // the current one (7d vs prior 7d, 30d vs prior 30d, 90d vs prior 90d) —
    // replaces the old "first day vs last day of the trend chart" delta,
    // which understated real change on longer ranges since it ignored
    // everything that happened in between the two endpoints.
    const windows = getComparisonWindows(range);
    const [previousSummaryReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate: windows.previous.startDate, endDate: windows.previous.endDate }],
      metrics: [{ name: "activeUsers" }, { name: "sessions" }],
    });

    // Active users by country
    const [countryReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "country" }],
      metrics: [{ name: "activeUsers" }],
      orderBys: [{ metric: { metricName: "activeUsers" }, desc: true }],
      limit: 10,
    });

    // Active users by city
    const [cityReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "city" }],
      metrics: [{ name: "activeUsers" }],
      orderBys: [{ metric: { metricName: "activeUsers" }, desc: true }],
      limit: 10,
    });

    // New vs returning users
    const [newVsReturningReport] = await client.runReport({
      property: `properties/${propertyId}`,
      dateRanges: [{ startDate, endDate: "today" }],
      dimensions: [{ name: "newVsReturning" }],
      metrics: [{ name: "activeUsers" }],
      orderBys: [{ metric: { metricName: "activeUsers" }, desc: true }],
    });

    const trend = (trendReport.rows || []).map((row) => ({
      date: row.dimensionValues?.[0]?.value,
      activeUsers: Number(row.metricValues?.[0]?.value || 0),
      sessions: Number(row.metricValues?.[1]?.value || 0),
    }));

    const channels = (channelReport.rows || []).map((row) => ({
      channel: row.dimensionValues?.[0]?.value,
      sessions: Number(row.metricValues?.[0]?.value || 0),
    }));

    const summaryRow = summaryReport.rows?.[0];
    const summary = {
      activeUsers: Number(summaryRow?.metricValues?.[0]?.value || 0),
      sessions: Number(summaryRow?.metricValues?.[1]?.value || 0),
      engagementRate: Number(summaryRow?.metricValues?.[2]?.value || 0),
      topChannel: channels[0]?.channel || "\u2014",
    };

    const previousSummaryRow = previousSummaryReport.rows?.[0];
    const previousSummary = {
      activeUsers: Number(previousSummaryRow?.metricValues?.[0]?.value || 0),
      sessions: Number(previousSummaryRow?.metricValues?.[1]?.value || 0),
    };

    const deltas = {
      activeUsers: percentDelta(summary.activeUsers, previousSummary.activeUsers),
      sessions: percentDelta(summary.sessions, previousSummary.sessions),
    };

    const geoCountries = (countryReport.rows || []).map((row) => ({
      country: row.dimensionValues?.[0]?.value,
      activeUsers: Number(row.metricValues?.[0]?.value || 0),
    }));

    const geoCities = (cityReport.rows || []).map((row) => ({
      city: row.dimensionValues?.[0]?.value,
      activeUsers: Number(row.metricValues?.[0]?.value || 0),
    }));

    const newVsReturning = (newVsReturningReport.rows || []).map((row) => ({
      segment: row.dimensionValues?.[0]?.value,
      activeUsers: Number(row.metricValues?.[0]?.value || 0),
    }));

    const queryMs = Date.now() - t0;

    return NextResponse.json({
      summary,
      trend,
      channels,
      geoCountries,
      geoCities,
      newVsReturning,
      deltas,
      queryMs,
      range,
    });
  } catch (error: any) {
    console.error("GA4 executive report error:", error);
    return NextResponse.json(
      { error: "Failed to fetch GA4 data", detail: error?.message },
      { status: 500 }
    );
  }
}
