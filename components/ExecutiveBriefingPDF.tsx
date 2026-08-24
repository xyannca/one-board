// components/ExecutiveBriefingPDF.tsx
//
// One-click Executive Briefing generator. Structure and color rules are
// defined in executive-briefing-design-spec.md (see CLAUDE.md) — this file
// implements that spec, it does not redefine it.
//
// Core discipline, unchanged from the original version of this file:
// every field rendered here must trace back to a value already computed
// and displayed on-screen before export. This component does not compute
// or fetch anything new — it only lays out data it's given. That includes
// `status` on each KPI/risk: the CALLER decides whether a number is
// critical/watch/ok (based on real thresholds), this component only
// renders whatever status it's handed. It must never infer status from
// direction alone (e.g. "up = good") — that's wrong for metrics like
// attrition where up is bad, and was a real bug in the previous version.

import type { Style } from "@react-pdf/types";
import { Document, Page, Text, View, StyleSheet, Svg, Path, Font } from "@react-pdf/renderer";

Font.register({
  family: "Inter",
  fonts: [
    { src: "/fonts/Inter-Regular.woff", fontWeight: 400 },
    { src: "/fonts/Inter-Bold.woff", fontWeight: 700 },
    { src: "/fonts/Inter-Bold.woff", fontWeight: 800 },
  ],
});

Font.registerHyphenationCallback((word) => [word]);

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type Status = "critical" | "watch" | "ok" | "neutral";

export interface BriefingFact {
  id: string;
  metric: string;
  icon?: string;           // small glyph, e.g. "👥" — optional, purely decorative
  value: number | string;
  unit?: string;
  status?: Status;         // defaults to "neutral" if omitted — caller must set
                            // this deliberately per real threshold logic, not
                            // inferred from delta direction
  statusLabel?: string;    // colored pill text, e.g. "Rising · Watch", "Below Target",
                            // "Stable" — a real status judgment. NOT for plain
                            // descriptive text that isn't a status (see `note`).
  note?: string;           // plain caption below the value/delta, no pill/color —
                            // for descriptive-but-not-a-status content, e.g. a
                            // share percentage ("62% traffic share") on a KPI
                            // that has no real target to be on/off track against
  delta?: {
    direction: "up" | "down";
    label: string;         // fully formatted by caller, e.g. "0.7pp (< 8% Target)"
  };
}

export interface BriefingRisk {
  label: string;
  detail: string;          // must restate a fact already shown — no new claims
  status: Status;           // MUST match the status of the KPI this describes —
                             // this is the color-consistency rule from the spec
}

export interface ComparisonChart {
  title: string;
  subtitle?: string;
  points: { label: string; value: number; display: string }[]; // real readings only —
                                                                  // never interpolate
                                                                  // points you don't have
  footnote?: string;
  gauge?: {
    label: string;
    value: number;          // 0-100
    target: number;         // 0-100, drawn as a marker line
    display: string;
    footnote?: string;
  };
}

export interface DrilldownTable {
  // Only populate this if the underlying dataset genuinely has this
  // dimension (e.g. GA4's real device/browser breakdown). Do not
  // construct a table for a dimension that isn't actually in the data —
  // see spec §"Per-tab drill-down data status".
  title: string;
  columns: string[];
  rows: { cells: (string | number)[]; status: Status }[];
}

export interface BriefingData {
  reportTitle: string;      // e.g. "HR & Workforce Performance"
  sectionLabel: string;     // e.g. "HR View"
  audience?: string;        // defaults to "Executive Leadership" — describes the
                             // reader, not the topic, so it normally doesn't vary by tab
  dataSource: string;
  period: string;
  generatedAt: string;
  kpis: BriefingFact[];
  narrative: string;        // supports **bold** inline markup for the real
                             // conclusion clause(s) — see RichText below
  chart?: ComparisonChart;
  risks?: BriefingRisk[];
  drilldown?: DrilldownTable;
  keyObservations?: string[];
  recommendations?: string[]; // matches AiNarrativeResult.recommendations exactly —
                               // these are already-complete sentences from the AI,
                               // not a title+detail pair
}

// ---------------------------------------------------------------------------
// Status → color mapping (single source of truth — every colored element
// in this file reads from here, nothing hardcodes a status color separately)
// ---------------------------------------------------------------------------

const STATUS: Record<Status, { text: string; pillBg: string; rowBg: string }> = {
  critical: { text: "#b03a2e", pillBg: "#fdecea", rowBg: "#fdf1ef" },
  watch:    { text: "#a06c00", pillBg: "#fdf3e0", rowBg: "#faf6ee" },
  ok:       { text: "#0f7a44", pillBg: "#e8f5ee", rowBg: "#f0f8f3" },
  neutral:  { text: "#45505f", pillBg: "#eef0f3", rowBg: "#f7f8f9" },
};

const INK = "#10151f";
const INK_SOFT = "#5a6472";
const BORDER = "#c9ced6";
const BRAND = "#0f9d68";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const styles = StyleSheet.create({
  page: { paddingHorizontal: 56, paddingTop: 56, paddingBottom: 60, fontFamily: "Inter", fontSize: 10.5, color: INK },

  // Header
  headerRow: { flexDirection: "row", justifyContent: "space-between", alignItems: "flex-start", borderBottomWidth: 2, borderBottomColor: INK, paddingBottom: 16, marginBottom: 26 },
  brandRow: { flexDirection: "row", alignItems: "center", marginBottom: 5 },
  logoMark: { width: 16, height: 16, borderRadius: 4, backgroundColor: BRAND, alignItems: "center", justifyContent: "center", marginRight: 6 },
  logoLetter: { color: "white", fontSize: 9, fontWeight: 800 },
  wordmark: { fontSize: 12, fontWeight: 800, color: INK, marginRight: 6 },
  eyebrow: { fontSize: 8.5, color: INK_SOFT, fontWeight: 600 },
  titleRow: { flexDirection: "row", alignItems: "center" },
  reportTitle: { fontSize: 20, fontWeight: 700, color: INK, marginRight: 10 },
  internalTag: { fontSize: 7.5, fontWeight: 700, color: "#a06c00", borderWidth: 1, borderColor: "#a06c00", borderRadius: 2, paddingVertical: 1, paddingHorizontal: 5 },
  metaBlock: { alignItems: "flex-end" },
  metaLine: { fontSize: 8, color: "#33404f", marginBottom: 2 },
  metaLabel: { fontWeight: 700, color: INK },

  // Executive summary
  execBox: { borderLeftWidth: 3, borderLeftColor: INK, paddingLeft: 16, marginBottom: 32 },
  execLabel: { fontSize: 8.5, fontWeight: 800, color: INK, textTransform: "uppercase", marginBottom: 5, letterSpacing: 0.5 },
  execText: { fontSize: 11, lineHeight: 1.65, color: "#202a38" },

  // KPI grid
  kpiRow: { flexDirection: "row", borderTopWidth: 1, borderTopColor: BORDER, borderBottomWidth: 1, borderBottomColor: BORDER, marginBottom: 28 },
  kpiBox: { flex: 1, paddingVertical: 16, paddingHorizontal: 14, borderRightWidth: 1, borderRightColor: BORDER, borderTopWidth: 3 },
  kpiBoxLast: { flex: 1, paddingVertical: 16, paddingHorizontal: 14, borderTopWidth: 3 },
  kpiLabelRow: { flexDirection: "row", alignItems: "center", marginBottom: 4 },
  kpiIcon: { fontSize: 8, marginRight: 3 },
  kpiLabel: { fontSize: 7.5, fontWeight: 800, color: "#45505f", textTransform: "uppercase" },
  // fontSize stepped down one notch (22->20) as a general margin-of-safety
  // adjustment, independent of the digit-truncation investigation — see
  // exportBriefingPDF's debug logging for that. Not presented as a fix.
  kpiValue: { fontSize: 20, fontWeight: 800, color: INK, marginBottom: 5 },
  kpiDeltaRow: { flexDirection: "row", alignItems: "center" },
  kpiDeltaText: { fontSize: 8.5, fontWeight: 700 },
  pill: { marginTop: 5, alignSelf: "flex-start", paddingVertical: 2, paddingHorizontal: 6, borderRadius: 7 },
  pillText: { fontSize: 6, fontWeight: 800, textTransform: "uppercase" },
  kpiNote: { fontSize: 7, color: "#5a6472", marginTop: 4 },

  // Badges (View-based, not ::before/::after — see spec's implementation note)
  badge: { width: 9, height: 9, borderRadius: 4.5, alignItems: "center", justifyContent: "center", marginRight: 3 },
  badgeText: { color: "white", fontSize: 6, fontWeight: 800 },

  // Section titles
  sectionTitle: { fontSize: 9.5, fontWeight: 800, color: INK, textTransform: "uppercase", borderBottomWidth: 1.5, borderBottomColor: INK, paddingBottom: 6, marginBottom: 12, letterSpacing: 0.4 },

  // Chart + watch items row
  bodyRow: { flexDirection: "row", gap: 28, marginBottom: 28 },
  col: { flex: 1 },

  chartLabel: { fontSize: 7.5, color: INK_SOFT, marginBottom: 8 },
  barRow: { flexDirection: "row", alignItems: "center", marginBottom: 14 },
  barMonth: { width: 28, fontSize: 7.5, fontWeight: 700, color: "#45505f" },
  barTrack: { flex: 1, height: 20, backgroundColor: "#eef0f3", borderRadius: 2 },
  // Fixed-width + right-aligned + bold is the one narrow, fragile combo in
  // this file — widened past what any realistic percentage value needs
  // (was 32, tight enough that a font-metric mismatch between renderers
  // could clip the leading digit) as a defensive margin.
  // fontSize also stepped down one notch (8.5->7.5), same independent
  // margin-of-safety adjustment as kpiValue above.
  barValue: { width: 46, textAlign: "right", fontSize: 7.5, fontWeight: 800, color: INK },
  chartNote: { fontSize: 6.5, color: "#8a94a3", marginTop: 2 },

  gaugeRow: { flexDirection: "row", alignItems: "center", marginTop: 4 },
  gaugeLabel: { width: 82, fontSize: 7, color: "#45505f", fontWeight: 600 },
  gaugeTrack: { flex: 1, height: 5, backgroundColor: "#eef0f3", borderRadius: 3 },
  gaugeNum: { width: 46, textAlign: "right", fontSize: 7, fontWeight: 800, color: INK },

  watchItem: { flexDirection: "row", alignItems: "flex-start", marginBottom: 10, padding: 11 },
  watchText: { fontSize: 8.5, lineHeight: 1.5, color: "#202a38", flex: 1 },

  // Drilldown table
  table: { borderWidth: 1, borderColor: BORDER, marginBottom: 16 },
  tableHeaderRow: { flexDirection: "row", backgroundColor: "#f2f5f9", borderBottomWidth: 1, borderBottomColor: BORDER },
  tableRow: { flexDirection: "row", borderBottomWidth: 1, borderBottomColor: "#eef0f3" },
  tableCellHead: { flex: 1, fontSize: 6.5, fontWeight: 800, color: "#45505f", textTransform: "uppercase", padding: 6 },
  tableCell: { flex: 1, fontSize: 7.5, color: INK, padding: 6 },
  tableStatusCell: { flex: 1, fontSize: 7.5, fontWeight: 700, padding: 6 },

  // Recommendations
  decisionRow: { flexDirection: "row", borderTopWidth: 1, borderTopColor: BORDER, borderBottomWidth: 1, borderBottomColor: BORDER },
  decisionCard: { flex: 1, borderRightWidth: 1, borderRightColor: BORDER, padding: 18 },
  decisionCardLast: { flex: 1, padding: 18 },
  decisionHeadRow: { flexDirection: "row", alignItems: "flex-start" },
  decisionNum: { width: 13, height: 13, borderRadius: 6.5, backgroundColor: INK, alignItems: "center", justifyContent: "center", marginRight: 6 },
  decisionNumText: { color: "white", fontSize: 7, fontWeight: 800 },
  decisionHead: { fontSize: 9.5, fontWeight: 800, color: INK, flex: 1 },
  decisionBody: { fontSize: 8.5, color: "#33404f", lineHeight: 1.5 },

  // Footer
  footer: { position: "absolute", bottom: 30, left: 40, right: 40, borderTopWidth: 1, borderTopColor: BORDER, paddingTop: 7 },
  footerText: { fontSize: 6.5, color: "#8a94a3", lineHeight: 1.4 },
});

// ---------------------------------------------------------------------------
// Small building blocks
// ---------------------------------------------------------------------------

/** Renders **bold** spans inline. Lets the caller mark which clause of the
 *  narrative is the real conclusion, without this component guessing. */
function RichText({ text, style }: { text: string; style?: Style }) {
  const parts = text.split(/(\*\*.*?\*\*)/g).filter((p) => p.length > 0);
  return (
    <Text style={style}>
      {parts.map((part, i) =>
        part.startsWith("**") && part.endsWith("**") ? (
          <Text key={i} style={{ fontWeight: 700, color: INK }}>{part.slice(2, -2)}</Text>
        ) : (
          <Text key={i}>{part}</Text>
        )
      )}
    </Text>
  );
}

/** Direction arrow drawn from borders, not a Unicode glyph — confirmed by
 *  test-render that even basic arrow characters (U+2193 etc.) don't
 *  reliably resolve in standard PDF fonts. Geometry is font-independent. */
function ArrowTriangle({ direction, color }: { direction: "up" | "down"; color: string }) {
  const base = { width: 0, height: 0, marginRight: 4, borderLeftWidth: 4, borderRightWidth: 4, borderLeftColor: "transparent" as const, borderRightColor: "transparent" as const };
  return (
    <View style={
      direction === "down"
        ? { ...base, borderTopWidth: 6, borderTopColor: color }
        : { ...base, borderBottomWidth: 6, borderBottomColor: color }
    } />
  );
}

function Badge({ status, symbol }: { status: Status; symbol: "!" | "check" }) {
  const color = STATUS[status].text;
  return (
    <View style={[styles.badge, { backgroundColor: color }]}>
      {symbol === "check" ? (
        <Svg width={6} height={6} viewBox="0 0 24 24">
          <Path d="M3 13 L9 19 L21 5" stroke="white" strokeWidth={4} fill="none" strokeLinecap="round" strokeLinejoin="round" />
        </Svg>
      ) : (
        <Text style={styles.badgeText}>!</Text>
      )}
    </View>
  );
}

function KPICard({ fact, isLast }: { fact: BriefingFact; isLast?: boolean }) {
  const status = fact.status ?? "neutral";
  const c = STATUS[status];
  const showBadge = status === "critical" || status === "watch" || status === "ok";
  return (
    <View style={[isLast ? styles.kpiBoxLast : styles.kpiBox, { borderTopColor: status === "neutral" ? "transparent" : c.text }]}>
      {/* Note: fact.icon is intentionally NOT rendered here. Emoji glyphs
          garble in PDF text rendering (confirmed by test-render) even
          with Inter registered — Inter has no color-emoji glyphs. A
          proper fix would be real SVG icons, not attempted here. */}
      <View style={styles.kpiLabelRow}>
        <Text style={styles.kpiLabel}>{fact.metric}</Text>
      </View>
      <Text style={styles.kpiValue}>{fact.value}{fact.unit ?? ""}</Text>
      {fact.delta ? (
        <View style={styles.kpiDeltaRow}>
          {showBadge ? <Badge status={status} symbol={status === "ok" ? "check" : "!"} /> : <ArrowTriangle direction={fact.delta.direction} color={c.text} />}
          <Text style={[styles.kpiDeltaText, { color: c.text }]}>{fact.delta.label}</Text>
        </View>
      ) : null}
      {fact.statusLabel ? (
        <View style={[styles.pill, { backgroundColor: c.pillBg }]}>
          <Text style={[styles.pillText, { color: c.text }]}>{fact.statusLabel}</Text>
        </View>
      ) : null}
      {/* Plain descriptive caption, not a status — no pill/color. For
          content like a share percentage that has no real target to be
          on/off track against (see BriefingFact.note). */}
      {fact.note ? <Text style={styles.kpiNote}>{fact.note}</Text> : null}
    </View>
  );
}

function Chart({ chart }: { chart: ComparisonChart }) {
  const max = Math.max(...chart.points.map((p) => p.value), chart.gauge?.value ?? 0, chart.gauge?.target ?? 0);
  return (
    <View>
      <Text style={styles.sectionTitle}>{chart.title}</Text>
      {chart.subtitle ? <Text style={styles.chartLabel}>{chart.subtitle}</Text> : null}
      {chart.points.map((p, i) => (
        <View key={i} style={styles.barRow}>
          <Text style={styles.barMonth}>{p.label}</Text>
          <View style={styles.barTrack}>
            <View style={{
              height: "100%", borderRadius: 2, width: `${(p.value / max) * 100}%`,
              backgroundColor: i === chart.points.length - 1 ? "#b03a2e" : "#9aa5b1",
            }} />
          </View>
          <Text style={styles.barValue}>{p.display}</Text>
        </View>
      ))}
      {chart.footnote ? <Text style={styles.chartNote}>{chart.footnote}</Text> : null}

      {chart.gauge ? (
        <View style={{ marginTop: 10 }}>
          <View style={styles.gaugeRow}>
            <Text style={styles.gaugeLabel}>{chart.gauge.label}</Text>
            <View style={styles.gaugeTrack}>
              <View style={{ height: "100%", borderRadius: 3, width: `${chart.gauge.value}%`, backgroundColor: "#b8860b" }} />
            </View>
            <Text style={styles.gaugeNum}>{chart.gauge.display}</Text>
          </View>
          {chart.gauge.footnote ? <Text style={styles.chartNote}>{chart.gauge.footnote}</Text> : null}
        </View>
      ) : null}
    </View>
  );
}

function WatchItems({ risks }: { risks: BriefingRisk[] }) {
  return (
    <View>
      <Text style={styles.sectionTitle}>Watch Items</Text>
      {risks.map((r, i) => {
        const c = STATUS[r.status];
        const showBadge = r.status !== "neutral";
        return (
          <View key={i} style={[styles.watchItem, { backgroundColor: c.rowBg, borderLeftWidth: 3, borderLeftColor: c.text }]}>
            {showBadge ? <Badge status={r.status} symbol={r.status === "ok" ? "check" : "!"} /> : null}
            <Text style={styles.watchText}>
              <Text style={{ fontWeight: 700, color: INK }}>{r.label} </Text>
              {r.detail}
            </Text>
          </View>
        );
      })}
    </View>
  );
}

function Drilldown({ table }: { table: DrilldownTable }) {
  return (
    <View style={{ marginBottom: 16 }}>
      <Text style={styles.sectionTitle}>{table.title}</Text>
      <View style={styles.table}>
        <View style={styles.tableHeaderRow}>
          {table.columns.map((c, i) => (
            <Text key={i} style={styles.tableCellHead}>{c}</Text>
          ))}
        </View>
        {table.rows.map((row, i) => {
          const c = STATUS[row.status];
          return (
            <View key={i} style={[styles.tableRow, { backgroundColor: row.status === "neutral" ? "white" : c.rowBg }]}>
              {row.cells.map((cell, j) => (
                <Text
                  key={j}
                  style={j === row.cells.length - 1 ? [styles.tableStatusCell, { color: c.text }] : styles.tableCell}
                >
                  {cell}
                </Text>
              ))}
            </View>
          );
        })}
      </View>
    </View>
  );
}

// ---------------------------------------------------------------------------
// Main component
// ---------------------------------------------------------------------------

export function ExecutiveBriefingPDF({ data }: { data: BriefingData }) {
  return (
    <Document>
      <Page size="A4" style={styles.page}>
        {/* Header */}
        <View style={styles.headerRow}>
          <View>
            <View style={styles.brandRow}>
              <View style={styles.logoMark}><Text style={styles.logoLetter}>O</Text></View>
              <Text style={styles.wordmark}>OneBoard</Text>
              <Text style={styles.eyebrow}>Executive Briefing</Text>
            </View>
            <View style={styles.titleRow}>
              <Text style={styles.reportTitle}>{data.reportTitle}</Text>
              <Text style={styles.internalTag}>INTERNAL</Text>
            </View>
          </View>
          <View style={styles.metaBlock}>
            <Text style={styles.metaLine}><Text style={styles.metaLabel}>Period: </Text>{data.period}</Text>
            <Text style={styles.metaLine}><Text style={styles.metaLabel}>Audience: </Text>{data.audience ?? "Executive Leadership"}</Text>
            <Text style={styles.metaLine}><Text style={styles.metaLabel}>Generated: </Text>{data.generatedAt}</Text>
          </View>
        </View>

        {/* Executive Summary */}
        <View style={styles.execBox}>
          <Text style={styles.execLabel}>Executive Summary</Text>
          <RichText text={data.narrative} style={styles.execText} />
          {data.keyObservations?.length ? (
            <View style={{ marginTop: 8 }}>
              {data.keyObservations.map((o, i) => (
                <View key={i} style={{ flexDirection: "row", marginBottom: 3 }}>
                  <Text style={{ fontSize: 9, marginRight: 5, color: "#5a6472" }}>{"•"}</Text>
                  <Text style={{ fontSize: 9.5, lineHeight: 1.5, color: "#33404f", flex: 1 }}>{o}</Text>
                </View>
              ))}
            </View>
          ) : null}
        </View>

        {/* KPI strip */}
        <View style={styles.kpiRow}>
          {data.kpis.map((k, i) => <KPICard key={k.id} fact={k} isLast={i === data.kpis.length - 1} />)}
        </View>

        {/* Chart + Watch Items */}
        {(data.chart || data.risks?.length) ? (
          <View style={styles.bodyRow}>
            {data.chart ? <View style={styles.col}><Chart chart={data.chart} /></View> : null}
            {data.risks?.length ? <View style={styles.col}><WatchItems risks={data.risks} /></View> : null}
          </View>
        ) : null}

        {/* Drill-down table — only rendered if the caller provided one;
            see spec: only build this from a real dimension in the dataset */}
        {data.drilldown ? <Drilldown table={data.drilldown} /> : null}

        {/* Recommendations */}
        {data.recommendations?.length ? (
          <View>
            <Text style={styles.sectionTitle}>Recommended Next Steps</Text>
            <View style={styles.decisionRow}>
              {data.recommendations.map((rec, i) => (
                <View key={i} style={i === data.recommendations!.length - 1 ? styles.decisionCardLast : styles.decisionCard}>
                  <View style={styles.decisionHeadRow}>
                    <View style={styles.decisionNum}><Text style={styles.decisionNumText}>{i + 1}</Text></View>
                    <Text style={styles.decisionBody}>{rec}</Text>
                  </View>
                </View>
              ))}
            </View>
          </View>
        ) : null}

        {/* Footer — `fixed` (not just `position: absolute`) so it's excluded
            from the page-break flow calculation entirely and rendered as a
            page-bottom overlay. Without `fixed`, react-pdf still treats it as
            needing room in the normal flow; when the content above nearly
            fills page 1, the whole block gets pushed onto an otherwise-empty
            page 2 instead of just sitting at the bottom of page 1. */}
        <View style={styles.footer} fixed>
          <Text style={styles.footerText}>
            Every figure above is computed directly from {data.dataSource}. No causal
            explanations or unobserved figures are included — only what the dashboard
            itself measured this period.
          </Text>
        </View>
      </Page>
    </Document>
  );
}
