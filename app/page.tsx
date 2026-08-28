"use client";

import { useEffect, useRef, useState } from "react";
import type { ReactNode, ComponentType } from "react";
import { detectAnomalies } from "./api/ai-narrative/_lib/anomalies";
import type { AiNarrativeFact, AiAnomaly, AiNarrativeResult } from "../types/ai-narrative";
import type { BriefingData, BriefingFact, BriefingRisk, ComparisonChart, DrilldownTable, Status } from "../components/ExecutiveBriefingPDF";
import type { ChartSpec, ColumnProfile } from "../types/kpi-spec";
import { profileColumns, sampleRows as sampleRowsForProfile } from "./api/kpi-suggest/_lib/columnProfiler";
import { computeChartSpec } from "./api/kpi-suggest/_lib/aggregationEngine";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  AreaChart,
  Area,
  LineChart,
  Line,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  Radar,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  Legend,
  LabelList,
} from "recharts";
import { ComposableMap, Geographies, Geography, Marker } from "react-simple-maps";
import Papa from "papaparse";
import {
  Users,
  Activity,
  Gauge,
  Compass,
  TrendingUp,
  Globe2,
  MapPin,
  PieChart as PieChartIcon,
  FileText,
  Flag,
  Clock,
  FileWarning,
  Radio,
  FileSpreadsheet,
  ArrowUpRight,
  ArrowDownRight,
  Minus,
  Smartphone,
  Layers,
  Cpu,
  Zap,
  Sparkles,
  Copy,
  CheckCircle2,
  Filter,
  Code,
  X,
  Lock,
  Upload,
  Sun,
  Moon,
  Loader2,
} from "lucide-react";

/* ============== Data shapes ============== */

type ExecutiveData = {
  summary: {
    activeUsers: number;
    sessions: number;
    engagementRate: number;
    topChannel: string;
  };
  trend: { date: string; activeUsers: number; sessions: number }[];
  channels: { channel: string; sessions: number }[];
  geoCountries?: { country: string; activeUsers: number }[];
  geoCities?: { city: string; activeUsers: number }[];
  newVsReturning?: { segment: string; activeUsers: number }[];
  deltas?: { activeUsers: number | null; sessions: number | null };
  queryMs?: number;
};

type MarketingData = {
  sources: { source: string; sessions: number }[];
  pages: { title: string; views: number }[];
  landingPages: { page: string; sessions: number }[];
  trend?: { date: string; sessions: number }[];
  channels?: { channel: string; sessions: number }[];
  funnel?: { sessions: number; engagedSessions: number; keyEvents: number | null };
  deltas?: { sessions: number | null };
  queryMs?: number;
};

type OperationsData = {
  summary: { bounceRate: number; avgSessionDuration: number };
  notFoundPages: { title: string; views: number }[];
  pages?: { title: string; views: number }[];
  devices?: { device: string; sessions: number }[];
  browsers?: { browser: string; sessions: number }[];
  operatingSystems?: { os: string; sessions: number }[];
  quota?: { consumed: number; remaining: number } | null;
  queryMs?: number;
};

const TABS = ["Executive", "Marketing", "Operations", "Custom Data"] as const;
type Tab = (typeof TABS)[number];
const TAB_LABELS: Record<Tab, string> = {
  Executive: "Executive View",
  Marketing: "Marketing & Funnel",
  Operations: "Operations & Diagnostics",
  "Custom Data": "Custom Data"
};
  
const RANGES = [
  { value: "7d", label: "7 Days" },
  { value: "30d", label: "30 Days" },
  { value: "90d", label: "90 Days" },
] as const;
type Range = (typeof RANGES)[number]["value"];

const DEFAULT_RANGE: Range = "30d";

// Export Report is temporarily hidden from the UI (not removed) — the
// underlying react-pdf/textkit character-dropping bug (see the extensive
// comment on KPICard in components/ExecutiveBriefingPDF.tsx) only affects
// the PDF export path, never the on-screen dashboard, so the safest
// mitigation is to stop users from triggering that path at all until
// there's real upstream progress. ExecutiveBriefingPDF.tsx, every
// *ToBriefing adapter (executiveToBriefing/marketingToBriefing/
// operationsToBriefing/customDataToBriefing), and handleExportBriefing
// are all left fully intact — flip this back to true to re-expose the
// button once it's safe to.
const EXPORT_REPORT_ENABLED = false;

/* ============== Design tokens ============== */

const COLORS = {
  bg: "var(--color-bg)",
  surface: "var(--color-surface)",
  ink: "var(--color-ink)",
  inkSoft: "var(--color-ink-soft)",
  inkFaint: "var(--color-ink-faint)",
  line: "var(--color-line)",
  track: "var(--color-track)",
  accent: "var(--color-accent)",
  accentSoft: "var(--color-accent-soft)",
  up: "var(--color-up)",
  upSoft: "var(--color-up-soft)",
  down: "var(--color-down)",
  downSoft: "var(--color-down-soft)",
  amberSoft: "var(--color-amber-soft)",
  amberInk: "var(--color-amber-ink)",
  blue: "var(--color-blue)",
  blueSoft: "var(--color-blue-soft)",
  teal: "var(--color-teal)",
  indigo: "var(--color-indigo)",
  indigoSoft: "var(--color-indigo-soft)",
};

const DONUT_COLORS = [
  "var(--donut-1)", "var(--donut-2)", "var(--donut-3)",
  "var(--donut-4)", "var(--donut-5)", "var(--donut-6)",
];

const SHADOW = "var(--color-shadow)";
const SHADOW_HOVER = "var(--color-shadow-hover)";

const FONT_STACK =
  '"Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif';
const MONO_STACK = '"JetBrains Mono", ui-monospace, "SF Mono", Menlo, monospace';

type IconType = ComponentType<{ size?: number; strokeWidth?: number; color?: string; className?: string }>;

/* ============== Helpers ============== */

function rangeLabel(range: Range) {
  if (range === "30d") return "last 30 days";
  if (range === "90d") return "last 90 days";
  return "last 7 days";
}

function tickInterval(length: number) {
  return Math.max(0, Math.ceil(length / 8) - 1);
}

function formatDate(raw: string) {
  const month = raw.slice(4, 6);
  const day = raw.slice(6, 8);
  return `${month}/${day}`;
}

function formatDuration(seconds: number) {
  const m = Math.floor(seconds / 60);
  const s = Math.round(seconds % 60);
  return `${m}m ${s}s`;
}

// The one formatting function every whole-number display in the app goes
// through — KPI cards and bar chart value labels alike — so the same kind
// of number never shows thousands separators in one place and not the
// other for no reason.
function formatInteger(n: number): string {
  return Math.round(n).toLocaleString();
}

function useGa4<T>(endpoint: string, range: Range) {
  const [data, setData] = useState<T | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isFetching, setIsFetching] = useState(false);
  const [clientMs, setClientMs] = useState<number | undefined>(undefined);

  useEffect(() => {
    setIsFetching(true);
    const start = performance.now();
    fetch(`${endpoint}?range=${range}`)
      .then((res) => res.json())
      .then((json) => {
        setClientMs(Math.round(performance.now() - start));
        if (json.error) setError(json.error);
        else {
          setError(null);
          setData(json);
        }
      })
      .catch(() => {
        setClientMs(Math.round(performance.now() - start));
        setError("Could not reach the GA4 endpoint.");
      })
      .finally(() => setIsFetching(false));
  }, [endpoint, range]);

  return { data, error, isFetching, clientMs };
}

function useCountUp(value: number, duration = 700) {
  const [display, setDisplay] = useState(value);
  const prevRef = useRef(value);

  useEffect(() => {
    const from = prevRef.current;
    const to = value;
    if (from === to) return;

    const prefersReducedMotion =
      typeof window !== "undefined" &&
      window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;

    if (prefersReducedMotion) {
      setDisplay(to);
      prevRef.current = to;
      return;
    }

    const start = performance.now();
    let raf: number;
    const tick = (now: number) => {
      const t = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - t, 3);
      setDisplay(from + (to - from) * eased);
      if (t < 1) raf = requestAnimationFrame(tick);
      else prevRef.current = to;
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [value, duration]);

  return display;
}

/* ============== Small building blocks ============== */

function ErrorBox({ message }: { message: string }) {
  return (
    <div className="rounded-2xl p-4 text-sm" style={{ background: COLORS.downSoft, color: "#991B1B" }}>
      {message}
    </div>
  );
}

function RefreshingNote({ show }: { show: boolean }) {
  if (!show) return null;
  return (
    <p className="text-xs mb-3 flex items-center gap-1.5" style={{ color: COLORS.inkFaint }}>
      <span className="w-1 h-1 rounded-full animate-pulse" style={{ background: COLORS.accent }} />
      Refreshing…
    </p>
  );
}

function SegmentedControl<T extends string>({
  options,
  value,
  onChange,
  size = "md",
  variant = "neutral",
}: {
  options: readonly { value: T; label: ReactNode }[];
  value: T;
  onChange: (v: T) => void;
  size?: "sm" | "md";
  variant?: "neutral" | "accent";
}) {
  const pad = size === "sm" ? "px-3 py-1.5 text-xs" : "px-4 py-2 text-xs";
  return (
    <div className="inline-flex p-1 rounded-2xl gap-0.5" style={{ background: COLORS.track }}>
      {options.map((o) => {
        const active = o.value === value;
        const activeBg = variant === "accent" ? COLORS.accent : COLORS.surface;
        const activeColor = variant === "accent" ? "#fff" : COLORS.ink;
        const activeShadow =
          variant === "accent"
            ? "0 2px 10px -2px rgba(5,150,105,0.4)"
            : "0 1px 2px rgba(15,23,42,0.06), 0 1px 6px rgba(15,23,42,0.05)";
        const activeBorder = variant === "accent" ? "1px solid transparent" : `1px solid ${COLORS.line}`;
        return (
          <button
            key={o.value}
            onClick={() => onChange(o.value)}
            className={`rounded-xl font-semibold transition-all ${pad}`}
            style={{
              color: active ? activeColor : COLORS.inkSoft,
              background: active ? activeBg : "transparent",
              boxShadow: active ? activeShadow : "none",
              border: active ? activeBorder : "1px solid transparent",
            }}
          >
            {o.label}
          </button>
        );
      })}
    </div>
  );
}

function Card({
  title,
  icon: Icon,
  badge,
  children,
  className = "",
}: {
  title: string;
  icon?: IconType;
  badge?: ReactNode;
  children: ReactNode;
  className?: string;
}) {
  return (
    <div
      className={`rounded-2xl p-5 transition-shadow duration-200 ${className}`}
      style={{
        background: COLORS.surface,
        border: `1px solid ${COLORS.line}`,
        boxShadow: SHADOW,
        backdropFilter: "blur(14px)",
        WebkitBackdropFilter: "blur(14px)",
      }}
    >
      <div className="flex items-center justify-between mb-4 gap-2">
        <div className="flex items-center gap-2 min-w-0">
          {Icon && <Icon size={14} strokeWidth={2} color={COLORS.inkSoft} />}
          <p
            className="text-[11px] font-bold uppercase tracking-wider truncate"
            style={{ color: COLORS.inkSoft }}
          >
            {title}
          </p>
        </div>
        {badge}
      </div>
      {children}
    </div>
  );
}

function Sparkline({ data, color }: { data: { label: string; value: number }[]; color: string }) {
  if (!data || data.length < 2) return <div style={{ width: 88, height: 32 }} />;
  return (
    <div style={{ width: 88, height: 32 }}>
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={data} margin={{ top: 2, right: 2, bottom: 2, left: 2 }}>
          <Line
            type="monotone"
            dataKey="value"
            stroke={color}
            strokeWidth={1.75}
            dot={false}
            isAnimationActive={false}
          />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}

function DeltaBadge({ value, invert = false }: { value: number | null | undefined; invert?: boolean }) {
  if (value === null || value === undefined || !Number.isFinite(value)) {
    return (
      <span className="text-xs" style={{ color: COLORS.inkFaint }}>
        —
      </span>
    );
  }
  const rounded = Math.round(value * 10) / 10;
  const isFlat = Math.abs(rounded) < 0.1;
  const isUp = rounded > 0;
  const good = isFlat ? null : invert ? !isUp : isUp;
  const color = isFlat ? COLORS.inkFaint : good ? COLORS.up : COLORS.down;
  const bg = isFlat ? COLORS.track : good ? COLORS.upSoft : COLORS.downSoft;
  const Arrow = isFlat ? Minus : isUp ? ArrowUpRight : ArrowDownRight;
  return (
    <span
      className="inline-flex items-center gap-0.5 rounded-full px-1.5 py-0.5 text-xs font-semibold"
      style={{ color, background: bg, fontVariantNumeric: "tabular-nums" }}
    >
      <Arrow size={12} strokeWidth={2} />
      {Math.abs(rounded).toFixed(1)}%
    </span>
  );
}

function KpiCard({
  icon: Icon,
  iconColor,
  iconBg,
  label,
  value,
  suffix = "",
  decimals = 0,
  sparklineData,
  sparklineColor,
  delta,
  invertDelta,
  caption,
  displayValue,
}: {
  icon: IconType;
  iconColor: string;
  iconBg: string;
  label: string;
  value: number;
  suffix?: string;
  decimals?: number;
  sparklineData?: { label: string; value: number }[];
  sparklineColor?: string;
  delta?: number | null;
  invertDelta?: boolean;
  caption?: string;
  displayValue?: string;
}) {
  const animated = useCountUp(displayValue !== undefined ? 0 : value);
  const shown =
    displayValue !== undefined
      ? displayValue
      : decimals > 0
        ? animated.toFixed(decimals)
        : formatInteger(animated);

  return (
    <div
      className="rounded-2xl p-5 transition-shadow duration-200 hover:shadow-lg"
      style={{
        background: COLORS.surface,
        border: `1px solid ${COLORS.line}`,
        boxShadow: SHADOW,
        backdropFilter: "blur(14px)",
        WebkitBackdropFilter: "blur(14px)",
      }}
    >
      <div className="flex items-center justify-between mb-3">
        {/* Was a single-line `truncate` — a long title (e.g. an
            AI-suggested Custom Data KPI name) got cut to an unreadable
            fragment with "…" after just a few words. line-clamp-2 lets it
            wrap onto a second line first, only falling back to ellipsis if
            it's still too long after that — same "wrap before you
            truncate" approach already used for long string values in
            ExecutiveBriefingPDF.tsx's kpiValueClamped. */}
        <p className="text-[11px] font-bold uppercase tracking-wider line-clamp-2" style={{ color: COLORS.inkSoft }}>
          {label}
        </p>
        <div
          className="w-6 h-6 rounded-full flex items-center justify-center shrink-0"
          style={{ background: iconBg }}
        >
          <Icon size={13} strokeWidth={2.25} color={iconColor} />
        </div>
      </div>

      <div className="flex items-end justify-between gap-2">
        <p
          className="font-extrabold leading-none truncate"
          style={{
            color: COLORS.ink,
            fontVariantNumeric: "tabular-nums",
            letterSpacing: "-0.02em",
            fontSize: 15,
          }}
        >
          {shown}
          {suffix}
        </p>
        {sparklineData && sparklineData.length > 1 && (
          <Sparkline data={sparklineData} color={sparklineColor ?? iconColor} />
        )}
      </div>

      <div className="mt-2 h-5 flex items-center">
        {delta !== undefined && delta !== null ? (
          <DeltaBadge value={delta} invert={invertDelta} />
        ) : caption ? (
          <span className="text-xs truncate font-medium" style={{ color: COLORS.inkFaint }}>
            {caption}
          </span>
        ) : null}
      </div>
    </div>
  );
}

function ChartTooltip({ active, payload, label }: any) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div
      className="rounded-xl px-3 py-2 text-xs"
      style={{ background: "#0F172A", color: "#fff", boxShadow: SHADOW_HOVER }}
    >
      <p className="mb-1" style={{ color: "#94A3B8" }}>
        {label}
      </p>
      {payload.map((p: any, i: number) => (
        <p key={i} className="font-semibold" style={{ fontVariantNumeric: "tabular-nums", color: "#fff" }}>
          {typeof p.value === "number" ? p.value.toLocaleString() : p.value}
        </p>
      ))}
    </div>
  );
}

function BarListChart({
  title,
  icon,
  rows,
  barColor,
  emptyLabel = "No data in this range.",
  className = "",
  decimals,
}: {
  title: string;
  icon?: IconType;
  rows: { label: string; value: number }[];
  barColor?: string;
  emptyLabel?: string;
  className?: string;
  // When set to 2, the value label is formatted to exactly 2 decimal
  // places (toFixed) instead of as a whole number — e.g. Custom Data's
  // sum/avg-over-a-decimal-column results, which must always show 2
  // digits (0.50, not 0.5). Every other caller displays whole numbers, via
  // the same formatInteger() a KPI card uses — never left unformatted.
  decimals?: number;
}) {
  const chartData = rows.map((r) => ({
    label: r.label.length > 26 ? r.label.slice(0, 24) + "…" : r.label,
    value: r.value,
  }));
  const height = Math.max(120, chartData.length * 38);
  const labelFormatter = (v: string | number | boolean | null | undefined) =>
    typeof v === "number" ? (decimals !== undefined ? v.toFixed(decimals) : formatInteger(v)) : v;

  return (
    <Card title={title} icon={icon} className={className}>
      {rows.length === 0 ? (
        <p className="text-sm" style={{ color: COLORS.inkFaint }}>
          {emptyLabel}
        </p>
      ) : (
        <ResponsiveContainer width="100%" height={height}>
          <BarChart data={chartData} layout="vertical" margin={{ left: 8, right: 44, top: 4, bottom: 4 }}>
            {/* Longest bar is scaled to 90% of the plot width (not 100%) so
                its value label — drawn just past the bar's end — always has
                real room to its right instead of running into the plot's
                edge, on top of the fixed right margin above. */}
            <XAxis type="number" hide domain={[0, (dataMax: number) => (dataMax > 0 ? dataMax / 0.9 : 1)]} />
            <YAxis
              type="category"
              dataKey="label"
              width={130}
              fontSize={12}
              stroke={COLORS.inkSoft}
              tickLine={false}
              axisLine={false}
            />
            <Tooltip content={<ChartTooltip />} cursor={{ fill: COLORS.track }} />
            <Bar dataKey="value" fill={barColor ?? COLORS.accent} radius={[8, 8, 8, 8]} barSize={10}>
              <LabelList
                dataKey="value"
                position="right"
                formatter={labelFormatter}
                style={{ fill: COLORS.ink, fontSize: 12, fontWeight: 600 }}
              />
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      )}
    </Card>
  );
}

function ShareDonut({
  title,
  icon,
  rows,
  valueLabel = "sessions",
  colors,
}: {
  title: string;
  icon?: IconType;
  rows: { label: string; value: number }[];
  valueLabel?: string;
  colors?: string[];
}) {
  const total = rows.reduce((sum, r) => sum + r.value, 0);
  const palette = colors ?? DONUT_COLORS;

  return (
    <Card title={title} icon={icon}>
      {rows.length === 0 ? (
        <p className="text-sm" style={{ color: COLORS.inkFaint }}>
          No data in this range.
        </p>
      ) : (
        <div className="flex items-center gap-4">
          <ResponsiveContainer width="100%" height={200}>
            <PieChart>
              <Pie
                data={rows}
                dataKey="value"
                nameKey="label"
                innerRadius={54}
                outerRadius={82}
                paddingAngle={3}
                cornerRadius={4}
              >
                {rows.map((_, i) => (
                  <Cell key={i} fill={palette[i % palette.length]} stroke="none" />
                ))}
              </Pie>
              <Tooltip content={<ChartTooltip />} />
              <Legend
                layout="vertical"
                verticalAlign="middle"
                align="right"
                iconType="circle"
                iconSize={7}
                formatter={(value) => (
                  <span className="text-xs font-medium" style={{ color: COLORS.inkSoft }}>
                    {value}
                  </span>
                )}
              />
            </PieChart>
          </ResponsiveContainer>
        </div>
      )}
      {total > 0 && (
        <p
          className="text-xs mt-2 text-center font-medium"
          style={{ color: COLORS.inkFaint, fontVariantNumeric: "tabular-nums" }}
        >
          {total.toLocaleString()} {valueLabel} total
        </p>
      )}
    </Card>
  );
}

/* ============== Skeleton loaders ============== */

function SkeletonKpiRow({ count }: { count: number }) {
  return (
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
      {Array.from({ length: count }).map((_, i) => (
        <div
          key={i}
          className="rounded-2xl p-5"
          style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW }}
        >
          <div className="ob-skeleton rounded-full h-3 w-20 mb-4" />
          <div className="ob-skeleton rounded-lg h-7 w-16" />
        </div>
      ))}
    </div>
  );
}

function SkeletonChartGrid({ count }: { count: number }) {
  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
      {Array.from({ length: count }).map((_, i) => (
        <div
          key={i}
          className="rounded-2xl p-5"
          style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW }}
        >
          <div className="ob-skeleton rounded-full h-3 w-32 mb-4" />
          <div className="ob-skeleton rounded-xl" style={{ height: 180 }} />
        </div>
      ))}
    </div>
  );
}

function ExecutiveSkeleton() {
  return (
    <div>
      <SkeletonKpiRow count={4} />
      <SkeletonChartGrid count={4} />
    </div>
  );
}

function MarketingSkeleton() {
  return (
    <div>
      <SkeletonKpiRow count={4} />
      <SkeletonChartGrid count={4} />
    </div>
  );
}

function OperationsSkeleton() {
  return (
    <div>
      <SkeletonKpiRow count={3} />
      <SkeletonChartGrid count={3} />
    </div>
  );
}

/* ============== Raw payload modal (real data, not decoration) ============== */

function RawPayloadModal({ payload, onClose }: { payload: unknown; onClose: () => void }) {
  const [copied, setCopied] = useState(false);
  const json = JSON.stringify(payload, null, 2);

  function copy() {
    navigator.clipboard.writeText(json);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ background: "rgba(15,23,42,0.7)" }}
      onClick={onClose}
    >
      <div
        className="rounded-3xl max-w-2xl w-full p-6 space-y-4"
        style={{
          background: "#0F172A",
          color: "#E2E8F0",
          boxShadow: SHADOW_HOVER,
          border: "1px solid rgba(255,255,255,0.08)",
          fontFamily: MONO_STACK,
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          className="flex items-center justify-between pb-3 text-xs"
          style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}
        >
          <div className="flex items-center gap-2 font-bold" style={{ color: "#34D399" }}>
            <Code size={15} />
            <span style={{ fontFamily: FONT_STACK }}>Live Data Payload</span>
          </div>
          <button onClick={onClose} style={{ color: "#94A3B8" }} aria-label="Close">
            <X size={16} />
          </button>
        </div>
        <pre
          className="p-4 rounded-2xl overflow-x-auto leading-relaxed text-[11px]"
          style={{ background: "#020617", color: "#34D399", maxHeight: 340 }}
        >
          {json}
        </pre>
        <div className="flex justify-between items-center pt-1">
          <span className="text-[11px]" style={{ color: "#64748B", fontFamily: FONT_STACK }}>
            Exactly what each tab's API route returned, live.
          </span>
          <button
            onClick={copy}
            className="px-4 py-2 rounded-xl font-semibold text-xs flex items-center gap-2"
            style={{ background: "#059669", color: "#fff", fontFamily: FONT_STACK }}
          >
            {copied ? <CheckCircle2 size={13} /> : <Copy size={13} />}
            {copied ? "Copied" : "Copy JSON"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============== Source panel (real GA4 status, real Excel file picker) ============== */

type CustomDataState = {
  connected: boolean;
  filename: string | null;
  rows: number | null;
  source: "upload" | "demo" | null;
  // Full parsed dataset — the aggregation engine always computes from this,
  // never from the small sample sent to /api/kpi-suggest.
  data: Record<string, unknown>[] | null;
  profile: ColumnProfile[] | null;
  // KPI/chart suggestion runs ONCE, right after parsing, via
  // /api/kpi-suggest — never per-render. The actual numbers are computed
  // separately, in code, from the full dataset (see aggregationEngine.ts).
  specs: ChartSpec[] | null;
  suggesting: boolean;
  suggestError: string | null;
};

// Single source of truth for "has the Custom Data tab actually finished
// rendering something" — Custom Data never writes into `latestData` (it
// has its own local state, not a GA4 fetch), so anything that gates on
// "has this tab loaded" (e.g. Export Report) must check this instead of
// latestData, or it wrongly reports the tab as still loading even after
// its KPI cards are on screen.
function isCustomDataReady(state: CustomDataState): boolean {
  return state.connected && !!state.data && !!state.specs && state.specs.length > 0;
}

function SourcePanel({
  activeEndpoint,
  customData,
  onFileSelect,
  customDataActive,
}: {
  activeEndpoint: string;
  customData: CustomDataState;
  onFileSelect: (file: File) => void;
  customDataActive: boolean;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  // Border highlight only indicates which source the CURRENT tab is
  // reading from. It does not imply the other source is broken or
  // unavailable — GA4 stays fully "live" in its own copy for the other
  // three tabs regardless of which tab is active right now.
  const excelBorderActive = customData.suggesting || customDataActive;

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
      {/* GA4 — content never changes with tab; only the border highlight
          toggles off when this tab isn't reading from it. */}
      <div
        className="ob-ga4-card rounded-2xl p-4 hover:-translate-y-0.5"
        style={{
          background: `linear-gradient(to bottom, ${COLORS.surface}, ${COLORS.accentSoft})`,
          border: customDataActive ? `1px solid ${COLORS.line}` : `1px solid var(--color-accent-border)`,
          boxShadow: SHADOW,
          backdropFilter: "blur(14px)",
          WebkitBackdropFilter: "blur(14px)",
        }}
      >
        <div className="flex items-center justify-between mb-2">
          <span className="font-bold text-xs uppercase tracking-wider" style={{ color: COLORS.ink }}>
            GA4 DirectLink
          </span>
          <span
            className="px-2.5 py-0.5 text-[10px] font-bold rounded-full flex items-center gap-1.5"
            style={{ background: COLORS.upSoft, color: COLORS.up }}
          >
            <span className="w-1.5 h-1.5 rounded-full ob-live-dot" style={{ background: COLORS.up }} />
            Active Stream
          </span>
        </div>
        <p className="text-xs font-medium" style={{ color: COLORS.inkSoft }}>
          Live GA4 Data API DirectQuery
        </p>
        <div className="mt-3 text-[11px] font-semibold flex items-center gap-1.5" style={{ color: COLORS.up }}>
          <CheckCircle2 size={13} />
          <span style={{ fontFamily: MONO_STACK }}>{activeEndpoint}</span>
        </div>
      </div>

      {/* Excel / CSV — border highlights on the Custom Data tab (same
          data-source category) or while a just-uploaded file is being
          profiled. Text content reflects the REAL upload state only — it
          never claims a sample dataset is "connected" via upload. */}
      <div
        onClick={() => inputRef.current?.click()}
        className="ob-excel-card rounded-2xl p-4 cursor-pointer hover:-translate-y-0.5"
        style={{
          background: excelBorderActive
            ? `linear-gradient(to bottom, ${COLORS.surface}, ${COLORS.accentSoft})`
            : COLORS.surface,
          border: excelBorderActive ? `1.5px solid ${COLORS.accent}` : `1px solid ${COLORS.line}`,
          boxShadow: excelBorderActive ? `0 0 0 3px ${COLORS.accentSoft}, ${SHADOW}` : SHADOW,
          transition: "border-color 200ms ease, box-shadow 200ms ease",
        }}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) onFileSelect(file);
          }}
        />
        <div className="flex items-center justify-between mb-2">
          <span className="font-bold text-xs uppercase tracking-wider" style={{ color: COLORS.ink }}>
            Excel / CSV Import
          </span>
          <span
            className="px-2.5 py-0.5 text-[10px] font-bold rounded-full"
            style={{
              background: customData.connected ? COLORS.accentSoft : COLORS.track,
              color: customData.connected ? COLORS.accent : COLORS.inkSoft,
            }}
          >
            {customData.connected ? (customData.source === "demo" ? "Demo Loaded" : "File Selected") : "Import Mode"}
          </span>
        </div>
        <p className="text-xs font-medium truncate" style={{ color: COLORS.inkSoft }}>
          {customData.filename ?? "Click to auto-generate KPIs from any spreadsheet"}
        </p>
        <div
          className="ob-excel-cta mt-3 text-[11px] font-semibold flex items-center gap-1.5"
          style={{ color: excelBorderActive ? COLORS.accent : COLORS.inkFaint }}
        >
          <Upload size={13} />
          <span>
            {customData.connected
              ? customData.suggesting
                ? `${customData.rows ?? "…"} rows parsed — suggesting KPIs…`
                : customData.specs
                ? `${customData.rows ?? customData.data?.length ?? 0} rows parsed — ${customData.specs.length} KPI card${customData.specs.length === 1 ? "" : "s"} generated`
                : customData.suggestError
                ? "KPI suggestion failed — see error above"
                : customData.rows !== null
                ? `${customData.rows} rows parsed from ${customData.filename ?? "file"}`
                : "File selected — ready to parse"
              : "+ Upload Data File (.xlsx/.xls/.csv)"}
          </span>
        </div>
      </div>

      {/* ERP */}
      <div
        className="rounded-2xl p-4 opacity-80"
        style={{
          background: COLORS.surface,
          border: `1px solid ${COLORS.line}`,
          backdropFilter: "blur(14px)",
          WebkitBackdropFilter: "blur(14px)",
        }}
      >
        <div className="flex items-center justify-between mb-2">
          <span className="font-bold text-xs uppercase tracking-wider" style={{ color: COLORS.inkSoft }}>
            ERP / CRM Connector
          </span>
          <span
            className="px-2.5 py-0.5 text-[10px] font-medium rounded-full"
            style={{ background: COLORS.line, color: COLORS.inkSoft }}
          >
            Standby
          </span>
        </div>
        <p className="text-xs font-medium" style={{ color: COLORS.inkFaint }}>
          SAP S/4HANA &amp; Dynamics 365 Connector
        </p>
        <div className="mt-3 text-[11px] font-semibold flex items-center gap-1.5" style={{ color: COLORS.inkFaint }}>
          <Lock size={13} />
          Enterprise API Standby
        </div>
      </div>
    </div>
  );
}



/* ============== Country choropleth (real react-simple-maps + real GA4 data) ============== */

const GEO_URL = "/world-countries-110m.json";

const COUNTRY_ALIASES: Record<string, string> = {
  "united states": "united states of america",
  russia: "russian federation",
  "south korea": "republic of korea",
  "north korea": "dem. rep. korea",
  czechia: "czech republic",
  vietnam: "viet nam",
};

const COUNTRY_CENTROIDS: Record<string, [number, number]> = {
  canada: [-106.35, 56.13],
  "united states": [-95.71, 37.09],
  mexico: [-102.55, 23.63],
  brazil: [-51.93, -14.24],
  argentina: [-63.62, -38.42],
  "united kingdom": [-3.44, 55.38],
  ireland: [-8.24, 53.41],
  germany: [10.45, 51.17],
  france: [2.21, 46.23],
  spain: [-3.75, 40.46],
  italy: [12.57, 41.87],
  netherlands: [5.29, 52.13],
  belgium: [4.47, 50.5],
  switzerland: [8.23, 46.82],
  portugal: [-8.22, 39.4],
  sweden: [18.64, 60.13],
  norway: [8.47, 60.47],
  denmark: [9.5, 56.26],
  finland: [25.75, 61.92],
  poland: [19.15, 51.92],
  "czech republic": [15.47, 49.82],
  czechia: [15.47, 49.82],
  austria: [14.55, 47.52],
  russia: [105.32, 61.52],
  china: [104.2, 35.86],
  japan: [138.25, 36.2],
  "south korea": [127.77, 35.91],
  india: [78.96, 20.59],
  australia: [133.78, -25.27],
  "new zealand": [174.89, -40.9],
  "south africa": [22.94, -30.56],
  nigeria: [8.68, 9.08],
  egypt: [30.8, 26.82],
  indonesia: [113.92, -0.79],
  philippines: [121.77, 12.88],
  thailand: [100.99, 15.87],
  vietnam: [108.28, 14.06],
  singapore: [103.82, 1.35],
  turkey: [35.24, 38.96],
  "saudi arabia": [45.08, 23.89],
  "united arab emirates": [53.85, 23.42],
  israel: [34.85, 31.05],
  pakistan: [69.35, 30.38],
  bangladesh: [90.36, 23.68],
  colombia: [-74.3, 4.57],
  chile: [-71.54, -35.68],
  peru: [-75.02, -9.19],
};

function CountryChoropleth({ rows }: { rows: { country: string; activeUsers: number }[] }) {
  const [hover, setHover] = useState<{ name: string; value: number } | null>(null);
  const byName = new Map<string, number>();
  rows.forEach((r) => byName.set(r.country.trim().toLowerCase(), r.activeUsers));
  const max = rows.reduce((m, r) => Math.max(m, r.activeUsers), 0);
  const totalUsers = rows.reduce((s, r) => s + r.activeUsers, 0);

  function valueFor(geoName: string): number | null {
    const key = geoName.trim().toLowerCase();
    if (byName.has(key)) return byName.get(key)!;
    for (const [ga4Name, mapName] of Object.entries(COUNTRY_ALIASES)) {
      if (mapName === key && byName.has(ga4Name)) return byName.get(ga4Name)!;
    }
    return null;
  }

  function colorFor(value: number | null) {
    if (value === null || max === 0) return "var(--color-map-empty)";
    const intensity = 0.28 + 0.72 * (value / max);
    return `rgba(5, 150, 105, ${intensity})`;
  }

  if (rows.length === 0) {
    return <p className="text-sm" style={{ color: COLORS.inkFaint }}>No data in this range.</p>;
  }

  const MARKER_COLORS = [COLORS.accent, COLORS.blue, COLORS.indigo, COLORS.teal, "#B45309"];
  const labeled = [...rows]
    .sort((a, b) => b.activeUsers - a.activeUsers)
    .map((r) => ({ ...r, centroid: COUNTRY_CENTROIDS[r.country.trim().toLowerCase()] }))
    .filter((r): r is typeof r & { centroid: [number, number] } => !!r.centroid)
    .slice(0, 5);

  return (
    <div style={{ position: "relative" }}>
      <div
        style={{
          background: "var(--color-map-canvas)",
          borderRadius: 16,
          padding: 12,
        }}
      >
        <ComposableMap
          projection="geoEqualEarth"
          projectionConfig={{ scale: 240, center: [-95, 42] }}
          width={800}
          height={420}
          style={{ width: "100%", height: "auto" }}
        >
        <Geographies geography={GEO_URL}>
          {({ geographies }: { geographies: any[] }) =>
            geographies.map((geo) => {
              const name: string = geo.properties?.name ?? "";
              const value = valueFor(name);
              return (
                <Geography
                  key={geo.rsmKey}
                  geography={geo}
                  onMouseEnter={() => value !== null && setHover({ name, value })}
                  onMouseLeave={() => setHover(null)}
                  style={{
                    default: { fill: colorFor(value), stroke: COLORS.surface, strokeWidth: 0.5, outline: "none" },
                    hover: {
                      fill: value !== null ? "#047857" : colorFor(value),
                      stroke: COLORS.surface,
                      strokeWidth: 0.5,
                      outline: "none",
                    },
                    pressed: { fill: "#065F46", stroke: COLORS.surface, strokeWidth: 0.5, outline: "none" },
                  }}
                />
              );
            })
          }
        </Geographies>

        {labeled.map((r, i) => {
          const dotColor = MARKER_COLORS[i % MARKER_COLORS.length];
          return (
            <Marker key={`dot-${r.country}`} coordinates={r.centroid}>
              <circle
                r={9}
                fill="#fff"
                opacity={0.6}
                className="ob-geo-ping"
                style={{ filter: "drop-shadow(0 0 1px rgba(15,23,42,0.3))" }}
              />
              <circle r={9} fill="none" stroke={dotColor} strokeWidth={1.5} opacity={0.7} className="ob-geo-ping" />
              <circle r={4.5} fill={dotColor} stroke="#fff" strokeWidth={1.5} />
            </Marker>
          );
        })}

        {labeled.map((r, i) => {
          const dotColor = MARKER_COLORS[i % MARKER_COLORS.length];
          const label = `${r.country} (${r.activeUsers.toLocaleString()})`;
          const pillWidth = label.length * 8 + 26;
          return (
            <Marker key={`pill-${r.country}`} coordinates={r.centroid}>
              <g transform={`translate(10, -20)`}>
                <rect
                  x={0}
                  y={0}
                  width={pillWidth}
                  height={26}
                  rx={13}
                  ry={13}
                  className="ob-map-pill-bg"
                  stroke={dotColor}
                  strokeWidth={1.5}
                  style={{ filter: "drop-shadow(0 1px 2px rgba(15,23,42,0.15))" }}
                />
                <circle cx={14} cy={13} r={3.5} fill={dotColor} />
                <text
                  x={23}
                  y={18}
                  className="ob-map-pill-text"
                  style={{ fontSize: 14, fontWeight: 700, fontFamily: FONT_STACK, ["--pill-color" as any]: dotColor }}
                >
                  {label}
                </text>
              </g>
            </Marker>
          );
        })}
      </ComposableMap>
      </div>
      {hover && (
        <div
          className="text-xs font-semibold px-2.5 py-1 rounded-lg"
          style={{
            position: "absolute",
            top: 8,
            left: 8,
            background: "#0F172A",
            color: "#fff",
            fontVariantNumeric: "tabular-nums",
          }}
        >
          {hover.name} · {hover.value.toLocaleString()} users
        </div>
      )}
    </div>
  );
}

/* ============== Live Insights (real derived numbers, no LLM call) ============== */

type InsightsInput = {
  range: string;
  activeUsers: number;
  sessions: number;
  engagementRate: number;
  topChannel: string;
  topChannelSharePct: number | null;
  topCountry: string | null;
  topCountrySharePct: number | null;
};

function buildInsightBullets(input: InsightsInput): string[] {
  const bullets: string[] = [];

  if (input.topChannelSharePct !== null) {
    bullets.push(
      `${input.topChannel} is the top channel, driving ${input.topChannelSharePct.toFixed(0)}% of sessions.`
    );
  }

  if (input.topCountry && input.topCountrySharePct !== null) {
    bullets.push(
      `${input.topCountry} leads by geography, at ${input.topCountrySharePct.toFixed(0)}% of active users.`
    );
  }

  bullets.push(
    `${input.activeUsers.toLocaleString()} active users generated ${input.sessions.toLocaleString()} sessions at a ${(
      input.engagementRate * 100
    ).toFixed(0)}% engagement rate over the ${input.range}.`
  );

  return bullets.slice(0, 3);
}

function LiveInsightsCard({
  input,
  facts,
  anomalies,
  context,
  result,
  onGenerated,
}: {
  input: InsightsInput | null;
  facts: AiNarrativeFact[];
  anomalies?: AiAnomaly[];
  context?: Record<string, number | string>;
  result: AiNarrativeResult | undefined;
  onGenerated: (result: AiNarrativeResult) => void;
}) {

  const [copied, setCopied] = useState(false);
  const bullets = input ? buildInsightBullets(input) : [];

  function copy() {
    if (!input) return;
    const text = `ONEBOARD LIVE INSIGHTS (${input.range.toUpperCase()}):\n${bullets
      .map((b) => `• ${b}`)
      .join("\n")}`;
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  return (
    <div
      className="rounded-2xl p-5 flex flex-col justify-between"
      style={{
        background: "var(--color-highlight-bg)",
        color: "var(--color-highlight-text)",
        border: "1px solid var(--color-highlight-border)",
        boxShadow: SHADOW,
        backdropFilter: "blur(14px)",
        WebkitBackdropFilter: "blur(14px)",
      }}
    >
      <div>
        <div
          className="flex items-center justify-between pb-3 mb-3"
          style={{ borderBottom: "1px solid var(--color-highlight-divider)" }}
        >
          <div className="flex items-center gap-2 text-xs font-bold" style={{ color: COLORS.accent }}>
            <Sparkles size={14} strokeWidth={2} color="#F59E0B" />
            Live Insights
          </div>
          <span
            className="text-[10px] font-semibold px-2 py-0.5 rounded-full"
            style={{ background: "var(--color-highlight-btn-bg)", color: COLORS.accent }}
          >
            From live data
          </span>

        </div>

        {!input ? (
          <p className="text-xs" style={{ color: "var(--color-highlight-text-soft)" }}>
            Loading…
          </p>
        ) : (
          <ul className="space-y-2.5 text-xs leading-relaxed" style={{ color: "var(--color-highlight-text-soft)" }}>
            {bullets.map((b, i) => (
              <li key={i} className="flex items-start gap-2">
                <span style={{ color: COLORS.accent }} className="font-bold">
                  •
                </span>
                <span>{b}</span>
              </li>
            ))}
          </ul>
        )}
      </div>

      <div className="flex gap-2 mt-4">
        <button
          onClick={copy}
          disabled={!input}
          className="flex-1 py-2.5 rounded-xl text-xs font-bold flex items-center justify-center gap-2 transition active:scale-95 disabled:opacity-50"
          style={{ background: "var(--color-highlight-btn-bg)", color: "var(--color-highlight-text)" }}
        >
          {copied ? <CheckCircle2 size={13} /> : <Copy size={13} />}
          {copied ? "Copied" : "Copy Insights"}
        </button>
      </div>
    </div>
  );
}

function AiNarrativeCard({
  facts,
  anomalies,
  context,
  result,
  onGenerated,
}: {
  facts: AiNarrativeFact[];
  anomalies?: AiAnomaly[];
  context?: Record<string, number | string>;
  result: AiNarrativeResult | undefined;
  onGenerated: (result: AiNarrativeResult) => void;
}) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  async function generate() {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch("/api/ai-narrative", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ facts, anomalies: anomalies ?? [], context: context ?? null }),
      });
      const json = await res.json();
      if (json.error) {
        setError(json.error);
      } else {
        onGenerated(json);
      }
    } catch {
      setError("Could not reach the AI narrative endpoint.");
    } finally {
      setLoading(false);
    }
  }

  function copy() {
    if (!result) return;
    const parts = [
      result.overview,
      ...result.keyObservations,
      ...result.alerts.map((a) => `Alert: ${a.label} — ${a.detail}`),
      ...result.recommendations.map((r) => `Recommendation: ${r}`),
    ];
    navigator.clipboard.writeText(parts.join("\n"));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  return (
    <div
      className="rounded-2xl p-5 flex flex-col justify-between"
      style={{
        background: COLORS.surface,
        border: `1px solid ${COLORS.indigo}`,
        boxShadow: SHADOW,
        backdropFilter: "blur(14px)",
        WebkitBackdropFilter: "blur(14px)",
      }}
    >
      <div>
        <div
          className="flex items-center justify-between pb-3 mb-3"
          style={{ borderBottom: `1px solid ${COLORS.line}` }}
        >
          <div className="flex items-center gap-2 text-xs font-bold" style={{ color: COLORS.indigo }}>
            <Sparkles size={14} strokeWidth={2} color="#059669" className="ob-sparkle-twinkle" />
            AI Narrative
          </div>
          <span
            className="text-[10px] font-semibold px-2 py-0.5 rounded-full"
            style={{ background: COLORS.indigoSoft, color: COLORS.indigo }}
          >
            LLM Engine
          </span>
        </div>

        {error && (
          <p className="text-xs mb-2" style={{ color: COLORS.down }}>
            {error}
          </p>
        )}

        <div>
          {result ? (
            <div className="space-y-3">
              <p className="text-xs leading-relaxed" style={{ color: COLORS.inkSoft }}>
                {result.overview}
              </p>

              {result.keyObservations.length > 0 && (
                <ul className="space-y-1.5 text-xs leading-relaxed" style={{ color: COLORS.inkSoft }}>
                  {result.keyObservations.map((o, i) => (
                    <li key={i} className="flex items-start gap-2">
                      <span style={{ color: COLORS.indigo }} className="font-bold">•</span>
                      <span>{o}</span>
                    </li>
                  ))}
                </ul>
              )}

              {result.alerts.length > 0 && (
                <div className="pt-2" style={{ borderTop: `1px solid ${COLORS.line}` }}>
                  <p className="text-[10px] font-bold uppercase tracking-wider mb-1.5" style={{ color: COLORS.amberInk }}>
                    Watch Items
                  </p>
                  {result.alerts.map((a, i) => (
                    <p key={i} className="text-xs leading-relaxed mb-1" style={{ color: COLORS.inkSoft }}>
                      <span style={{ fontWeight: 700 }}>{a.label}:</span> {a.detail}
                    </p>
                  ))}
                </div>
              )}

              {result.recommendations.length > 0 && (
                <div className="pt-2" style={{ borderTop: `1px solid ${COLORS.line}` }}>
                  <p className="text-[10px] font-bold uppercase tracking-wider mb-1.5" style={{ color: COLORS.indigo }}>
                    Suggested Next Steps
                  </p>
                  {result.recommendations.map((r, i) => (
                    <p key={i} className="text-xs leading-relaxed mb-1" style={{ color: COLORS.inkSoft }}>
                      {r}
                    </p>
                  ))}
                </div>
              )}
            </div>
          ) : (
            <p className="text-xs" style={{ color: COLORS.inkFaint }}>
              Synthesizes the metrics already shown above into a structured summary — no new
              data, generated only when you click below.
            </p>
          )}
        </div>
        
      </div>

      <div className="flex gap-2 mt-4">
        <button
          onClick={generate}
          disabled={loading}
          className="flex-1 py-2.5 rounded-xl text-xs font-bold flex items-center justify-center gap-2 transition active:scale-95 disabled:opacity-60"
          style={{ background: COLORS.indigoSoft, color: COLORS.indigo }}
        >
          <Sparkles size={13} color="#059669" className="ob-sparkle-twinkle" />
          {loading ? "Generating…" : result ? "Regenerate" : "Generate AI Summary"}
        </button>
        {result && (
          <button
            onClick={copy}
            className="py-2.5 px-3 rounded-xl text-xs font-bold flex items-center justify-center gap-2 transition active:scale-95"
            style={{ background: COLORS.track, color: COLORS.inkSoft }}
          >
            {copied ? <CheckCircle2 size={13} /> : <Copy size={13} />}
          </button>
        )}
      </div>
    </div>
  );
}

/* ============== Executive ============== */

function NewVsReturningBar({ data }: { data: { segment: string; activeUsers: number }[] }) {
  const total = data.reduce((sum, d) => sum + d.activeUsers, 0);
  const segLabel = (s: string) => (s === "new" ? "New" : s === "returning" ? "Returning" : s);
  const segColor = (s: string, i: number) =>
    s === "new" ? COLORS.accent : s === "returning" ? COLORS.blue : DONUT_COLORS[i % DONUT_COLORS.length];

  return (
    <Card title="New vs. returning users" icon={Users}>
      {total === 0 ? (
        <p className="text-sm" style={{ color: COLORS.inkFaint }}>
          No data in this range.
        </p>
      ) : (
        <>
          <div className="h-2.5 rounded-full overflow-hidden flex" style={{ background: COLORS.track }}>
            {data.map((d, i) => (
              <div
                key={d.segment}
                style={{ width: `${(d.activeUsers / total) * 100}%`, background: segColor(d.segment, i) }}
              />
            ))}
          </div>
          <div className="flex flex-wrap gap-4 mt-3">
            {data.map((d, i) => (
              <div key={d.segment} className="flex items-center gap-1.5 text-xs font-medium" style={{ color: COLORS.inkSoft }}>
                <span className="w-2 h-2 rounded-full shrink-0" style={{ background: segColor(d.segment, i) }} />
                {segLabel(d.segment)} · {((d.activeUsers / total) * 100).toFixed(0)}%
              </div>
            ))}
          </div>
        </>
      )}
    </Card>
  );
}

// On-screen counterpart to the PDF's cross-department Watch Items (see
// buildCrossDeptWatchGroups) — same groups, same severity→color mapping
// (anomalySeverityToStatus), just rendered as Tailwind/JSX instead of
// react-pdf primitives. Renders nothing when there's nothing to show —
// no "no issues" placeholder invented.
function CrossDeptWatchSummary({ groups }: { groups: DeptWatchGroup[] }) {
  if (groups.length === 0) return null;
  return (
    <Card title="Cross-Department Watch Items" icon={FileWarning}>
      <div className="space-y-5">
        {groups.map((g) => (
          <div key={g.department}>
            <div className="flex items-center justify-between mb-1.5">
              <span className="text-xs font-bold uppercase tracking-wider" style={{ color: COLORS.ink }}>
                {g.label}
              </span>
            </div>
            {g.anomalies.map((a) => {
              const status = anomalySeverityToStatus(a.severity);
              const color = status === "critical" ? COLORS.down : COLORS.amberInk;
              const bg = status === "critical" ? COLORS.downSoft : COLORS.amberSoft;
              return (
                <div
                  key={a.id}
                  className="rounded-lg px-3 py-2 mt-2 text-xs"
                  style={{ background: bg, borderLeft: `3px solid ${color}` }}
                >
                  <span className="font-semibold" style={{ color: COLORS.ink }}>
                    {a.label}
                  </span>{" "}
                  <span style={{ color: COLORS.inkSoft }}>{a.detail}</span>
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </Card>
  );
}

function ExecutiveView({
  range,
  onData,
  onNarrative,
  narrative,
  marketingData,
  operationsData,
}: {
  range: Range;
  onData?: (data: ExecutiveData, clientMs?: number) => void;
  onNarrative?: (result: AiNarrativeResult) => void;
  narrative?: AiNarrativeResult;
  marketingData?: MarketingData;
  operationsData?: OperationsData;
}) {

  const { data, error, isFetching, clientMs } = useGa4<ExecutiveData>("/api/ga4/executive", range);

  useEffect(() => {
    if (data) onData?.(data, clientMs);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data]);

  if (error) return <ErrorBox message={error} />;
  if (!data) return <ExecutiveSkeleton />;

  const trend = data.trend.map((row) => ({ ...row, label: formatDate(row.date) }));
  const activeUsersDelta = data.deltas?.activeUsers ?? null;
  const sessionsDelta = data.deltas?.sessions ?? null;
  const totalChannelSessions = data.channels.reduce((s, c) => s + c.sessions, 0);
  const topChannelShare =
    totalChannelSessions > 0 ? ((data.channels[0]?.sessions ?? 0) / totalChannelSessions) * 100 : null;
  const maxChannelSessions = data.channels[0]?.sessions || 1;

  const totalGeoUsers = (data.geoCountries ?? []).reduce((s, c) => s + c.activeUsers, 0);

  // Cross-department facts for the Executive AI Narrative — Option B:
  // KPI cards above stay untouched, only the AI's input gets richer.
  const executiveFacts = executiveCrossDeptFacts(data, range, marketingData, operationsData);
  const executiveAnomalies: AiAnomaly[] = [];
  const executiveContext: Record<string, number | string> | undefined = undefined;

  return (
    
    <div>
      <RefreshingNote show={isFetching} />

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <KpiCard
          icon={Users}
          iconColor={COLORS.accent}
          iconBg={COLORS.accentSoft}
          label="Active Users"
          value={data.summary.activeUsers}
          sparklineData={trend.map((r) => ({ label: r.label, value: r.activeUsers }))}
          sparklineColor={COLORS.accent}
          delta={activeUsersDelta}
        />
        <KpiCard
          icon={Activity}
          iconColor={COLORS.blue}
          iconBg={COLORS.blueSoft}
          label="Sessions"
          value={data.summary.sessions}
          sparklineData={trend.map((r) => ({ label: r.label, value: r.sessions }))}
          sparklineColor={COLORS.blue}
          delta={sessionsDelta}
        />
        <KpiCard
          icon={TrendingUp}
          iconColor={COLORS.teal}
          iconBg="#CCFBF1"
          label="Engagement Rate"
          value={data.summary.engagementRate * 100}
          decimals={0}
          suffix="%"
        />
        <KpiCard
          icon={Globe2}
          iconColor={COLORS.indigo}
          iconBg={COLORS.indigoSoft}
          label="Top Channel"
          value={0}
          displayValue={data.summary.topChannel}
          caption={topChannelShare !== null ? `${topChannelShare.toFixed(0)}% total traffic share` : undefined}
        />
      </div>

      <div id="ob-chart-executive" className="mb-4">
        <Card
          title={`Live Sessions Trend · ${rangeLabel(range)}`}
          icon={TrendingUp}
          badge={activeUsersDelta !== null ? <DeltaBadge value={activeUsersDelta} /> : undefined}
        >
          <ResponsiveContainer width="100%" height={220}>
            <AreaChart data={trend} margin={{ top: 4, right: 4, bottom: 0, left: -16 }}>
              <defs>
                <linearGradient id="ob-gradient-active-users" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={COLORS.accent} stopOpacity={0.25} />
                  <stop offset="100%" stopColor={COLORS.accent} stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={COLORS.line} vertical={false} strokeDasharray="3 3" />
              <XAxis
                dataKey="label"
                fontSize={11}
                stroke={COLORS.inkFaint}
                tickLine={false}
                axisLine={false}
                interval={tickInterval(trend.length)}
              />
              <YAxis fontSize={11} stroke={COLORS.inkFaint} tickLine={false} axisLine={false} allowDecimals={false} />
              <Tooltip content={<ChartTooltip />} cursor={{ stroke: COLORS.line }} />
              <Area
                type="monotone"
                dataKey="activeUsers"
                stroke={COLORS.accent}
                strokeWidth={2}
                fill="url(#ob-gradient-active-users)"
                dot={false}
                activeDot={{ r: 4, fill: COLORS.accent, stroke: "#fff", strokeWidth: 2 }}
              />
            </AreaChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4">
        <Card title={`Sessions by Channel · ${rangeLabel(range)}`} icon={Radio}>
          <div className="space-y-3">
            {data.channels.map((c) => {
              const pct = Math.max(4, (c.sessions / maxChannelSessions) * 100);
              return (
                <div key={c.channel}>
                  <div className="flex items-center justify-between text-xs mb-1 font-medium">
                    <span style={{ color: COLORS.ink }}>{c.channel}</span>
                    <span style={{ color: COLORS.ink, fontVariantNumeric: "tabular-nums" }}>
                      {c.sessions.toLocaleString()}
                    </span>
                  </div>
                  <div className="h-1.5 rounded-full" style={{ background: COLORS.track }}>
                    <div className="h-1.5 rounded-full" style={{ width: `${pct}%`, background: COLORS.accent }} />
                  </div>
                </div>
              );
            })}
          </div>
        </Card>

        {data.newVsReturning && data.newVsReturning.length > 0 ? (
          <NewVsReturningBar data={data.newVsReturning} />
        ) : (
          <Card title="New vs. returning users" icon={Users}>
            <p className="text-sm" style={{ color: COLORS.inkFaint }}>
              No data in this range.
            </p>
          </Card>
        )}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-5 gap-4">
        <Card title="Geographical Telemetry Map" icon={Globe2} className="lg:col-span-3">
          <CountryChoropleth rows={data.geoCountries ?? []} />
        </Card>

        <BarListChart
          title={`Active Metro Cities · ${rangeLabel(range)}`}
          icon={MapPin}
          rows={(data.geoCities ?? []).map((g) => ({ label: g.city, value: g.activeUsers }))}
          barColor={COLORS.accent}
          className="lg:col-span-2"
        />
      </div>

      <div className="mt-4">
        <CrossDeptWatchSummary
          groups={buildCrossDeptWatchGroups(marketingData, operationsData)}
        />
      </div>

      <div className="mt-4">
        <AiNarrativeCard
          facts={executiveFacts}
          anomalies={executiveAnomalies}
          context={executiveContext}
          result={narrative}
          onGenerated={(result) => onNarrative?.(result)}
        />
      </div>
    </div>
  );
}

/* ============== Marketing ============== */

function FunnelVisualizer({
  funnel,
  range,
}: {
  funnel: { sessions: number; engagedSessions: number; keyEvents: number | null };
  range: Range;
}) {
  const steps = [
    { label: "Sessions", value: funnel.sessions, color: COLORS.blue },
    { label: "Engaged sessions", value: funnel.engagedSessions, color: COLORS.teal },
  ];
  if (funnel.keyEvents !== null) {
    steps.push({ label: "Key events", value: funnel.keyEvents, color: COLORS.inkFaint });
  }
  const max = steps[0]?.value || 1;

  const usingFallback = funnel.keyEvents === 0;
  const completionBasis = usingFallback ? funnel.engagedSessions : steps[steps.length - 1]?.value ?? 0;
  const completionRate = funnel.sessions > 0 ? (completionBasis / funnel.sessions) * 100 : 0;

  return (
    <Card
      title={`Conversion Funnel · ${rangeLabel(range)}`}
      icon={Filter}
      badge={
        <span
          className="text-xs font-bold px-2.5 py-1 rounded-lg"
          style={{ background: COLORS.accentSoft, color: COLORS.accent, fontVariantNumeric: "tabular-nums" }}
        >
          {completionRate.toFixed(1)}% completion
        </span>
      }
    >
      <div className="space-y-3">
        {steps.map((s, i) => {
          const pct = s.value === 0 ? 0 : Math.max(4, (s.value / max) * 100);
          return (
            <div key={s.label}>
              <div className="flex items-center justify-between text-xs mb-1 font-medium">
                <span style={{ color: COLORS.ink }}>
                  {i + 1}. {s.label}
                </span>
                <span style={{ color: COLORS.ink, fontVariantNumeric: "tabular-nums" }}>
                  {s.value.toLocaleString()}
                </span>
              </div>
              <div className="h-2.5 rounded-full" style={{ background: COLORS.track }}>
                <div className="h-2.5 rounded-full" style={{ width: `${pct}%`, background: s.color }} />
              </div>
            </div>
          );
        })}
      </div>
      {funnel.keyEvents === null && (
        <p className="text-xs mt-3" style={{ color: COLORS.inkFaint }}>
          No key events configured in this GA4 property yet — showing the two funnel steps available today.
        </p>
      )}
      {usingFallback && (
        <p className="text-xs mt-3" style={{ color: COLORS.inkFaint }}>
          No key events recorded in this window, so completion above reflects engaged sessions instead.
        </p>
      )}
    </Card>
  );
}

function ChannelRadar({ channels }: { channels: { channel: string; sessions: number }[] }) {
  if (channels.length < 3) {
    return (
      <Card title="Channel Mix" icon={Compass}>
        <p className="text-sm" style={{ color: COLORS.inkFaint }}>
          Need at least 3 channels with traffic to plot a radar shape — currently {channels.length}.
        </p>
      </Card>
    );
  }
  return (
    <Card title="Channel Mix" icon={Compass}>
      <ResponsiveContainer width="100%" height={220}>
        <RadarChart data={channels}>
          <PolarGrid stroke={COLORS.line} />
          <PolarAngleAxis dataKey="channel" tick={{ fontSize: 10, fill: COLORS.inkSoft }} />
          <PolarRadiusAxis tick={false} axisLine={false} />
          <Radar dataKey="sessions" stroke={COLORS.accent} fill={COLORS.accent} fillOpacity={0.35} />
          <Tooltip content={<ChartTooltip />} />
        </RadarChart>
      </ResponsiveContainer>
    </Card>
  );
}

function MarketingView({
  range,
  onData,
  onNarrative,
  narrative,
}: {
  range: Range;
  onData?: (data: MarketingData, clientMs?: number) => void;
  onNarrative?: (result: AiNarrativeResult) => void;
  narrative?: AiNarrativeResult;
}) {
  const { data, error, isFetching, clientMs } = useGa4<MarketingData>("/api/ga4/marketing", range);

  useEffect(() => {
    if (data) onData?.(data, clientMs);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data]);

  if (error) return <ErrorBox message={error} />;
  if (!data) return <MarketingSkeleton />;

  const totalSessions = data.sources.reduce((s, x) => s + x.sessions, 0);
  const topSource = data.sources[0];
  const topSourceShare = totalSessions > 0 && topSource ? (topSource.sessions / totalSessions) * 100 : null;
  const topPage = data.pages[0];
  const topLanding = data.landingPages[0];
  const sessionsDelta = data.deltas?.sessions ?? null;

  return (
    <div>
      <RefreshingNote show={isFetching} />

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <KpiCard
          icon={Activity}
          iconColor={COLORS.accent}
          iconBg={COLORS.accentSoft}
          label="Total Traffic Sessions"
          value={totalSessions}
          delta={sessionsDelta}
          caption={sessionsDelta === null && topSource ? `via ${topSource.source}` : undefined}
        />
        <KpiCard
          icon={Compass}
          iconColor={COLORS.indigo}
          iconBg={COLORS.indigoSoft}
          label="Top Source"
          value={0}
          displayValue={topSource ? topSource.source : "—"}
          caption={topSourceShare !== null ? `${topSourceShare.toFixed(0)}% overall share` : undefined}
        />
        <KpiCard
          icon={FileText}
          iconColor={COLORS.blue}
          iconBg={COLORS.blueSoft}
          label="Top Page Views"
          value={topPage ? topPage.views : 0}
          caption={topPage ? topPage.title : undefined}
        />
        <KpiCard
          icon={Flag}
          iconColor={COLORS.teal}
          iconBg="#CCFBF1"
          label="Landing Page Views"
          value={topLanding ? topLanding.sessions : 0}
          caption={topLanding ? topLanding.page : undefined}
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-4">
        <div id="ob-chart-marketing"className="lg:col-span-2">
          {data.funnel ? (
            <FunnelVisualizer funnel={data.funnel} range={range} />
          ) : (
            <Card title={`Conversion Funnel · ${rangeLabel(range)}`} icon={Filter}>
              <p className="text-sm" style={{ color: COLORS.inkFaint }}>
                No data in this range.
              </p>
            </Card>
          )}
        </div>
        <ChannelRadar channels={data.channels ?? []} />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <ShareDonut
          title={`Traffic Source · ${rangeLabel(range)}`}
          icon={PieChartIcon}
          rows={data.sources.map((s) => ({ label: s.source, value: s.sessions }))}
          valueLabel="sessions"
          colors={[COLORS.accent, COLORS.inkFaint, COLORS.blue, COLORS.indigo, COLORS.teal, "#CBD5E1"]}
        />
        <BarListChart
          title={`Landing Page Acquisition · ${rangeLabel(range)}`}
          icon={Flag}
          barColor={COLORS.teal}
          rows={data.landingPages.map((p) => ({ label: p.page, value: p.sessions }))}
        />
      </div>

      <div className="mt-4">
        {(() => {
          const { anomalies, context } = marketingAnomaliesAndContext(data);
          return (
            <AiNarrativeCard
              facts={marketingToBriefing(data, range).kpis}
              anomalies={anomalies}
              context={context}
              result={narrative}
              onGenerated={(result) => onNarrative?.(result)}
            />
          );
        })()}
      </div>
    </div>
  );
}



/* ============== Operations ============== */

function OperationsView({
  range,
  onData,
  onNarrative,
  narrative,
}: {
  range: Range;
  onData?: (data: OperationsData, clientMs?: number) => void;
  onNarrative?: (result: AiNarrativeResult) => void;
  narrative?: AiNarrativeResult;
}) {

  const { data, error, isFetching, clientMs } = useGa4<OperationsData>("/api/ga4/operations", range);

  useEffect(() => {
    if (data) onData?.(data, clientMs);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data]);

  if (error) return <ErrorBox message={error} />;
  if (!data) return <OperationsSkeleton />;

  const deviceRows = (data.devices ?? []).map((d) => ({
    label: d.device.charAt(0).toUpperCase() + d.device.slice(1),
    value: d.sessions,
  }));
  const browserRows = (data.browsers ?? []).map((b) => ({ label: b.browser, value: b.sessions }));
  const osRows = (data.operatingSystems ?? []).map((o) => ({ label: o.os, value: o.sessions }));
  const hasDeviceCharts = deviceRows.length > 0 || browserRows.length > 0 || osRows.length > 0;

  const quotaTotal = data.quota ? data.quota.consumed + data.quota.remaining : null;
  const quotaPct = quotaTotal && quotaTotal > 0 ? (data.quota!.consumed / quotaTotal) * 100 : null;

  return (
    <div>
      <RefreshingNote show={isFetching} />

      {/* 1. Top row — headline metric cards */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-6">
        <KpiCard
          icon={Gauge}
          iconColor={COLORS.accent}
          iconBg={COLORS.accentSoft}
          label="Bounce Rate"
          value={data.summary.bounceRate * 100}
          decimals={0}
          suffix="%"
        />
        <KpiCard
          icon={Clock}
          iconColor={COLORS.blue}
          iconBg={COLORS.blueSoft}
          label="Avg. Engagement Duration"
          value={0}
          displayValue={formatDuration(data.summary.avgSessionDuration)}
        />
        {data.quota ? (
          <KpiCard
            icon={Cpu}
            iconColor={COLORS.teal}
            iconBg="#CCFBF1"
            label="API Rate Quota"
            value={data.quota.consumed}
            caption={
              quotaPct !== null
                ? `${quotaPct.toFixed(2)}% of hourly quota used`
                : `${data.quota.remaining.toLocaleString()} tokens remaining`
            }
          />
        ) : (
          <KpiCard
            icon={Cpu}
            iconColor={COLORS.inkFaint}
            iconBg={COLORS.track}
            label="API Rate Quota"
            value={0}
            displayValue="—"
            caption="Quota data unavailable this request"
          />
        )}
      </div>

      {/* 2. Middle — device / browser / OS breakdown charts */}
      {hasDeviceCharts && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-4">
          <div id="ob-chart-operations">
            <ShareDonut title={`Device Category · ${rangeLabel(range)}`} icon={Smartphone} rows={deviceRows} />
          </div>
          <BarListChart
            title={`Browser Matrix · ${rangeLabel(range)}`}
            icon={Layers}
            barColor={COLORS.blue}
            rows={browserRows}
          />
          <BarListChart
            title={`Operating System · ${rangeLabel(range)}`}
            icon={Cpu}
            barColor={COLORS.indigo}
            rows={osRows}
          />
        </div>
      )}

      {/* 3. Bottom — general page traffic on the left, diagnostic
          route logs on the right, presented as routine monitoring data. */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <BarListChart
          title={`Top Content Pages · ${rangeLabel(range)}`}
          icon={FileText}
          barColor={COLORS.blue}
          rows={(data.pages ?? []).map((p) => ({ label: p.title, value: p.views }))}
        />

        <div
          className="rounded-2xl p-5 transition-shadow duration-200"
          style={{
            background: COLORS.surface,
            border: `1px solid ${COLORS.line}`,
            boxShadow: SHADOW,
            backdropFilter: "blur(14px)",
            WebkitBackdropFilter: "blur(14px)",
          }}
        >
          <div className="flex items-center justify-between mb-4 gap-2">
            <div
              className="flex items-center gap-2 min-w-0 px-2.5 py-1 rounded-lg"
              style={{ background: COLORS.amberSoft, border: `1px solid rgba(245,158,11,0.35)` }}
            >
              <FileWarning size={14} strokeWidth={2} color={COLORS.amberInk} />
              <p className="text-[11px] font-bold uppercase tracking-wider truncate" style={{ color: COLORS.amberInk }}>
                Diagnostic Route Logs
              </p>
            </div>
            <span className="text-[11px] font-medium" style={{ color: COLORS.accent }}>
              Live Endpoint Inspector
            </span>
          </div>
          {data.notFoundPages.length === 0 ? (
            <p className="text-sm" style={{ color: COLORS.inkFaint }}>
              No route issues detected in the {rangeLabel(range)}.
            </p>
          ) : (
            <div className="space-y-2.5">
              {data.notFoundPages.map((p) => {
                const highSeverity = p.views >= 5;
                return (
                  <div
                    key={p.title}
                    className="flex flex-col sm:flex-row sm:items-center gap-2 sm:gap-3 rounded-xl p-3"
                    style={{ border: `1px solid ${COLORS.line}` }}
                  >
                    <div className="flex items-center gap-3 min-w-0">
                      <span
                        className="shrink-0 px-2 py-1 rounded-lg text-[10px] font-bold"
                        style={{ background: COLORS.downSoft, color: COLORS.down, fontFamily: MONO_STACK }}
                      >
                        HTTP 404
                      </span>
                      <span
                        className="truncate text-xs font-medium"
                        style={{ color: COLORS.ink, fontFamily: MONO_STACK }}
                      >
                        {p.title}
                      </span>
                    </div>
                    <div className="flex items-center gap-3 sm:ml-auto shrink-0">
                      <span
                        className="text-xs font-medium"
                        style={{ color: COLORS.inkSoft, fontVariantNumeric: "tabular-nums" }}
                      >
                        {p.views} occurrences
                      </span>
                      <span
                        className="shrink-0 px-2 py-1 rounded-lg text-[10px] font-bold"
                        style={{
                          background: highSeverity ? COLORS.downSoft : COLORS.track,
                          color: highSeverity ? COLORS.down : COLORS.inkSoft,
                        }}
                      >
                        {highSeverity ? "High Severity" : "Low Severity"}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      <div className="mt-4">
        <AiNarrativeCard
          facts={operationsToBriefing(data, range).kpis}
          anomalies={operationsAnomalies(data)}
          result={narrative}
          onGenerated={(result) => onNarrative?.(result)}
        />
      </div>
    </div>
  );
}

/* ============== Custom Data ============== */
// One dedicated tab for any uploaded file, regardless of shape. No fixed
// department/metricKey schema: the file's columns are profiled in code
// (columnProfiler.ts), AI suggests a handful of KPI/chart specs from that
// profile alone (never the full dataset), then every number actually
// shown is computed in code from the FULL uploaded dataset
// (aggregationEngine.ts) — the AI only ever chose what to compute and how.

function TrendLineChart({ title, rows }: { title: string; rows: { label: string; value: number }[] }) {
  return (
    <Card title={title} icon={TrendingUp}>
      {rows.length === 0 ? (
        <p className="text-sm" style={{ color: COLORS.inkFaint }}>
          No data to chart.
        </p>
      ) : (
        <ResponsiveContainer width="100%" height={220}>
          <LineChart data={rows} margin={{ top: 4, right: 24, bottom: 0, left: -16 }}>
            <CartesianGrid stroke={COLORS.line} vertical={false} strokeDasharray="3 3" />
            <XAxis
              dataKey="label"
              fontSize={11}
              stroke={COLORS.inkFaint}
              tickLine={false}
              axisLine={false}
              interval={tickInterval(rows.length)}
              padding={{ right: 16 }}
            />
            <YAxis fontSize={11} stroke={COLORS.inkFaint} tickLine={false} axisLine={false} allowDecimals={false} />
            <Tooltip content={<ChartTooltip />} />
            <Line type="monotone" dataKey="value" stroke={COLORS.accent} strokeWidth={2} dot={rows.length <= 12} />
          </LineChart>
        </ResponsiveContainer>
      )}
    </Card>
  );
}

function CustomDataView({
  state,
  onLoadDemo,
  demoLoading,
}: {
  state: CustomDataState;
  onLoadDemo: () => void;
  demoLoading: boolean;
}) {
  if (!state.connected) {
    return (
      <div
        className="rounded-2xl p-10 flex flex-col items-center text-center gap-3"
        style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW }}
      >
        <FileSpreadsheet size={28} color={COLORS.inkFaint} />
        <p className="text-sm font-semibold" style={{ color: COLORS.ink }}>
          Upload a spreadsheet to auto-generate KPIs
        </p>
        <p className="text-xs max-w-[480px]" style={{ color: COLORS.inkFaint }}>
          OneBoard profiles the file's columns, asks AI to suggest a handful of KPI cards and
          charts, then computes every number in code from your full dataset — nothing is
          estimated or invented.
        </p>
        <p className="text-[11px] max-w-[480px]" style={{ color: COLORS.inkFaint }}>
          .csv, .xlsx, or .xls · needs a header row + at least one data row · only the first sheet is read
        </p>
        <button
          onClick={onLoadDemo}
          disabled={demoLoading}
          className="mt-2 text-xs font-bold px-4 py-2 rounded-xl flex items-center gap-2 transition active:scale-95 disabled:opacity-60"
          style={{ background: COLORS.accentSoft, color: COLORS.accent, border: "1px solid transparent" }}
        >
          <FileSpreadsheet size={14} />
          {demoLoading ? "Loading…" : "Load Sample Data"}
        </button>
      </div>
    );
  }

  if (state.suggesting) {
    return (
      <div>
        <SkeletonKpiRow count={4} />
        <SkeletonChartGrid count={2} />
      </div>
    );
  }

  if (state.suggestError) {
    return <ErrorBox message={state.suggestError} />;
  }

  if (!state.specs || state.specs.length === 0 || !state.data) {
    return (
      <p className="text-sm" style={{ color: COLORS.inkFaint }}>
        No KPI suggestions could be generated for this file.
      </p>
    );
  }

  // Real computation, every render: the AI's specs are only a suggestion
  // of what/how — the numbers themselves always come from the full parsed
  // dataset, never the small sample sent to /api/kpi-suggest.
  const results = state.specs.map((spec) => computeChartSpec(state.data!, spec, state.profile ?? undefined));
  const kpiResults = results.filter((r) => r.spec.chartType === "kpi");
  const chartResults = results.filter((r) => r.spec.chartType !== "kpi");

  return (
    <div>
      <p className="text-xs font-medium mb-4" style={{ color: COLORS.inkFaint }}>
        {state.rows} row{state.rows === 1 ? "" : "s"} from {state.filename} · every number below
        computed from the full file
      </p>

      {kpiResults.length > 0 && (
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
          {kpiResults.map((r, i) => {
            const isNumeric = typeof r.value === "number";
            return (
              <KpiCard
                key={i}
                icon={Gauge}
                iconColor={COLORS.accent}
                iconBg={COLORS.accentSoft}
                label={r.spec.title}
                value={isNumeric ? (r.value as number) : 0}
                decimals={isNumeric ? r.decimals : 0}
                displayValue={isNumeric ? undefined : r.value === null ? "—" : String(r.value)}
              />
            );
          })}
        </div>
      )}

      {chartResults.length > 0 && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {chartResults.map((r, i) => {
            if (r.spec.chartType === "bar") {
              return (
                <BarListChart
                  key={i}
                  title={r.spec.title}
                  icon={Layers}
                  barColor={COLORS.accent}
                  rows={r.series}
                  decimals={r.decimals === 2 ? 2 : undefined}
                />
              );
            }
            if (r.spec.chartType === "pie") {
              return <ShareDonut key={i} title={r.spec.title} icon={PieChartIcon} rows={r.series} valueLabel="value" />;
            }
            return <TrendLineChart key={i} title={r.spec.title} rows={r.series} />;
          })}
        </div>
      )}
    </div>
  );
}

/* ============== Export Briefing adapters ============== */
// Each function turns one tab's already-displayed data into the
// BriefingData shape ExecutiveBriefingPDF expects. Narrative text here is
// built from the same rule-based logic already on screen (buildInsightBullets
// for Executive, plain sentences for the others) — no new numbers, no AI
// call yet. This keeps the exported PDF honest: everything in it was
// already visible in the dashboard before export.

function deltaFromPercent(pct: number | null | undefined, period: string): BriefingFact["delta"] {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return undefined;
  const direction: "up" | "down" = pct >= 0 ? "up" : "down";
  return { direction, label: `${Math.abs(pct).toFixed(1)}% ${period}` };
}

function executiveToBriefing(
  data: ExecutiveData,
  range: Range,
  marketingData?: MarketingData,
  operationsData?: OperationsData
): BriefingData {
  const totalChannelSessions = data.channels.reduce((s, c) => s + c.sessions, 0);
  const topChannelShare =
    totalChannelSessions > 0 ? ((data.channels[0]?.sessions ?? 0) / totalChannelSessions) * 100 : null;
  const totalGeoUsers = (data.geoCountries ?? []).reduce((s, c) => s + c.activeUsers, 0);
  const topCountryShare =
    data.geoCountries && data.geoCountries.length > 0 && totalGeoUsers > 0
      ? (data.geoCountries[0].activeUsers / totalGeoUsers) * 100
      : null;

  const bullets = buildInsightBullets({
    range: rangeLabel(range),
    activeUsers: data.summary.activeUsers,
    sessions: data.summary.sessions,
    engagementRate: data.summary.engagementRate,
    topChannel: data.summary.topChannel,
    topChannelSharePct: topChannelShare,
    topCountry: data.geoCountries?.[0]?.country ?? null,
    topCountrySharePct: topCountryShare,
  });

  return {
    reportTitle: "OneBoard Executive Briefing",
    sectionLabel: "Executive View",
    dataSource: "GA4 Live",
    period: rangeLabel(range),
    generatedAt: new Date().toLocaleString(),
    kpis: [
      {
        id: "active_users",
        metric: "Active Users",
        value: data.summary.activeUsers,
        delta: deltaFromPercent(data.deltas?.activeUsers, "vs prior period"),
      },
      {
        id: "sessions",
        metric: "Sessions",
        value: data.summary.sessions,
        delta: deltaFromPercent(data.deltas?.sessions, "vs prior period"),
      },
      {
        id: "engagement_rate",
        metric: "Engagement Rate",
        value: Math.round(data.summary.engagementRate * 100),
        unit: "%",
      },
      {
        id: "top_channel",
        metric: "Top Channel",
        value: data.summary.topChannel,
        note: topChannelShare !== null ? `${topChannelShare.toFixed(0)}% traffic share` : undefined,
      },
    ],
    narrative: bullets.join(" "),
    risks: crossDeptWatchGroupsToRisks(buildCrossDeptWatchGroups(marketingData, operationsData)),
  };
}

function marketingToBriefing(data: MarketingData, range: Range): BriefingData {
  const totalSessions = data.sources.reduce((s, x) => s + x.sessions, 0);
  const topSource = data.sources[0];
  const topSourceShare = totalSessions > 0 && topSource ? (topSource.sessions / totalSessions) * 100 : null;

  const narrativeParts: string[] = [];
  if (topSource && topSourceShare !== null) {
    narrativeParts.push(`${topSource.source} is the top traffic source, at ${topSourceShare.toFixed(0)}% of sessions.`);
  }
  if (data.pages[0]) {
    narrativeParts.push(`The top content page is "${data.pages[0].title}" with ${data.pages[0].views.toLocaleString()} views.`);
  }
  narrativeParts.push(`Total traffic across ${rangeLabel(range)} was ${totalSessions.toLocaleString()} sessions.`);

  // Drilldown: full session distribution by source — the real dimension
  // behind the "Top Source" KPI's share note, broken out row by row instead
  // of collapsing to one number. Real GA4 sources only, no invented rows.
  // Neutral status throughout: there's no internal target for how traffic
  // *should* split across sources, so no row earns a color judgment —
  // same rule as the "Top Source" KPI itself (note, not statusLabel).
  const drilldown: DrilldownTable | undefined =
    data.sources.length > 0
      ? {
          title: "Session Distribution by Source",
          columns: ["Source", "Sessions", "Share of Total"],
          rows: data.sources.map((s) => ({
            cells: [s.source, s.sessions, totalSessions > 0 ? `${((s.sessions / totalSessions) * 100).toFixed(0)}%` : "—"],
            status: "neutral",
          })),
        }
      : undefined;

  return {
    reportTitle: "OneBoard Marketing Briefing",
    sectionLabel: "Marketing & Funnel",
    dataSource: "GA4 Live",
    period: rangeLabel(range),
    generatedAt: new Date().toLocaleString(),
    kpis: [
      { id: "total_sessions", metric: "Total Traffic Sessions", value: totalSessions },
      {
        id: "top_source",
        metric: "Top Source",
        value: topSource ? topSource.source : "—",
        note: topSourceShare !== null ? `${topSourceShare.toFixed(0)}% overall share` : undefined,
      },
      { id: "top_page_views", metric: "Top Page Views", value: data.pages[0]?.views ?? 0 },
      { id: "landing_page_views", metric: "Landing Page Views", value: data.landingPages[0]?.sessions ?? 0 },
    ],
    narrative: narrativeParts.join(" "),
    drilldown,
  };
}

function operationsToBriefing(data: OperationsData, range: Range): BriefingData {
  const hasIssues = data.notFoundPages.length > 0;
  const total404 = data.notFoundPages.reduce((s, p) => s + p.views, 0);

  const narrativeParts: string[] = [
    `Bounce rate for ${rangeLabel(range)} was ${Math.round(data.summary.bounceRate * 100)}%, with an average engagement duration of ${formatDuration(data.summary.avgSessionDuration)}.`,
  ];
  if (hasIssues) {
    narrativeParts.push(`${data.notFoundPages.length} route(s) returned HTTP 404 during this window, totaling ${total404} occurrences.`);
  } else {
    narrativeParts.push(`No route issues were detected in this window.`);
  }

  // route_issues previously carried no status at all (defaulted to
  // neutral/no color) while the Watch Item for the exact same number was
  // rendered in "watch" amber — two components disagreeing about the color
  // of one metric, the exact bug class CLAUDE.md's color-consistency rule
  // calls out. Computed once here and reused by both the KPI pill and the
  // drilldown table rows below — status computed once, never re-derived.
  const routeStatus: Status = hasIssues ? "watch" : "ok";
  const routeStatusLabel = hasIssues ? "Needs Review" : "Clear";

  // Drilldown: the 404 count broken out into the actual routes and their
  // occurrence counts (data.notFoundPages is already real GA4 page-path +
  // views data — see CLAUDE.md's "Marketing & Operations: CONFIRMED real
  // dimensional data" note). Same status as the KPI/watch item above.
  const drilldown: DrilldownTable | undefined = hasIssues
    ? {
        title: "Route Issues (404s) — Detail",
        columns: ["Route", "Occurrences"],
        rows: data.notFoundPages.map((p) => ({
          cells: [p.title, p.views],
          status: routeStatus,
        })),
      }
    : undefined;

  return {
    reportTitle: "OneBoard Operations Briefing",
    sectionLabel: "Operations & Diagnostics",
    dataSource: "GA4 Live",
    period: rangeLabel(range),
    generatedAt: new Date().toLocaleString(),
    kpis: [
      { id: "bounce_rate", metric: "Bounce Rate", value: Math.round(data.summary.bounceRate * 100), unit: "%" },
      { id: "avg_engagement", metric: "Avg. Engagement Duration", value: formatDuration(data.summary.avgSessionDuration) },
      {
        id: "route_issues",
        metric: "Route Issues (404s)",
        value: data.notFoundPages.length,
        status: routeStatus,
        statusLabel: routeStatusLabel,
      },
    ],
    narrative: narrativeParts.join(" "),
    drilldown,
  };
}

// Unlike every other tab's adapter, this one has no fixed set of named
// metrics to map — Custom Data's KPIs/charts are whatever AI suggested for
// THIS file's columns (see /api/kpi-suggest), so this must handle any
// number of specs with any titles, not a hardcoded list. Every number is
// still recomputed here from the full dataset (never read off state as a
// cached display string), same rule as CustomDataView on screen — the two
// surfaces call the exact same computeChartSpec, so they can't disagree.
function customDataToBriefing(state: CustomDataState): BriefingData {
  const rows = state.data ?? [];
  const specs = state.specs ?? [];
  const results = specs.map((spec) => computeChartSpec(rows, spec, state.profile ?? undefined));

  const kpiResults = results.filter((r) => r.spec.chartType === "kpi");
  const chartResults = results.filter((r) => r.spec.chartType !== "kpi");

  // Same display rule CustomDataView uses on screen: decimals is computed
  // once by the aggregation engine per spec, never re-derived here.
  const formatValue = (value: number, decimals: number) =>
    decimals > 0 ? value.toFixed(decimals) : formatInteger(value);

  const kpis: BriefingFact[] = kpiResults.map((r, i) => ({
    id: `custom_${i}`,
    metric: r.spec.title,
    // No real target/threshold exists for an arbitrary AI-suggested KPI —
    // status is deliberately omitted (defaults to neutral), never guessed.
    value:
      typeof r.value === "number"
        ? formatValue(r.value, r.decimals)
        : r.value === null
          ? "—"
          : r.value,
  }));

  // BriefingData has room for exactly one chart and one drilldown table —
  // every other tab's adapter already makes this same simplification (e.g.
  // Marketing's on-screen tab has four charts; its briefing picks one
  // drilldown table).
  //
  // The Chart component's bar rows use a small fixed-width label column
  // (~28pt — sized for 3-letter month abbreviations, the only labels it
  // was ever built to show). A category name that runs long (e.g. a course
  // title) wraps into an unreadable multi-line mess there. The drilldown
  // table's cell is a flexible flow layout that handles long text fine. So:
  // prefer a "line" spec (its groups are always date-bucketed to "YYYY-MM"
  // or "YYYY" — always short) for the chart slot; only use a bar/pie spec
  // there if every one of its labels is actually short enough to fit.
  // Anything not picked for chart still gets a shot at the drilldown table.
  const hasOnlyShortLabels = (r: (typeof chartResults)[number]) => r.series.every((p) => p.label.length <= 14);
  const chartResult =
    chartResults.find((r) => r.spec.chartType === "line" && r.series.length > 0) ??
    chartResults.find(
      (r) => (r.spec.chartType === "bar" || r.spec.chartType === "pie") && r.series.length > 0 && hasOnlyShortLabels(r)
    );
  const chart: ComparisonChart | undefined = chartResult
    ? {
        title: chartResult.spec.title,
        points: chartResult.series.map((p) => ({
          label: p.label,
          value: p.value,
          display: formatValue(p.value, chartResult.decimals),
        })),
      }
    : undefined;

  const drilldownResult = chartResults.find((r) => r !== chartResult && r.series.length > 0);
  const drilldown: DrilldownTable | undefined = drilldownResult
    ? {
        title: drilldownResult.spec.title,
        columns: [drilldownResult.spec.groupByColumn ?? "Category", "Value"],
        rows: drilldownResult.series.map((p) => ({
          cells: [p.label, formatValue(p.value, drilldownResult.decimals)],
          status: "neutral" as Status,
        })),
      }
    : undefined;

  const rowCount = rows.length;
  const narrative =
    `This report summarizes ${kpis.length} KPI${kpis.length === 1 ? "" : "s"}` +
    (chartResults.length > 0
      ? ` and ${chartResults.length} chart breakdown${chartResults.length === 1 ? "" : "s"}`
      : "") +
    `, computed from all ${rowCount} row${rowCount === 1 ? "" : "s"} of ${state.filename ?? "the uploaded file"}.`;

  return {
    reportTitle: "OneBoard Custom Data Briefing",
    sectionLabel: "Custom Data",
    dataSource: state.filename ?? "Uploaded file",
    period: `${rowCount} row${rowCount === 1 ? "" : "s"} · full dataset`,
    generatedAt: new Date().toLocaleString(),
    kpis,
    narrative,
    chart,
    drilldown,
  };
}

// Converts the AI's alerts ({label, detail}, no status — see route.ts's
// AiNarrativeResult) into BriefingRisk[] by matching each alert back to
// the KPI it's describing and reusing THAT KPI's already-computed status
// (e.g. operationsToBriefing's routeStatus). This is the actual mechanism
// behind the color-consistency rule: status is computed once and every
// other place it appears just looks it up — nothing recomputes it
// independently, so nothing can disagree with it.
function alertsToRisks(
  alerts: { label: string; detail: string }[],
  kpis: BriefingFact[]
): BriefingRisk[] {
  return alerts.map((a) => {
    const text = `${a.label} ${a.detail}`.toLowerCase();
    const matched = kpis.find((k) => text.includes(String(k.id).toLowerCase()));
    return {
      label: a.label,
      detail: a.detail,
      // An alert existing at all means something was flagged. If no KPI
      // match is found, "watch" is the safe floor — never silently "ok".
      status: matched?.status ?? "watch",
    };
  });
}

function marketingAnomaliesAndContext(data: MarketingData) {
  const trendSeries = (data.trend ?? []).map((t) => ({
    label: formatDate(t.date),
    value: t.sessions,
  }));

  const anomalies = detectAnomalies(trendSeries, "Sessions", {
    trendMinStreak: 3,
  });

  const totalSessions = data.sources.reduce((s, x) => s + x.sessions, 0);
  const ranked = [...data.sources].sort((a, b) => b.sessions - a.sessions);

  const context: Record<string, number | string> = { totalSessions };
  if (ranked[0]) {
    context.topSourceName = ranked[0].source;
    context.topSourceSessions = ranked[0].sessions;
  }
  if (ranked[1]) {
    context.secondSourceName = ranked[1].source;
    context.secondSourceSessions = ranked[1].sessions;
  }

  return { anomalies, context };
}

// Auto-detects whether an anomaly signal describes an objective technical
// fault (broken link, server error, etc.) — based on well-known technical
// markers in the text, not a subjective judgment call. Runs in code, at
// signal-creation time, not via the AI — so classification is identical
// every time, and nobody has to remember to hand-tag each new signal.
const TECHNICAL_ISSUE_PATTERN = /\b(4\d{2}|5\d{2})\b|\berror\b|\bfailed\b|\btimeout\b|\bbroken\b|\bnot found\b/i;

function classifyAnomalyType(
  label: string,
  detail: string,
  fallback: AiAnomaly["type"]
): AiAnomaly["type"] {
  const text = `${label} ${detail}`;
  return TECHNICAL_ISSUE_PATTERN.test(text) ? "technical_issue" : fallback;
}

// Operations has no time series to run detectAnomalies against, but a
// recorded 404 IS a real, already-computed signal — not something that
// needs inventing. The severity threshold (>=5 occurrences) mirrors the
// exact rule OperationsView already uses to render "High/Low Severity"
// badges on screen, so this isn't a new judgment call, just the same one
// also passed to the AI.
function operationsAnomalies(data: OperationsData): AiAnomaly[] {
  if (data.notFoundPages.length === 0) return [];

  const totalOccurrences = data.notFoundPages.reduce((s, p) => s + p.views, 0);
  const highSeverity = data.notFoundPages.length > 0;
  const label = "Route errors detected";
  const detail = `${data.notFoundPages.length} route(s) returned HTTP 404 during this window, totaling ${totalOccurrences} occurrences.`;

  return [
    {
      id: "operations-404-routes",
      type: classifyAnomalyType(label, detail, "target_breach"),
      label,
      detail,
      severity: highSeverity ? "warning" : "info",
    },
  ];
}

async function exportBriefingPDF(briefing: BriefingData) {
  const [{ pdf, Font }, { ExecutiveBriefingPDF }] = await Promise.all([
    import("@react-pdf/renderer"),
    import("../components/ExecutiveBriefingPDF"),
  ]);

  // Reported bug: a handful of KPI values render short by their leading
  // character(s) ("3.2%" -> ".2%", "16m 2s" -> "16 2s") — confirmed for
  // real on user-submitted PDFs (their raw text runs are genuinely short
  // those characters, not a viewer/clipping artifact) and reproduced
  // directly, but only once in ~19 real-Chrome export attempts against
  // otherwise-identical code and data. Ruled out: font glyphs (verified
  // intact via fontTools), container width (measured the actual embedded
  // glyph advances at each Text's real font size — "94.1%" needs ~59pt,
  // the narrowest box here has ~93pt available, nowhere close to
  // overflowing), and plain repetition/HMR reload (15-30+ clean exports
  // in various stress runs). The one thing that did track with the one
  // reproduction: it was on Chrome 151 (the user's real browser), never
  // on Playwright's pinned Chrome-for-Testing 143 build — consistent
  // with a rare async timing race rather than a deterministic bug.
  //
  // `Font.register()` in ExecutiveBriefingPDF.tsx only queues the custom
  // Inter sources — the actual fetch+parse of the .woff files is async
  // and, per @react-pdf/font, lazy: it resolves on first use inside the
  // layout pass this same render triggers. Explicitly awaiting `Font.load`
  // for every weight this component actually uses (400/700/800) before
  // starting layout closes that window outright — if this was a
  // load-still-in-flight race, there is nothing left to race once these
  // promises have resolved. This is a targeted concurrency fix, not
  // another font-size/width guess (both already tried, per the git
  // history on this file, and neither one is what closed this).
  await Promise.all(
    [400, 700, 800].map((fontWeight) => Font.load({ fontFamily: "Inter", fontWeight }))
  );

  const blob = await pdf(<ExecutiveBriefingPDF data={briefing} />).toBlob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `OneBoard-${briefing.sectionLabel.replace(/\s+/g, "-")}-Briefing-${new Date().toISOString().slice(0, 10)}.pdf`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Picks only the fields AiNarrativeFact actually declares — BriefingFact
// now also carries icon/status/statusLabel (PDF-only presentation fields),
// which have no meaning to the AI narrative endpoint.
function toNarrativeFact(f: BriefingFact, idPrefix: string, metricPrefix: string): AiNarrativeFact {
  return {
    id: `${idPrefix}_${f.id}`,
    metric: `${metricPrefix} — ${f.metric}`,
    value: f.value,
    unit: f.unit,
    delta: f.delta,
  };
}

function executiveCrossDeptFacts(
  execData: ExecutiveData,
  range: Range,
  marketingData?: MarketingData,
  operationsData?: OperationsData
): AiNarrativeFact[] {
  const facts: AiNarrativeFact[] = executiveToBriefing(execData, range).kpis.map((f) =>
    toNarrativeFact(f, "executive", "Executive")
  );

  if (marketingData) {
    facts.push(
      ...marketingToBriefing(marketingData, range).kpis.map((f) => toNarrativeFact(f, "marketing", "Marketing"))
    );
  }

  if (operationsData) {
    facts.push(
      ...operationsToBriefing(operationsData, range).kpis.map((f) => toNarrativeFact(f, "operations", "Operations"))
    );
  }

  return facts;
}

// Severity -> Status mapping for the rule-based anomaly detectors
// (detectAnomalies/operationsAnomalies — real signals, not AI-derived).
// Used by both the on-screen CrossDeptWatchSummary and the PDF's risks,
// so the same anomaly is never colored differently on the two surfaces.
function anomalySeverityToStatus(severity: AiAnomaly["severity"]): Status {
  return severity === "warning" ? "critical" : "watch";
}

interface DeptWatchGroup {
  department: "marketing" | "operations";
  label: string;
  anomalies: AiAnomaly[];
}

// Executive's cross-department Watch Items summary: reads what each tab's
// existing rule-based anomaly detector has ALREADY computed and re-labels
// it by source department. Executive never recomputes a judgment of its
// own here — same principle as executiveCrossDeptFacts above, just for
// risks instead of plain facts. A department with no anomalies is omitted
// entirely, not shown as an empty placeholder.
function buildCrossDeptWatchGroups(
  marketingData: MarketingData | undefined,
  operationsData: OperationsData | undefined
): DeptWatchGroup[] {
  const groups: DeptWatchGroup[] = [];

  if (marketingData) {
    const anomalies = marketingAnomaliesAndContext(marketingData).anomalies;
    if (anomalies.length > 0) {
      groups.push({ department: "marketing", label: "Marketing", anomalies });
    }
  }

  if (operationsData) {
    const anomalies = operationsAnomalies(operationsData);
    if (anomalies.length > 0) {
      groups.push({ department: "operations", label: "Operations", anomalies });
    }
  }

  return groups;
}

// Converts DeptWatchGroup[] into the PDF's BriefingRisk[] — one row per
// anomaly, labeled with its source department so the PDF's flat Watch
// Items list still reads as grouped by origin. Same status computation as
// the on-screen CrossDeptWatchSummary (anomalySeverityToStatus) — the two
// surfaces read the same underlying groups, so a signal can't be colored
// differently between them.
function crossDeptWatchGroupsToRisks(groups: DeptWatchGroup[]): BriefingRisk[] {
  const risks: BriefingRisk[] = [];
  for (const g of groups) {
    for (const a of g.anomalies) {
      risks.push({
        label: `${g.label} — ${a.label}`,
        detail: a.detail,
        status: anomalySeverityToStatus(a.severity),
      });
    }
  }
  return risks;
}

async function generateAiNarrative(
  facts: AiNarrativeFact[],
  anomalies?: AiAnomaly[],
  context?: Record<string, number | string>
): Promise<AiNarrativeResult | null> {
  try {
    const res = await fetch("/api/ai-narrative", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ facts, anomalies: anomalies ?? [], context: context ?? null }),
    });
    const json = await res.json();
    if (json.error) {
      console.error("AI narrative generation failed:", json.error);
      return null;
    }
    return json as AiNarrativeResult;
  } catch (err) {
    console.error("Could not reach the AI narrative endpoint:", err);
    return null;
  }
}


/* ============== Page ============== */

export default function Home() {
  const [tab, setTab] = useState<Tab>("Executive");
  const [range, setRange] = useState<Range>(DEFAULT_RANGE);
  const [showPayload, setShowPayload] = useState(false);
  const [demoLoading, setDemoLoading] = useState(false);
  const [demoError, setDemoError] = useState<string | null>(null);
  const [briefingLoading, setBriefingLoading] = useState(false); 

  const [aiNarratives, setAiNarratives] = useState<{
    executive?: AiNarrativeResult;
    marketing?: AiNarrativeResult;
    operations?: AiNarrativeResult;
  }>({});

  const [customData, setCustomData] = useState<CustomDataState>({
    connected: false,
    filename: null,
    rows: null,
    source: null,
    data: null,
    profile: null,
    specs: null,
    suggesting: false,
    suggestError: null,
  });

  const [customDataError, setCustomDataError] = useState<string | null>(null);
  const [customDataLoading, setCustomDataLoading] = useState(false);

  const [latestData, setLatestData] = useState<{
    executive?: ExecutiveData;
    operations?: OperationsData;
    marketing?: MarketingData;
  }>({});
  const [lastMs, setLastMs] = useState<number | undefined>(undefined);
  const [tabMs, setTabMs] = useState<{
    executive?: number;
    marketing?: number;
    operations?: number;
  }>({});

  const [isDark, setIsDark] = useState(true);

  useEffect(() => {
    const saved = localStorage.getItem('oneboard-theme');
    const dark = saved !== 'light';
    setIsDark(dark);
    document.documentElement.classList.toggle('dark', dark);
  }, []);

  const toggleTheme = () => {
    setIsDark((prev) => {
      const next = !prev;
      document.documentElement.classList.toggle('dark', next);
      localStorage.setItem('oneboard-theme', next ? 'dark' : 'light');
      return next;
    });
  };


  function handleExecutiveData(data: ExecutiveData, ms?: number) {
    setLatestData((prev) => ({ ...prev, executive: data }));
    if (ms !== undefined) setLastMs(ms);
    setTabMs((prev) => ({ ...prev, executive: ms }));
  }

  function handleOperationsData(data: OperationsData, ms?: number) {
    setLatestData((prev) => ({ ...prev, operations: data }));
    if (ms !== undefined) setLastMs(ms);
     setTabMs((prev) => ({ ...prev, operations: ms }));
  }

  function handleMarketingData(data: MarketingData, ms?: number) {
    setLatestData((prev) => ({ ...prev, marketing: data }));
    if (ms !== undefined) setLastMs(ms);
    setTabMs((prev) => ({ ...prev, marketing: ms }));
  }

  function handleExecutiveNarrative(result: AiNarrativeResult) {
    setAiNarratives((prev) => ({ ...prev, executive: result}));
  }
  function handleMarketingNarrative(result: AiNarrativeResult) {
    setAiNarratives((prev) => ({ ...prev, marketing: result}));
  }
  function handleOperationsNarrative(result: AiNarrativeResult) {
    setAiNarratives((prev) => ({ ...prev, operations: result }));
  }


  // Shared by both upload paths (real file / demo file) — one profile ->
  // suggest call, right after parsing, never re-run per render. The
  // profile (column names/types/samples) is all the AI ever sees; the
  // actual chart numbers are computed afterward, in code, from the full
  // `rows` dataset (see aggregationEngine.ts / CustomDataView).
  async function suggestAndStore(
    filename: string,
    source: "upload" | "demo",
    rows: Record<string, unknown>[]
  ) {
    const profile = profileColumns(rows);
    setCustomData((prev) => ({
      ...prev,
      connected: true,
      filename,
      rows: rows.length,
      source,
      data: rows,
      profile,
      suggesting: true,
      specs: null,
      suggestError: null,
    }));
    try {
      const columns = Object.keys(rows[0] ?? {});
      const res = await fetch("/api/kpi-suggest", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ columns, profile, sampleRows: sampleRowsForProfile(rows) }),
      });
      const json = await res.json();
      if (!res.ok || json.error) {
        throw new Error(json.error || `HTTP ${res.status}`);
      }
      if (json.available === false) {
        throw new Error("AI KPI suggestion is unavailable (no API key configured).");
      }
      const specs: ChartSpec[] = json.specs ?? [];
      setCustomData((prev) => ({ ...prev, suggesting: false, specs }));
    } catch (e) {
      const message = e instanceof Error ? e.message : "Could not suggest KPIs for this data.";
      setCustomData((prev) => ({ ...prev, suggesting: false, suggestError: message }));
    }
  }

  async function handleFileSelect(file: File) {
  setCustomDataError(null);
  setCustomDataLoading(true);
  setTab("Custom Data");

  setCustomData((prev) => ({ ...prev, connected: true, filename: file.name, rows: null, source: "upload", data: null }));

  try {
    const isCsv = file.name.toLowerCase().endsWith(".csv");
    let rows: Record<string, unknown>[];

    if (isCsv) {
      const text = await file.text();
      const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
      if (parsed.errors.length > 0) {
        throw new Error(parsed.errors[0].message);
      }
      rows = parsed.data as Record<string, unknown>[];
    } else {
      // .xlsx / .xls
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) {
        throw new Error("This workbook has no sheets.");
      }
      const sheet = workbook.Sheets[firstSheetName];
      rows = XLSX.utils.sheet_to_json(sheet, { defval: null }) as Record<string, unknown>[];
    }

    if (rows.length === 0) {
      throw new Error("No data rows found — check the file has a header row plus at least one data row.");
    }

    await suggestAndStore(file.name, "upload", rows);
  } catch (e) {
    setCustomDataError(e instanceof Error ? e.message : "Could not parse this file.");
    setCustomData((prev) => ({ ...prev, connected: false, filename: null, rows: null, source: null, data: null, profile: null, specs: null }));
  } finally {
    setCustomDataLoading(false);
  }
}

async function handleLoadDemoData() {
    setDemoLoading(true);
    setDemoError(null);
    setTab("Custom Data");
    try {
      const res = await fetch("/demo-baseline.csv");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const text = await res.text();
      const parsed = Papa.parse<Record<string, unknown>>(text, { header: true, skipEmptyLines: true });
      await suggestAndStore("demo-baseline.csv", "demo", parsed.data);
    } catch (e) {
      setDemoError("Could not load the demo dataset.");
    } finally {
      setDemoLoading(false);
    }
  }

  async function handleExportBriefing() {
  setBriefingLoading(true);
  try {
    let briefing: BriefingData | null = null;
    let narrativeResult: AiNarrativeResult | null = null;
    let inputFacts: AiNarrativeFact[] | undefined;
    let inputAnomalies: AiAnomaly[] | undefined;
    let inputContext: Record<string, number | string> | undefined;

    if (tab === "Executive" && latestData.executive) {
      briefing = executiveToBriefing(
        latestData.executive,
        range,
        latestData.marketing,
        latestData.operations
      );
      narrativeResult = aiNarratives.executive ?? null;
      if (!narrativeResult) {
        inputFacts = executiveCrossDeptFacts(
          latestData.executive,
          range,
          latestData.marketing,
          latestData.operations
        );
      }
    } else if (tab === "Marketing" && latestData.marketing) {
      briefing = marketingToBriefing(latestData.marketing, range);
      narrativeResult = aiNarratives.marketing ?? null;
      if (!narrativeResult) {
        inputFacts = briefing.kpis;
        const mktPart = marketingAnomaliesAndContext(latestData.marketing);
        inputAnomalies = mktPart.anomalies;
        inputContext = mktPart.context;
      }
    } else if (tab === "Operations" && latestData.operations) {
      briefing = operationsToBriefing(latestData.operations, range);
      narrativeResult = aiNarratives.operations ?? null;
      if (!narrativeResult) {
        inputFacts = briefing.kpis;
        inputAnomalies = operationsAnomalies(latestData.operations);
      }
    } else if (tab === "Custom Data" && isCustomDataReady(customData)) {
      // Custom Data has no AI Narrative surface on screen (see
      // CustomDataView) and no fixed metric set to seed one from — leave
      // narrativeResult/inputFacts unset so the auto-generate block below
      // is skipped entirely; the plain, rule-based narrative
      // customDataToBriefing already wrote onto `briefing` stands as-is.
      briefing = customDataToBriefing(customData);
    }

    if (!briefing) {
      alert("Data for this tab hasn't finished loading yet — try again in a moment.");
      return;
    }

    // Export Briefing should always produce the complete report — it
    // shouldn't depend on whether "Generate AI Summary" was clicked first
    // on this tab. If nothing's been generated yet, generate it now.
    if (!narrativeResult && inputFacts) {
      const generated = await generateAiNarrative(inputFacts, inputAnomalies, inputContext);
      if (generated) {
        narrativeResult = generated;
        if (tab === "Executive") setAiNarratives((prev) => ({ ...prev, executive: generated }));
        if (tab === "Marketing") setAiNarratives((prev) => ({ ...prev, marketing: generated }));
        if (tab === "Operations") setAiNarratives((prev) => ({ ...prev, operations: generated }));
      }
    }

    if (narrativeResult) {
      briefing.narrative = narrativeResult.overview;
      briefing.keyObservations = narrativeResult.keyObservations;
      // Executive already has its own risks — the cross-department Watch
      // Items summary set by executiveToBriefing above (real, rule-based,
      // correctly labeled by source department). alertsToRisks matches the
      // AI's alerts against THIS tab's own kpis, which for Executive would
      // silently replace that with a weaker match. Every other tab still
      // gets its risks from alertsToRisks as before.
      if (tab !== "Executive") {
        briefing.risks = alertsToRisks(narrativeResult.alerts, briefing.kpis);
      }
      briefing.recommendations = narrativeResult.recommendations;
    }

    await exportBriefingPDF(briefing);
  } catch (e) {
    alert("Could not generate the briefing PDF.");
  } finally {
    setBriefingLoading(false);
  }
}


  const payload = {
    metadata: {
      engine: "OneBoard",
      activeTab: tab,
      selectedRange: range,
      lastQueryMs: lastMs,
      timestamp: new Date().toISOString(),
    },
    liveTelemetry: latestData,
    customDataSnapshot: customData,
  };

  return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", fontFamily: FONT_STACK }}>
      <style>{`
        @keyframes ob-fade-in {
          from { opacity: 0; transform: translateY(4px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes ob-shimmer {
          from { background-position: -200px 0; }
          to { background-position: 200px 0; }
        }
        .ob-fade-in { animation: ob-fade-in 260ms ease-out; }
        /* Inactive tab panels stay mounted (their data-fetching hooks keep
           running so cross-department data like the Executive briefing's
           marketing/operations facts are available without visiting those
           tabs first) but must not be display:none — a display:none
           box is 0x0, and Recharts' ResponsiveContainer measures that via
           ResizeObserver, logging "width(0) and height(0)" for every chart
           in every inactive tab on every layout pass. visibility:hidden
           keeps the box laid out (so charts measure their real size) while
           position:absolute removes it from flow so it doesn't reserve
           empty space below the active tab.  */
        .ob-tab-inactive {
          position: absolute;
          inset: 0;
          overflow: hidden;
          visibility: hidden;
          pointer-events: none;
          z-index: -1;
        }
        .ob-skeleton {
          background: linear-gradient(90deg, #EEF2F3 25%, #F6F8F9 37%, #EEF2F3 63%);
          background-size: 400px 100%;
          animation: ob-shimmer 1.4s ease-in-out infinite;
        }
        .ob-excel-card {
          transition: border-color 160ms ease, transform 160ms ease, box-shadow 160ms ease;
        }
        .ob-excel-card:hover {
          border-color: ${COLORS.accent} !important;
          box-shadow: ${SHADOW_HOVER};
        }
        .ob-excel-cta {
          transition: color 160ms ease;
        }
        .ob-excel-card:hover .ob-excel-cta {
          color: ${COLORS.accent} !important;
        }
        .ob-ga4-card {
          transition: border-color 160ms ease, transform 160ms ease, box-shadow 160ms ease;
        }
        .ob-ga4-card:hover {
          border-color: ${COLORS.accent} !important;
          box-shadow: ${SHADOW_HOVER};
        }
        .ob-header-btn {
          transition: border-color 160ms ease, transform 160ms ease, box-shadow 160ms ease;
        }
        .ob-header-btn:hover {
          transform: translateY(-1px);
          box-shadow: ${SHADOW_HOVER};
        }
        .ob-header-btn:not(.ob-header-btn-outline):hover {
          border-color: ${COLORS.accent} !important;
        }
        .ob-header-btn-outline:hover {
          border-color: ${COLORS.accent} !important;
          color: ${COLORS.accent} !important;
        }
        @keyframes ob-geo-ping {
          0% { transform: scale(1); opacity: 0.45; }
          75%, 100% { transform: scale(2.6); opacity: 0; }
        }
        .ob-geo-ping {
          transform-origin: center;
          transform-box: fill-box;
          animation: ob-geo-ping 2.4s cubic-bezier(0, 0, 0.2, 1) infinite;
        }
        @keyframes ob-live-dot {
          0%, 100% { opacity: 1; }
          50% { opacity: 0; }
        }
        .ob-live-dot {
          animation: ob-live-dot 1.6s ease-in-out infinite;
        }
        .ob-map-pill-bg {
          fill: #ffffff;
        }
        .dark .ob-map-pill-bg {
          fill: rgba(15, 23, 42, 0.92);
        }
        .ob-map-pill-text {
          fill: #0F172A;
        }
        .dark .ob-map-pill-text {
          fill: var(--pill-color, #10B981);
        }
        @media (prefers-reduced-motion: reduce) {
          .ob-fade-in, .ob-skeleton, .animate-pulse, .ob-geo-ping, .ob-live-dot { animation: none !important; }
        }
      `}</style>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:p-10 space-y-6">
        <header className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-4" style={{ borderBottom: `1px solid ${COLORS.line}` }}>
          <div className="flex items-center gap-3">
            <div
              className="w-11 h-11 rounded-2xl flex items-center justify-center font-black text-xl text-white"
              style={{ background: `linear-gradient(to bottom right, ${COLORS.accent}, ${COLORS.teal})`, boxShadow: "0 2px 8px rgba(5,150,105,0.3)" }}
            >
              O
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-3xl font-black tracking-tight" style={{ color: COLORS.ink }}>
                  OneBoard
                </h1>
                <span
                  className="text-[11px] font-bold px-2.5 py-0.5 rounded-full flex items-center gap-1"
                  style={{ background: COLORS.accentSoft, color: COLORS.accent }}
                >
                  <Sparkles size={11} /> Auto-Dashboard Engine
                </span>
              </div>
              <p className="text-xs font-medium mt-0.5" style={{ color: COLORS.inkSoft }}>
                Live GA4 Web Telemetry API + AI-Generated KPIs from Any File
              </p>
            </div>
          </div>

          <button
            onClick={toggleTheme}
            style={{
              background: COLORS.accentSoft,
              color: COLORS.accent,
              border: "1px solid transparent",
            }}
            className="ob-header-btn px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm transition-colors cursor-pointer"
          >
            {isDark ? <Sun size={13} /> : <Moon size={13} />}
            {isDark ? 'Light' : 'Dark'}
          </button>

          <div className="flex flex-wrap items-center gap-2.5">
            <div

              className="flex items-center gap-2 text-xs font-medium px-3.5 py-2 rounded-xl"
              style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW, fontFamily: MONO_STACK }}
              title="Round-trip time for the most recently loaded GA4 API response"
            >
              <span className="w-1.5 h-1.5 rounded-full ob-live-dot" style={{ background: COLORS.up }} />
              <Zap size={13} style={{ color: "#F59E0B" }} />
              <span style={{ color: COLORS.up, fontWeight: 700 }}>{lastMs !== undefined ? `${lastMs}ms` : "…"}</span>
            </div>

            <button
              onClick={handleLoadDemoData}
              disabled={demoLoading}
              className="ob-header-btn text-xs font-bold px-3.5 py-2 rounded-xl flex items-center gap-2 transition active:scale-95 disabled:opacity-60"
              style={{ background: COLORS.accentSoft, color: COLORS.accent, border: "1px solid transparent" }}
            >
              <FileSpreadsheet size={14} />
              {demoLoading ? "Loading…" : "Load Sample Data"}
            </button>

            {EXPORT_REPORT_ENABLED && (
              <button
              onClick={handleExportBriefing}
              disabled={briefingLoading}
              className="ob-header-btn text-xs font-bold px-3.5 py-2 rounded-xl flex items-center gap-2 transition active:scale-95 disabled:opacity-60"
              style={{ background: COLORS.indigoSoft, color: COLORS.indigo, border: "1px solid transparent" }}
              >
              {briefingLoading ? <Loader2 size={14} className="animate-spin" /> : <FileText size={14} />}
              {briefingLoading ? "Generating…" : "Export Report"}
              </button>
            )}

            <button
              onClick={() => setShowPayload(true)}
              className="ob-header-btn ob-header-btn-outline text-xs font-bold px-3.5 py-2 rounded-xl flex items-center gap-2 transition active:scale-95"
              style={{ background: COLORS.surface, color: COLORS.inkSoft, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW }}
            >
              <Code size={14} />
              Raw UDM Payload
            </button>
          </div>
        </header>

        {demoError && <ErrorBox message={demoError} />}
        {customDataError && <ErrorBox message={customDataError} />}

        <SourcePanel
          activeEndpoint={
            tab === "Custom Data"
              ? "/api/kpi-suggest (client-computed)"
              : `/api/ga4/${tab.toLowerCase()} (${
                  (tab === "Executive"
                    ? tabMs.executive
                    : tab === "Marketing"
                    ? tabMs.marketing
                    : tabMs.operations) ?? "…"
                }ms)`
          }
          customData={customData}
          onFileSelect={handleFileSelect}
          customDataActive={tab === "Custom Data"}
        />

        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="overflow-x-auto max-w-full">
            <SegmentedControl
              options={TABS.map((t) => ({
                value: t,
                label: (
                  <>
                    <span className="sm:hidden">{t}</span>
                    <span className="hidden sm:inline">{TAB_LABELS[t]}</span>
                  </>
                ),
              }))}
              value={tab}
              onChange={setTab}
              variant="accent"
            />
          </div>
          {tab !== "Custom Data" ? (
            <SegmentedControl options={RANGES} value={range} onChange={setRange} size="sm" variant="accent" />
          ) : (
            <span className="text-xs font-medium" style={{ color: COLORS.inkFaint }}>
              All uploaded rows
            </span>
          )}
        </div>

        <div style={{ position: "relative" }}>
          <div className={tab === "Executive" ? "ob-fade-in" : "ob-tab-inactive"}>
            <ExecutiveView
              range={range}
              onData={handleExecutiveData}
              onNarrative={handleExecutiveNarrative}
              narrative={aiNarratives.executive}
              marketingData={latestData.marketing}
              operationsData={latestData.operations}
            />
          </div>
          <div className={tab === "Marketing" ? "ob-fade-in" : "ob-tab-inactive"}>
            <MarketingView range={range} onData={handleMarketingData} onNarrative={handleMarketingNarrative} narrative={aiNarratives.marketing} />
          </div>
          <div className={tab === "Operations" ? "ob-fade-in" : "ob-tab-inactive"}>
            <OperationsView range={range} onData={handleOperationsData} onNarrative={handleOperationsNarrative} narrative={aiNarratives.operations} />
          </div>
          <div className={tab === "Custom Data" ? "ob-fade-in" : "ob-tab-inactive"}>
            <CustomDataView state={customData} onLoadDemo={handleLoadDemoData} demoLoading={demoLoading} />
          </div>
        </div>

        <div
          className="mt-6 pt-4 flex flex-wrap items-center justify-between gap-2 text-[8px]"
          style={{ borderTop: `1px solid ${COLORS.line}`, color: COLORS.inkFaint }}
        >
          <span>OneBoard · Built with Next.js, TypeScript &amp; the GA4 Data API</span>
          <span>Created on Aug 2026</span>
        </div>
      </main>

      {showPayload && <RawPayloadModal payload={payload} onClose={() => setShowPayload(false)} />}
    </div>
  );
}
