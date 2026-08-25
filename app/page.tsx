"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode, ComponentType } from "react";
import { detectAnomalies, detectSustainedTrend } from "./api/ai-narrative/_lib/anomalies";
import type { AiNarrativeFact, AiAnomaly, AiNarrativeResult } from "../types/ai-narrative";
import type { BriefingData, BriefingFact, BriefingRisk, DrilldownTable, Status } from "../components/ExecutiveBriefingPDF";
import type { ExcelDepartment, ExcelMetricKey, ExcelRowClassification } from "../types/excel-blend";
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
  UserMinus,
  GraduationCap,
  UserPlus,
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

type HRFact = {
  id: string;
  metric: string;
  value: number | string;
  unit?: string;
  delta?: { value: number; direction: "up" | "down"; period: string };
  context?: string;
};

type HRData = {
  source: "sample";
  meta: { label: string; generatedAt: string; note: string };
  facts: HRFact[];
  monthlySnapshot: {
    month: string;
    headcount: number;
    hires: number;
    terminations: number;
    attritionRate: number;
    complianceRate: number;
  }[];
  companyTargets: { attritionTargetMax: number; complianceTargetMin: number };
  queryMs?: number;
};


const TABS = ["Executive", "Marketing", "Operations", "HR"] as const;
type Tab = (typeof TABS)[number];
const TAB_LABELS: Record<Tab, string> = {
  Executive: "Executive View",
  Marketing: "Marketing & Funnel",
  Operations: "Operations & Diagnostics",
  HR: "HR View"
};
  
const RANGES = [
  { value: "7d", label: "7 Days" },
  { value: "30d", label: "30 Days" },
  { value: "90d", label: "90 Days" },
] as const;
type Range = (typeof RANGES)[number]["value"];

const DEFAULT_RANGE: Range = "30d";

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
        : Math.round(animated).toLocaleString();

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
        <p className="text-[11px] font-bold uppercase tracking-wider truncate" style={{ color: COLORS.inkSoft }}>
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
            fontSize: displayValue !== undefined ? 22 : 30,
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
}: {
  title: string;
  icon?: IconType;
  rows: { label: string; value: number }[];
  barColor?: string;
  emptyLabel?: string;
  className?: string;
}) {
  const chartData = rows.map((r) => ({
    label: r.label.length > 26 ? r.label.slice(0, 24) + "…" : r.label,
    value: r.value,
  }));
  const height = Math.max(120, chartData.length * 38);

  return (
    <Card title={title} icon={icon} className={className}>
      {rows.length === 0 ? (
        <p className="text-sm" style={{ color: COLORS.inkFaint }}>
          {emptyLabel}
        </p>
      ) : (
        <ResponsiveContainer width="100%" height={height}>
          <BarChart data={chartData} layout="vertical" margin={{ left: 8, right: 36, top: 4, bottom: 4 }}>
            <XAxis type="number" hide />
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

type ExcelState = {
  connected: boolean;
  filename: string | null;
  rows: number | null;
  source: "upload" | "demo" | null;
  data: Record<string, unknown>[] | null;
  // Classification runs ONCE, right after parsing, via /api/excel-classify
  // — never per-report. `classified` is persisted (see EXCEL_BLEND_STORAGE_KEY)
  // so a page reload reuses it instead of calling the AI again.
  classified: ExcelRowClassification[] | null;
  classifying: boolean;
  classifyError: string | null;
};

const EXCEL_BLEND_STORAGE_KEY = "oneboard-excel-blend-v1";

// One shared selector every tab uses to read only the rows AI tagged as
// its own department — no tab ever reads another department's rows.
function selectDeptRows(excel: ExcelState, department: ExcelDepartment): Record<string, unknown>[] {
  if (!excel.data || !excel.classified) return [];
  const idx = new Set(
    excel.classified.filter((c) => c.department === department).map((c) => c.rowIndex)
  );
  return excel.data.filter((_, i) => idx.has(i));
}

// Finds the first candidate column present in a row (case/whitespace
// insensitive) — same "don't assume exact header spelling" approach
// already used by matchExcelChannelTargets, shared here so the
// metric/target_value schema (Operations/HR/Executive) uses one lookup,
// not three copies.
function findColumn(sampleKeys: string[], candidates: string[]): string | undefined {
  return sampleKeys.find((k) => candidates.includes(k.trim().toLowerCase()));
}

const TARGET_VALUE_CANDIDATES = ["target_value", "target value", "target", "goal", "value"];

// Looks up a single metricKey's target value among a department's
// classified rows (used by HR's override and by matchExcelMetricTargets
// below). Returns null — never a guessed/default number — if no row with
// that metricKey exists or its target_value column can't be parsed.
function findMetricTargetFromExcel(excel: ExcelState, metricKey: ExcelMetricKey): number | null {
  if (!excel.data || !excel.classified) return null;
  const match = excel.classified.find((c) => c.metricKey === metricKey);
  if (!match) return null;
  const row = excel.data[match.rowIndex];
  if (!row) return null;
  const key = findColumn(Object.keys(row), TARGET_VALUE_CANDIDATES);
  if (!key) return null;
  const value = Number(row[key]);
  return Number.isFinite(value) ? value : null;
}

function summarizeClassification(classified: ExcelRowClassification[]): string {
  const labels: Record<string, string> = {
    marketing: "Marketing",
    operations: "Operations",
    hr: "HR",
    executive: "Executive",
  };
  const counts = new Map<string, number>();
  for (const c of classified) {
    const key = c.department ?? "unclassified";
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  const order = ["marketing", "operations", "hr", "executive", "unclassified"];
  return order
    .filter((k) => counts.has(k))
    .map((k) => `${counts.get(k)} ${labels[k] ?? "unclassified"}`)
    .join(" · ");
}

function SourcePanel({
  activeEndpoint,
  excel,
  onExcelFileSelect,
  hrActive,
  activeDepartment,
}: {
  activeEndpoint: string;
  excel: ExcelState;
  onExcelFileSelect: (file: File) => void;
  hrActive: boolean;
  activeDepartment: ExcelDepartment;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  // Border highlight only indicates which source the CURRENT tab is
  // reading from. It does not imply the other source is broken or
  // unavailable — GA4 stays fully "live" in its own copy for the other
  // three tabs regardless of which tab is active right now.
  //
  // Which tab "reads from" the blended file is now decided by the AI
  // classifier, not hardcoded to one tab — highlight whenever the
  // currently active tab actually has classified rows, or while
  // classification is still in flight for a just-uploaded file.
  const activeDeptHasRows = (excel.classified ?? []).some((c) => c.department === activeDepartment);
  const excelBorderActive = excel.classifying || activeDeptHasRows;

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
      {/* GA4 — content never changes with tab; only the border highlight
          toggles off when this tab isn't reading from it. */}
      <div
        className="ob-ga4-card rounded-2xl p-4 hover:-translate-y-0.5"
        style={{
          background: `linear-gradient(to bottom, ${COLORS.surface}, ${COLORS.accentSoft})`,
          border: hrActive ? `1px solid ${COLORS.line}` : `1px solid var(--color-accent-border)`,
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

      {/* Excel / CSV — border highlights on HR tab (same data-source
          category) or when the person has actually uploaded a file.
          Text content reflects the REAL upload state only — it never
          claims a sample dataset is "connected" via upload. */}
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
            if (file) onExcelFileSelect(file);
          }}
        />
        <div className="flex items-center justify-between mb-2">
          <span className="font-bold text-xs uppercase tracking-wider" style={{ color: COLORS.ink }}>
            Excel / CSV Import
          </span>
          <span
            className="px-2.5 py-0.5 text-[10px] font-bold rounded-full"
            style={{
              background: excel.connected ? COLORS.accentSoft : COLORS.track,
              color: excel.connected ? COLORS.accent : COLORS.inkSoft,
            }}
          >
            {excel.connected ? (excel.source === "demo" ? "Demo Loaded" : "File Selected") : "Import Mode"}
          </span>
        </div>
        <p className="text-xs font-medium truncate" style={{ color: COLORS.inkSoft }}>
          {excel.filename ?? "Click to blend offline baseline target spreadsheet"}
        </p>
        <div
          className="ob-excel-cta mt-3 text-[11px] font-semibold flex items-center gap-1.5"
          style={{ color: excelBorderActive ? COLORS.accent : COLORS.inkFaint }}
        >
          <Upload size={13} />
          <span>
            {excel.connected
              ? excel.classifying
                ? `${excel.rows ?? "…"} rows parsed — classifying with AI…`
                : excel.classified
                ? `${excel.rows ?? excel.classified.length} rows parsed — ${summarizeClassification(excel.classified)}`
                : excel.classifyError
                ? "Classification failed — see error above"
                : excel.rows !== null
                ? `${excel.rows} rows parsed from ${excel.filename ?? "file"} — ready to blend`
                : "File selected — ready to parse"
              : "+ Upload Offline Dataset (.xlsx/.csv)"}
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
              {g.targetRatio && (
                <span
                  className="text-xs font-semibold"
                  style={{ color: g.targetRatio.onTarget === g.targetRatio.total ? COLORS.up : COLORS.amberInk }}
                >
                  {g.targetRatio.onTarget}/{g.targetRatio.total} metrics on target
                </span>
              )}
            </div>
            {g.targetRatio && g.targetRatio.misses.length > 0 && (
              <p className="text-xs mb-2" style={{ color: COLORS.inkFaint }}>
                Off target: {g.targetRatio.misses.join(", ")}
              </p>
            )}
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
  hrData,
  excel,
}: {
  range: Range;
  onData?: (data: ExecutiveData, clientMs?: number) => void;
  onNarrative?: (result: AiNarrativeResult) => void;
  narrative?: AiNarrativeResult;
  marketingData?: MarketingData;
  operationsData?: OperationsData;
  hrData?: HRData;
  excel?: ExcelState;
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
  const executiveFacts = executiveCrossDeptFacts(data, range, marketingData, operationsData, hrData);
  const { anomalies: executiveAnomalies, context: executiveContext } = hrData
    ? hrAnomaliesAndContext(hrData)
    : { anomalies: [] as AiAnomaly[], context: undefined as Record<string, number | string> | undefined };

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

      {(() => {
        if (!excel?.classified) return null;
        const deptRows = selectDeptRows(excel, "executive");
        if (deptRows.length === 0) return null;
        const matches = matchExcelMetricTargets(excel, [
          { metricKey: "activeUsersTarget", label: "Active Users", actual: data.summary.activeUsers },
          { metricKey: "sessionsTarget", label: "Sessions", actual: data.summary.sessions },
          {
            metricKey: "engagementRateTarget",
            label: "Engagement Rate",
            actual: Math.round(data.summary.engagementRate * 100),
          },
        ]);
        if (!matches) return null;
        return (
          <div className="mt-4">
            <TargetBlendCard
              title="Executive Metrics vs. Excel Baseline Target"
              rows={matches.map((m) => ({
                ...m,
                display: m.label === "Engagement Rate" ? (n: number) => `${n}%` : undefined,
              }))}
              totalRows={deptRows.length}
              unmatchedNote="metric not tracked on this dashboard"
            />
          </div>
        );
      })()}

      <div className="mt-4">
        <CrossDeptWatchSummary
          groups={buildCrossDeptWatchGroups(marketingData, operationsData, hrData, excel)}
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

// Matches parsed Excel/CSV rows against the real, already-fetched GA4
// channel sessions. This is the actual "blend" step — it was previously
// missing entirely, which is why uploading or loading a baseline file had
// zero visible effect anywhere on the dashboard (excel.data was parsed and
// stored in state, but no view ever read it).
//
// Column matching is intentionally loose on NAME (a few common header
// spellings) but strict on VALUE: a row only produces a comparison if its
// channel text exact-matches (case-insensitive) a channel GA4 actually
// reported this period. Unmatched rows are surfaced as a count, never
// silently dropped or guessed at — same "don't invent it" rule as the
// PDF briefings.
function matchExcelChannelTargets(
  rows: Record<string, unknown>[] | null,
  channels: { channel: string; sessions: number }[]
): { channel: string; actual: number; target: number; onTrack: boolean }[] | null {
  if (!rows || rows.length === 0 || channels.length === 0) return null;

  const channelKeyCandidates = ["channel", "channel group", "source", "medium"];
  const targetKeyCandidates = ["target_sessions", "target sessions", "target", "goal", "goal_sessions"];

  const sampleKeys = Object.keys(rows[0] ?? {});
  const channelKey = sampleKeys.find((k) => channelKeyCandidates.includes(k.trim().toLowerCase()));
  const targetKey = sampleKeys.find((k) => targetKeyCandidates.includes(k.trim().toLowerCase()));
  if (!channelKey || !targetKey) return null;

  const byChannel = new Map(channels.map((c) => [c.channel.trim().toLowerCase(), c.sessions]));

  // onTrack computed once here — reused by TargetBlendCard's bars AND the
  // cross-department Watch Items summary, so the two can never disagree
  // about whether a channel hit its target.
  const matched: { channel: string; actual: number; target: number; onTrack: boolean }[] = [];
  for (const row of rows) {
    const name = String(row[channelKey] ?? "").trim();
    const target = Number(row[targetKey]);
    if (!name || !Number.isFinite(target)) continue;
    const actual = byChannel.get(name.toLowerCase());
    if (actual === undefined) continue; // real rule: no match, no row — never guess
    matched.push({ channel: name, actual, target, onTrack: actual >= target });
  }
  return matched.length > 0 ? matched : null;
}

// Generic actual-vs-target bar comparison — used by Marketing (rows keyed
// by channel name), Operations, and Executive (rows keyed by metric
// label). Same rendering regardless of caller; only the title/footnote
// text differ.
function TargetBlendCard({
  title,
  rows,
  totalRows,
  unmatchedNote,
}: {
  title: string;
  // onTrack is computed by the matcher (matchExcelChannelTargets /
  // matchExcelMetricTargets), not here — this card and the
  // cross-department Watch Items summary both read the same value instead
  // of each re-deriving actual-vs-target (and the lowerIsBetter direction)
  // independently.
  rows: { label: string; actual: number; target: number; onTrack: boolean; display?: (n: number) => string }[];
  totalRows: number;
  unmatchedNote: string;
}) {
  const unmatched = totalRows - rows.length;
  return (
    <Card title={title} icon={FileSpreadsheet}>
      <div className="space-y-3">
        {rows.map((r) => {
          const max = Math.max(r.actual, r.target, 1);
          const onTrack = r.onTrack;
          const statusColor = onTrack ? COLORS.up : COLORS.amberInk;
          const fmt = r.display ?? ((n: number) => String(n));
          return (
            <div key={r.label}>
              <div className="flex items-center justify-between text-xs font-semibold mb-1">
                <span style={{ color: COLORS.ink }}>{r.label}</span>
                <span style={{ color: statusColor }}>
                  {fmt(r.actual)} / {fmt(r.target)} target {onTrack ? "▲" : "▼"}
                </span>
              </div>
              <div className="flex gap-1">
                <div className="h-2 rounded-full flex-1" style={{ background: COLORS.track }}>
                  <div
                    className="h-2 rounded-full"
                    style={{ width: `${Math.min(100, (r.actual / max) * 100)}%`, background: statusColor }}
                  />
                </div>
              </div>
            </div>
          );
        })}
      </div>
      <p className="text-[11px] mt-3" style={{ color: COLORS.inkFaint }}>
        {rows.length} of {totalRows} baseline row{totalRows === 1 ? "" : "s"} matched{unmatched > 0 ? ` — ${unmatched} unmatched (${unmatchedNote})` : "."}
      </p>
    </Card>
  );
}

// Operations/Executive share the "metric,target_value" schema — this one
// matcher looks up each of a caller's known fields by metricKey among a
// department's classified rows. Never invents a comparison row: a field
// with no matching classified row (or an unparseable target_value) is
// simply omitted, same "no match, no row" rule as the channel matcher.
function matchExcelMetricTargets(
  excel: ExcelState,
  fields: { metricKey: ExcelMetricKey; label: string; actual: number; lowerIsBetter?: boolean }[]
): { label: string; actual: number; target: number; onTrack: boolean; lowerIsBetter?: boolean }[] | null {
  if (!excel.classified) return null;
  const rows: { label: string; actual: number; target: number; onTrack: boolean; lowerIsBetter?: boolean }[] = [];
  for (const f of fields) {
    const target = findMetricTargetFromExcel(excel, f.metricKey);
    if (target === null) continue;
    const onTrack = f.lowerIsBetter ? f.actual <= target : f.actual >= target;
    rows.push({ label: f.label, actual: f.actual, target, onTrack, lowerIsBetter: f.lowerIsBetter });
  }
  return rows.length > 0 ? rows : null;
}

function MarketingView({
  range,
  onData,
  onNarrative,
  narrative,
  excel,
}: {
  range: Range;
  onData?: (data: MarketingData, clientMs?: number) => void;
  onNarrative?: (result: AiNarrativeResult) => void;
  narrative?: AiNarrativeResult;
  excel?: ExcelState;
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

      {(() => {
        if (!excel?.classified) return null;
        const deptRows = selectDeptRows(excel, "marketing");
        if (deptRows.length === 0) return null;
        const matches = matchExcelChannelTargets(deptRows, data.channels ?? []);
        if (!matches) return null;
        return (
          <div className="mt-4">
            <TargetBlendCard
              title="Channel Sessions vs. Excel Baseline Target"
              rows={matches.map((m) => ({ label: m.channel, actual: m.actual, target: m.target, onTrack: m.onTrack }))}
              totalRows={deptRows.length}
              unmatchedNote="no live sessions for that channel name"
            />
          </div>
        );
      })()}

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
  excel,
}: {
  range: Range;
  onData?: (data: OperationsData, clientMs?: number) => void;
  onNarrative?: (result: AiNarrativeResult) => void;
  narrative?: AiNarrativeResult;
  excel?: ExcelState;
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

      {(() => {
        if (!excel?.classified) return null;
        const deptRows = selectDeptRows(excel, "operations");
        if (deptRows.length === 0) return null;
        const matches = matchExcelMetricTargets(excel, [
          {
            metricKey: "bounceRateTarget",
            label: "Bounce Rate",
            actual: Math.round(data.summary.bounceRate * 100),
            lowerIsBetter: true,
          },
          {
            metricKey: "avgEngagementDurationTargetSeconds",
            label: "Avg. Engagement Duration",
            actual: data.summary.avgSessionDuration,
          },
        ]);
        if (!matches) return null;
        return (
          <div className="mt-4">
            <TargetBlendCard
              title="Operations Metrics vs. Excel Baseline Target"
              rows={matches.map((m) => ({
                ...m,
                display: m.label === "Avg. Engagement Duration" ? formatDuration : (n: number) => `${n}%`,
              }))}
              totalRows={deptRows.length}
              unmatchedNote="metric not tracked on this dashboard"
            />
          </div>
        );
      })()}

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

function monthLabel(m: string) {
  const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const mo = Number(m.split("-")[1]);
  return names[mo - 1] ?? m;
}

function hrFactVisual(id: string): { icon: IconType; color: string; bg: string } {
  if (id === "headcount") return { icon: Users, color: COLORS.accent, bg: COLORS.accentSoft };
  if (id === "attrition") return { icon: UserMinus, color: COLORS.down, bg: COLORS.downSoft };
  if (id === "compliance") return { icon: GraduationCap, color: COLORS.blue, bg: COLORS.blueSoft };
  return { icon: UserPlus, color: COLORS.indigo, bg: COLORS.indigoSoft };
}

function HrFactCard({ fact }: { fact: HRFact }) {
  const { icon: Icon, color, bg } = hrFactVisual(fact.id);
  // attrition/compliance deltas are stored as percentage-point (pp)
  // differences, not relative percent change — see hrFacts.ts for why.
  const isRateMetric = fact.id === "attrition" || fact.id === "compliance";

  // "up" isn't always good: rising attrition is bad, rising compliance is
  // good. Flag metrics where a rise should read as a warning, not progress —
  // same inversion concept as the DeltaBadge component's `invert` prop used
  // elsewhere in this file (e.g. bounce rate).
  const invertGoodDirection = fact.id === "attrition";
  const isFlat = fact.delta ? fact.delta.value === 0 : false;
  const isUp = fact.delta?.direction === "up";
  const isGood = isFlat ? null : invertGoodDirection ? !isUp : isUp;
  const deltaColor = isFlat ? COLORS.inkFaint : isGood ? COLORS.up : COLORS.teal;
  const deltaArrow = isFlat ? "" : isUp ? "▲" : "▼";
  const deltaLabel = fact.delta
    ? `${deltaArrow}${deltaArrow ? " " : ""}${fact.delta.value}${isRateMetric ? "pp" : "%"} MoM`
    : null;

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
        <p className="text-[11px] font-bold uppercase tracking-wider truncate" style={{ color: COLORS.inkSoft }}>
          {fact.metric}
        </p>
        <div className="w-6 h-6 rounded-full flex items-center justify-center shrink-0" style={{ background: bg }}>
          <Icon size={13} strokeWidth={2.25} color={color} />
        </div>
      </div>

      <p
        className="font-extrabold leading-none truncate"
        style={{ color: COLORS.ink, fontVariantNumeric: "tabular-nums", letterSpacing: "-0.02em", fontSize: 30 }}
      >
        {fact.value}
        {fact.unit ?? ""}
      </p>

      <div className="mt-2 h-5 flex items-center">
        {deltaLabel ? (
          <span className="text-xs font-semibold" style={{ color: deltaColor }}>
            {deltaLabel}
          </span>
        ) : fact.context ? (
          <span className="text-xs truncate font-medium" style={{ color: COLORS.inkFaint }}>
            {fact.context}
          </span>
        ) : null}
      </div>
      {deltaLabel && fact.context && (
        <p className="text-[11px] mt-1 truncate" style={{ color: COLORS.inkFaint }}>
          {fact.context}
        </p>
      )}
    </div>
  );
}

function HrTrendChart({ snapshot }: { snapshot: HRData["monthlySnapshot"] }) {
  const chartData = snapshot.map((m) => ({
    label: monthLabel(m.month),
    attritionRate: m.attritionRate,
    complianceRate: m.complianceRate,
  }));
  return (
    <Card title="Attrition & Compliance Trend · Trailing 12 Months" icon={TrendingUp}>
      <ResponsiveContainer width="100%" height={220}>
        <LineChart data={chartData} margin={{ top: 4, right: 8, bottom: 0, left: 4 }}>
          <CartesianGrid stroke={COLORS.line} vertical={false} strokeDasharray="3 3" />
          <XAxis dataKey="label" fontSize={11} stroke={COLORS.inkFaint} tickLine={false} axisLine={false} />
          <YAxis
            yAxisId="attrition"
            fontSize={11}
            stroke={COLORS.teal}
            tickLine={false}
            axisLine={false}
            width={44}
            tickFormatter={(v) => `${v}%`}
          />
          <YAxis
            yAxisId="compliance"
            orientation="right"
            fontSize={11}
            stroke={COLORS.blue}
            tickLine={false}
            axisLine={false}
            width={50}
            domain={['dataMin - 2', 'dataMax + 2']}
            tickFormatter={(v) => `${v}%`}
          />
          <Tooltip content={<ChartTooltip />} />
          <Line
            yAxisId="attrition"
            type="monotone"
            dataKey="attritionRate"
            name="Attrition %"
            stroke={COLORS.teal}
            strokeWidth={2}
            dot={false}
          />
          <Line
            yAxisId="compliance"
            type="monotone"
            dataKey="complianceRate"
            name="Compliance %"
            stroke={COLORS.blue}
            strokeWidth={2}
            dot={false}
          />
          <Legend
            iconType="circle"
            iconSize={7}
            formatter={(value) => (
              <span className="text-xs font-medium" style={{ color: COLORS.inkSoft }}>
                {value}
              </span>
            )}
          />
        </LineChart>
      </ResponsiveContainer>
    </Card>
  );
}

function HrHeadcountChart({ snapshot }: { snapshot: HRData["monthlySnapshot"] }) {
  const chartData = snapshot.map((m) => ({
    label: monthLabel(m.month),
    hires: m.hires,
    terminations: m.terminations,
  }));
  return (
    <Card title="Hires vs. Terminations · Trailing 12 Months" icon={Users}>
      <ResponsiveContainer width="100%" height={220}>
        <BarChart data={chartData} margin={{ top: 4, right: 4, bottom: 0, left: -16 }}>
          <CartesianGrid stroke={COLORS.line} vertical={false} strokeDasharray="3 3" />
          <XAxis dataKey="label" fontSize={11} stroke={COLORS.inkFaint} tickLine={false} axisLine={false} />
          <YAxis fontSize={11} stroke={COLORS.inkFaint} tickLine={false} axisLine={false} allowDecimals={false} />
          <Tooltip content={<ChartTooltip />} />
          <Bar dataKey="hires" name="New Hires" fill={COLORS.accent} radius={[6, 6, 0, 0]} barSize={10}  />
          <Bar dataKey="terminations" name="Terminations" fill="#94A3B8" radius={[6, 6, 0, 0]} barSize={10} />
          <Legend
            iconType="circle"
            iconSize={7}
            formatter={(value) => (
              <span className="text-xs font-medium" style={{ color: COLORS.inkSoft }}>
                {value}
              </span>
            )}
          />
        </BarChart>
      </ResponsiveContainer>
    </Card>
  );
}

function HRSkeleton() {
  return (
    <div>
      <SkeletonKpiRow count={4} />
      <SkeletonChartGrid count={2} />
    </div>
  );
}

// HR's default targets (companyTargets from /api/hr — 95%/8%, verified
// and reused across this whole app: the PDF one-pager, the AI narrative's
// anomaly thresholds, everything) stay the default. If the AI classified
// any uploaded rows as HR's attritionTargetMax/complianceTargetMin, THOSE
// numbers override the default — never the other way around, and never a
// partial override that leaves the descriptive `context` string on the
// fact quoting the OLD number while the pill/threshold logic uses the
// NEW one. Mirrors the exact context formula from
// app/api/hr/_lib/hrFacts.ts's computeHRFacts so the two can never drift.
function mergeHrOverrides(data: HRData, excel?: ExcelState): HRData {
  if (!excel?.classified) return data;
  const attritionOverride = findMetricTargetFromExcel(excel, "attritionTargetMax");
  const complianceOverride = findMetricTargetFromExcel(excel, "complianceTargetMin");
  if (attritionOverride === null && complianceOverride === null) return data;

  const companyTargets = {
    attritionTargetMax: attritionOverride ?? data.companyTargets.attritionTargetMax,
    complianceTargetMin: complianceOverride ?? data.companyTargets.complianceTargetMin,
  };
  const latest = data.monthlySnapshot[data.monthlySnapshot.length - 1];

  return {
    ...data,
    companyTargets,
    facts: data.facts.map((f) => {
      if (f.id === "attrition" && latest) {
        return {
          ...f,
          context:
            latest.attritionRate <= companyTargets.attritionTargetMax
              ? `Within internal target (<${companyTargets.attritionTargetMax}%)`
              : `Above internal target (<${companyTargets.attritionTargetMax}%)`,
        };
      }
      if (f.id === "compliance" && latest) {
        return {
          ...f,
          context:
            latest.complianceRate >= companyTargets.complianceTargetMin
              ? `Meets internal target (>${companyTargets.complianceTargetMin}%)`
              : `Below internal target (>${companyTargets.complianceTargetMin}%)`,
        };
      }
      return f;
    }),
  };
}

function HRView({
  onData,
  onNarrative,
  narrative,
  excel,
}: {
  onData?: (data: HRData) => void;
  onNarrative?: (result: AiNarrativeResult) => void;
  narrative?: AiNarrativeResult;
  excel?: ExcelState;
}) {

  const { data, error, isFetching } = useGa4<HRData>("/api/hr", "30d");
  const mergedData = useMemo(() => (data ? mergeHrOverrides(data, excel) : null), [data, excel]);

  useEffect(() => {
    if (mergedData) onData?.(mergedData);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mergedData]);

  if (error) return <ErrorBox message={error} />;
  if (!mergedData || !data) return <HRSkeleton />;

  const overrideActive = mergedData.companyTargets !== data.companyTargets;

  return (
    <div>
      <RefreshingNote show={isFetching} />

      {overrideActive && (
        <p className="text-xs mb-3 flex items-center gap-1.5" style={{ color: COLORS.accent }}>
          <FileSpreadsheet size={13} />
          Target overridden by uploaded baseline — attrition ≤ {mergedData.companyTargets.attritionTargetMax}%
          &nbsp;·&nbsp; compliance ≥ {mergedData.companyTargets.complianceTargetMin}%
        </p>
      )}

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        {mergedData.facts.map((f) => (
          <HrFactCard key={f.id} fact={f} />
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
        <div id="ob-chart-hr">
          <HrTrendChart snapshot={mergedData.monthlySnapshot} />
        </div>
        <HrHeadcountChart snapshot={mergedData.monthlySnapshot} />
      </div>

      {(() => {
  const { anomalies, context } = hrAnomaliesAndContext(mergedData);
  return (
    <AiNarrativeCard
      facts={hrToBriefing(mergedData, "30d").kpis}
      anomalies={anomalies}
      context={context}
      result={narrative}
      onGenerated={(result) => onNarrative?.(result)}
    />
  );
})()}

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
  operationsData?: OperationsData,
  hrData?: HRData,
  excel?: ExcelState
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
    risks: crossDeptWatchGroupsToRisks(buildCrossDeptWatchGroups(marketingData, operationsData, hrData, excel)),
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
  // drilldown table rows below, same pattern as classifyHrStatus.
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

// Mirrors HrFactCard's existing `invertGoodDirection` logic — that
// component already correctly knows rising attrition is bad on screen; the
// PDF export must never disagree with what's already shown for the same
// number. Thresholds come from data.companyTargets — real fields already
// used by hrAnomaliesAndContext, nothing invented here.
function classifyHrStatus(
  fact: HRFact,
  targets: HRData["companyTargets"]
): { status: Status; statusLabel: string; deltaLabel: string; note?: string } {
  const value = typeof fact.value === "number" ? fact.value : parseFloat(String(fact.value));

  if (fact.id === "attrition") {
    const max = targets.attritionTargetMax;
    const critical = value >= max;
    const rising = fact.delta?.direction === "up" && (fact.delta?.value ?? 0) > 0;
    const status: Status = critical ? "critical" : rising ? "watch" : "ok";
    const statusLabel = critical ? "Above Target" : rising ? "Rising · Watch" : "On Track";
    const deltaLabel = fact.delta ? `${fact.delta.value}pp (< ${max}% Target)` : `< ${max}% Target`;
    return { status, statusLabel, deltaLabel };
  }

  if (fact.id === "compliance") {
    const min = targets.complianceTargetMin;
    const below = value < min;
    const status: Status = below ? "critical" : "ok";
    const statusLabel = below ? "Below Target" : "On Track";
    // Was `${min}.0%` — hardcoded a single trailing decimal on the
    // assumption `min` is always a whole number (true for the 95 that
    // ships as the default). An Excel-baseline override can be any
    // decimal (e.g. 99.95), which turned this into "99.95.0%". Just
    // interpolate the real number, same as the attrition branch above.
    const deltaLabel = below ? `Below ${min}% Target` : `Above ${min}% Target`;
    return { status, statusLabel, deltaLabel };
  }

  if (fact.id === "headcount") {
    // No real target exists for headcount. A small MoM move is noise, not
    // a status signal — stays neutral regardless of direction, matching
    // the "STABLE" pill already shown on screen.
    // fact.delta.value is already signed here (hrFacts.ts's toDeltaField,
    // used only for headcount, keeps the raw +/- pct — unlike
    // toPointDeltaField for attrition/compliance, which Math.abs()'s it).
    // Prepending another "+"/"-" on top double-signs a decrease into
    // "--0.3%". The ArrowTriangle already shown next to this label carries
    // the direction, same as HrFactCard's ▲/▼ glyph on screen — so just use
    // the raw value, matching that pattern instead of re-deriving a sign.
    const deltaLabel = fact.delta ? `${fact.delta.value}% MoM` : "—";
    return { status: "neutral", statusLabel: "Stable", deltaLabel };
  }

  // Remaining fact (top hire source) — informational, no target to breach.
  // Same pattern as Executive's "Top Channel" / Marketing's "Top Source":
  // the share text is a plain caption (note), not a colored status pill —
  // this number has no real target to be on/off track against, so it
  // shouldn't be dressed up as a status judgment.
  return {
    status: "neutral",
    statusLabel: "",
    deltaLabel: "",
    note: fact.context,
  };
}

// Converts the AI's alerts ({label, detail}, no status — see route.ts's
// AiNarrativeResult) into BriefingRisk[] by matching each alert back to
// the KPI it's describing and reusing THAT KPI's already-computed status.
// This is the actual mechanism behind the color-consistency rule: status
// is computed once (in classifyHrStatus, above) and every other place it
// appears just looks it up — nothing recomputes it independently, so
// nothing can disagree with it.
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

function hrToBriefing(data: HRData, range: Range): BriefingData {
  const kpis: BriefingFact[] = data.facts.map((f) => {
    const { status, statusLabel, deltaLabel, note } = classifyHrStatus(f, data.companyTargets);
    return {
      id: f.id,
      metric: f.metric,
      value: f.value,
      unit: f.unit,
      status,
      statusLabel,
      note,
      delta: f.delta ? { direction: f.delta.direction, label: deltaLabel } : undefined,
    };
  });

  // Two-point chart. Its start point must match whatever the Watch Items
  // text says the trend "since <Month>" is — both come from the exact same
  // detectSustainedTrend() run, not two independently-chosen starting
  // points, so the chart and the narrative can never disagree on where the
  // trend began (same rule as the color-consistency system). If no
  // sustained run reaches the latest point (e.g. attrition is flat), there's
  // no "since X" claim being made anywhere else on the page either, so it
  // falls back to the plain trailing-12-month first-vs-last reading.
  const attritionSeries = data.monthlySnapshot.map((m) => ({
    label: monthLabel(m.month),
    value: m.attritionRate,
  }));
  const { runs: attritionRuns } = detectSustainedTrend(attritionSeries, "Attrition", 3, true);
  const trailingRun = attritionRuns.find((r) => r.endIndex === data.monthlySnapshot.length - 1);

  const first = data.monthlySnapshot[trailingRun ? trailingRun.startIndex : 0];
  const last = data.monthlySnapshot[data.monthlySnapshot.length - 1];

  const chart = first && last ? {
    title: "Attrition Trend — Start vs. Current",
    subtitle: trailingRun
      ? `${trailingRun.endIndex - trailingRun.startIndex + 1}-month ${trailingRun.direction === "up" ? "upward" : "downward"} run, first vs. most recent reading`
      : "Trailing 12-month attrition, first vs. most recent reading",
    points: [
      { label: monthLabel(first.month), value: first.attritionRate, display: `${first.attritionRate}%` },
      { label: monthLabel(last.month), value: last.attritionRate, display: `${last.attritionRate}%` },
    ],
    gauge: {
      label: "Compliance vs. Target",
      value: last.complianceRate,
      target: data.companyTargets.complianceTargetMin,
      display: `${last.complianceRate}%`,
      footnote: `Target = ${data.companyTargets.complianceTargetMin}% internal minimum`,
    },
  } : undefined;

  const narrative = data.facts
    .map((f) => f.context)
    .filter((c): c is string => !!c)
    .join(" ");

  return {
    reportTitle: "HR & Workforce Performance",
    sectionLabel: "HR View",
    audience: "Executive Leadership",
    dataSource: "Sample HR Dataset (Illustrative)",
    period: "Trailing 12 Months",
    generatedAt: new Date().toLocaleString(),
    kpis,
    chart,
    narrative: narrative || "No additional context available for this period.",
  };
}


function hrAnomaliesAndContext(data: HRData) {
  const attritionSeries = data.monthlySnapshot.map((m) => ({
    label: monthLabel(m.month),
    value: m.attritionRate,
  }));
  const complianceSeries = data.monthlySnapshot.map((m) => ({
    label: monthLabel(m.month),
    value: m.complianceRate,
  }));

  const anomalies = [
    ...detectAnomalies(attritionSeries, "Attrition", {
      target: { max: data.companyTargets.attritionTargetMax },
      valueIsPercentage: true,
    }),
    ...detectAnomalies(complianceSeries, "Compliance", {
      target: { min: data.companyTargets.complianceTargetMin },
      valueIsPercentage: true,
    }),
  ];

  const context = {
    attritionTargetMax: data.companyTargets.attritionTargetMax,
    complianceTargetMin: data.companyTargets.complianceTargetMin,
    currentAttrition: data.monthlySnapshot[data.monthlySnapshot.length - 1]?.attritionRate,
    currentCompliance: data.monthlySnapshot[data.monthlySnapshot.length - 1]?.complianceRate,
  };

  return { anomalies, context };
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
  operationsData?: OperationsData,
  hrData?: HRData
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

  if (hrData) {
    facts.push(...hrToBriefing(hrData, range).kpis.map((f) => toNarrativeFact(f, "hr", "HR")));
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
  department: "marketing" | "operations" | "hr";
  label: string;
  targetRatio: { onTarget: number; total: number; misses: string[] } | null;
  anomalies: AiAnomaly[];
}

// Executive's cross-department Watch Items summary: reads what each tab
// has ALREADY computed (the Excel-baseline target matches from the blend
// feature, plus each tab's existing rule-based anomaly detector) and
// re-labels it by source department. Executive never recomputes a
// judgment of its own here — same principle as executiveCrossDeptFacts
// above, just for risks instead of plain facts. A department with
// neither a target ratio nor any anomalies is omitted entirely, not
// shown as an empty/zero placeholder.
function buildCrossDeptWatchGroups(
  marketingData: MarketingData | undefined,
  operationsData: OperationsData | undefined,
  hrData: HRData | undefined,
  excel: ExcelState | undefined
): DeptWatchGroup[] {
  const groups: DeptWatchGroup[] = [];

  if (marketingData) {
    let targetRatio: DeptWatchGroup["targetRatio"] = null;
    if (excel?.classified) {
      const deptRows = selectDeptRows(excel, "marketing");
      const matches = matchExcelChannelTargets(deptRows, marketingData.channels ?? []);
      if (matches) {
        targetRatio = {
          onTarget: matches.filter((m) => m.onTrack).length,
          total: matches.length,
          misses: matches.filter((m) => !m.onTrack).map((m) => m.channel),
        };
      }
    }
    const anomalies = marketingAnomaliesAndContext(marketingData).anomalies;
    if (targetRatio || anomalies.length > 0) {
      groups.push({ department: "marketing", label: "Marketing", targetRatio, anomalies });
    }
  }

  if (operationsData) {
    let targetRatio: DeptWatchGroup["targetRatio"] = null;
    if (excel?.classified) {
      const matches = matchExcelMetricTargets(excel, [
        {
          metricKey: "bounceRateTarget",
          label: "Bounce Rate",
          actual: Math.round(operationsData.summary.bounceRate * 100),
          lowerIsBetter: true,
        },
        {
          metricKey: "avgEngagementDurationTargetSeconds",
          label: "Avg. Engagement Duration",
          actual: operationsData.summary.avgSessionDuration,
        },
      ]);
      if (matches) {
        targetRatio = {
          onTarget: matches.filter((m) => m.onTrack).length,
          total: matches.length,
          misses: matches.filter((m) => !m.onTrack).map((m) => m.label),
        };
      }
    }
    const anomalies = operationsAnomalies(operationsData);
    if (targetRatio || anomalies.length > 0) {
      groups.push({ department: "operations", label: "Operations", targetRatio, anomalies });
    }
  }

  if (hrData) {
    // HR has no TargetBlendCard (it overrides companyTargets instead —
    // see mergeHrOverrides) — its ratio comes from the same
    // classifyHrStatus/hrToBriefing status every other HR surface
    // already uses, not a re-derivation.
    const hrKpis = hrToBriefing(hrData, "30d").kpis;
    const tracked = hrKpis.filter((k) => k.id === "attrition" || k.id === "compliance");
    const targetRatio: DeptWatchGroup["targetRatio"] =
      tracked.length > 0
        ? {
            onTarget: tracked.filter((k) => k.status !== "critical").length,
            total: tracked.length,
            misses: tracked.filter((k) => k.status === "critical").map((k) => k.metric),
          }
        : null;
    const anomalies = hrAnomaliesAndContext(hrData).anomalies;
    if (targetRatio || anomalies.length > 0) {
      groups.push({ department: "hr", label: "HR", targetRatio, anomalies });
    }
  }

  return groups;
}

// Converts DeptWatchGroup[] into the PDF's BriefingRisk[] — one risk row
// for the ratio (when there are any target-matched rows for that dept)
// plus one row per anomaly, each labeled with its source department so
// the PDF's flat Watch Items list still reads as grouped by origin. Same
// status computation as the on-screen CrossDeptWatchSummary
// (anomalySeverityToStatus) — the two surfaces read the same underlying
// groups, so a signal can't be colored differently between them.
function crossDeptWatchGroupsToRisks(groups: DeptWatchGroup[]): BriefingRisk[] {
  const risks: BriefingRisk[] = [];
  for (const g of groups) {
    if (g.targetRatio) {
      const { onTarget, total, misses } = g.targetRatio;
      risks.push({
        label: `${g.label} — ${onTarget}/${total} Metrics On Target`,
        detail: misses.length > 0 ? `Off target: ${misses.join(", ")}.` : "All matched metrics on target.",
        status: onTarget === total ? "ok" : "watch",
      });
    }
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
    hr?: AiNarrativeResult;
  }>({});

  const [excel, setExcel] = useState<ExcelState>({
    connected: false,
    filename: null,
    rows: null,
    source: null,
    data: null,
    classified: null,
    classifying: false,
    classifyError: null,
  });

  // Classification runs once per upload and is persisted so a page reload
  // reuses it instead of calling the AI again — see EXCEL_BLEND_STORAGE_KEY.
  useEffect(() => {
    try {
      const raw = localStorage.getItem(EXCEL_BLEND_STORAGE_KEY);
      if (!raw) return;
      const saved = JSON.parse(raw) as {
        filename: string | null;
        source: "upload" | "demo" | null;
        data: Record<string, unknown>[] | null;
        classified: ExcelRowClassification[] | null;
      };
      if (saved.data && saved.classified) {
        setExcel({
          connected: true,
          filename: saved.filename,
          rows: saved.data.length,
          source: saved.source,
          data: saved.data,
          classified: saved.classified,
          classifying: false,
          classifyError: null,
        });
      }
    } catch {
      // Corrupt/unavailable storage — start from the empty default, same
      // as never having uploaded anything.
    }
  }, []);

  const [excelError, setExcelError] = useState<string | null>(null);
  const [excelLoading, setExcelLoading] = useState(false);

  const [latestData, setLatestData] = useState<{
    executive?: ExecutiveData;
    operations?: OperationsData;
    marketing?: MarketingData;
    hr?: HRData;
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

  function handleHRData(data: HRData) {
    setLatestData((prev) => ({ ...prev, hr: data }));
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
  function handleHRNarrative(result: AiNarrativeResult) {
    setAiNarratives((prev) => ({ ...prev, hr: result }));
  }


  // Shared by both upload paths (real file / demo file) — one classify
  // call, right after parsing, never re-run per report. Persists the
  // result so a page reload reuses it (see EXCEL_BLEND_STORAGE_KEY).
  async function classifyAndStore(
    filename: string,
    source: "upload" | "demo",
    rows: Record<string, unknown>[]
  ) {
    setExcel((prev) => ({
      ...prev,
      connected: true,
      filename,
      rows: rows.length,
      source,
      data: rows,
      classifying: true,
      classified: null,
      classifyError: null,
    }));
    try {
      const columns = Object.keys(rows[0] ?? {});
      const res = await fetch("/api/excel-classify", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ columns, rows }),
      });
      const json = await res.json();
      if (!res.ok || json.error) {
        throw new Error(json.error || `HTTP ${res.status}`);
      }
      if (json.available === false) {
        throw new Error("AI classification is unavailable (no API key configured).");
      }
      const classified: ExcelRowClassification[] = json.classifications ?? [];
      setExcel((prev) => ({ ...prev, classifying: false, classified }));
      try {
        localStorage.setItem(
          EXCEL_BLEND_STORAGE_KEY,
          JSON.stringify({ filename, source, data: rows, classified })
        );
      } catch {
        // Storage unavailable — classification still works for this
        // session, just won't survive a reload.
      }
    } catch (e) {
      const message = e instanceof Error ? e.message : "Could not classify this data.";
      setExcel((prev) => ({ ...prev, classifying: false, classifyError: message }));
    }
  }

  async function handleExcelFileSelect(file: File) {
  setExcelError(null);
  setExcelLoading(true);

  setExcel((prev) => ({ ...prev, connected: true, filename: file.name, rows: null, source: "upload", data: null }));

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

    await classifyAndStore(file.name, "upload", rows);
  } catch (e) {
    setExcelError(e instanceof Error ? e.message : "Could not parse this file.");
    setExcel((prev) => ({ ...prev, connected: false, filename: null, rows: null, source: null, data: null, classified: null }));
  } finally {
    setExcelLoading(false);
  }
}

async function handleLoadDemoExcel() {
    setDemoLoading(true);
    setDemoError(null);
    try {
      const res = await fetch("/demo-baseline.csv");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const text = await res.text();
      const parsed = Papa.parse<Record<string, unknown>>(text, { header: true, skipEmptyLines: true });
      await classifyAndStore("demo-baseline.csv", "demo", parsed.data);
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
        latestData.operations,
        latestData.hr,
        excel
      );
      narrativeResult = aiNarratives.executive ?? null;
      if (!narrativeResult) {
        inputFacts = executiveCrossDeptFacts(
          latestData.executive,
          range,
          latestData.marketing,
          latestData.operations,
          latestData.hr
        );
        const hrPart = latestData.hr
          ? hrAnomaliesAndContext(latestData.hr)
          : { anomalies: [] as AiAnomaly[], context: undefined as Record<string, number | string> | undefined };
        inputAnomalies = hrPart.anomalies;
        inputContext = hrPart.context;
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
    } else if (tab === "HR" && latestData.hr) {
      briefing = hrToBriefing(latestData.hr, range);
      narrativeResult = aiNarratives.hr ?? null;
      if (!narrativeResult) {
        inputFacts = briefing.kpis;
        const hrPart = hrAnomaliesAndContext(latestData.hr);
        inputAnomalies = hrPart.anomalies;
        inputContext = hrPart.context;
      }
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
        if (tab === "HR") setAiNarratives((prev) => ({ ...prev, hr: generated }));
      }
    }

    if (narrativeResult) {
      briefing.narrative = narrativeResult.overview;
      briefing.keyObservations = narrativeResult.keyObservations;
      // Executive already has its own risks — the cross-department Watch
      // Items summary set by executiveToBriefing above (real, rule-based,
      // correctly labeled by source department). alertsToRisks matches the
      // AI's alerts against THIS tab's own kpis, which for Executive would
      // silently replace that with a weaker match (Executive's AI alerts
      // are seeded from HR anomalies alone, matched against Executive's
      // unrelated active_users/sessions/etc. kpi ids). Every other tab
      // still gets its risks from alertsToRisks as before.
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
    excelSnapshot: excel,
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
           marketing/operations/hr facts are available without visiting
           those tabs first) but must not be display:none — a display:none
           box is 0x0, and Recharts' ResponsiveContainer measures that via
           ResizeObserver, logging "width(0) and height(0)" for every chart
           in every inactive tab on every layout pass. visibility:hidden
           keeps the box laid out (so charts measure their real size) while
           position:absolute removes it from flow so it doesn't reserve
           empty space below the active tab.  */
        .ob-tab-inactive {
          position: absolute;
          inset: 0;
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
                Live GA4 Web Telemetry API + Excel Baseline Blending
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
              onClick={handleLoadDemoExcel}
              disabled={demoLoading}
              className="ob-header-btn text-xs font-bold px-3.5 py-2 rounded-xl flex items-center gap-2 transition active:scale-95 disabled:opacity-60"
              style={{ background: COLORS.accentSoft, color: COLORS.accent, border: "1px solid transparent" }}
            >
              <FileSpreadsheet size={14} />
              {demoLoading ? "Loading…" : "Load Sample Baseline"}
            </button>

            <button
            onClick={handleExportBriefing}
            disabled={briefingLoading}
            className="ob-header-btn text-xs font-bold px-3.5 py-2 rounded-xl flex items-center gap-2 transition active:scale-95 disabled:opacity-60"
            style={{ background: COLORS.indigoSoft, color: COLORS.indigo, border: "1px solid transparent" }}
            >
            {briefingLoading ? <Loader2 size={14} className="animate-spin" /> : <FileText size={14} />}
            {briefingLoading ? "Generating…" : "Export Report"}
            </button>

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
        {excelError && <ErrorBox message={excelError} />}

        <SourcePanel
          activeEndpoint={
            tab === "HR"
              ? "/api/hr (sample, client-side)"
              : `/api/ga4/${tab.toLowerCase()} (${
                  (tab === "Executive"
                    ? tabMs.executive
                    : tab === "Marketing"
                    ? tabMs.marketing
                    : tabMs.operations) ?? "…"
                }ms)`
          }
          excel={excel}

          onExcelFileSelect={handleExcelFileSelect}
          hrActive={tab === "HR"}
          activeDepartment={tab.toLowerCase() as ExcelDepartment}
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
          {tab !== "HR" ? (
            <SegmentedControl options={RANGES} value={range} onChange={setRange} size="sm" variant="accent" />
          ) : (
            <span className="text-xs font-medium" style={{ color: COLORS.inkFaint }}>
              Trailing 12 months
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
              hrData={latestData.hr}
              excel={excel}
            />
          </div>
          <div className={tab === "Marketing" ? "ob-fade-in" : "ob-tab-inactive"}>
            <MarketingView range={range} onData={handleMarketingData} onNarrative={handleMarketingNarrative} narrative={aiNarratives.marketing} excel={excel} />
          </div>
          <div className={tab === "Operations" ? "ob-fade-in" : "ob-tab-inactive"}>
            <OperationsView range={range} onData={handleOperationsData} onNarrative={handleOperationsNarrative} narrative={aiNarratives.operations} excel={excel} />
          </div>
          <div className={tab === "HR" ? "ob-fade-in" : "ob-tab-inactive"}>
            <HRView onData={handleHRData} onNarrative={handleHRNarrative} narrative={aiNarratives.hr} excel={excel} />
          </div>
        </div>

        <div
          className="mt-6 pt-4 flex flex-wrap items-center justify-between gap-2 text-[8px]"
          style={{ borderTop: `1px solid ${COLORS.line}`, color: COLORS.inkFaint }}
        >
          <span>OneBoard · Built with Next.js, TypeScript &amp; the GA4 Data API</span>
          <span>Created on Aug 08, 2026</span>
        </div>
      </main>

      {showPayload && <RawPayloadModal payload={payload} onClose={() => setShowPayload(false)} />}
    </div>
  );
}
