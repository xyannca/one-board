"use client";

import { useEffect, useRef, useState } from "react";
import type { ReactNode, ComponentType } from "react";
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
  // Real period-over-period comparison (e.g. this 30 days vs. the prior 30
  // days) computed server-side — not derived from the trend chart's
  // endpoints, which understates change on longer ranges.
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

const TABS = ["Executive", "Marketing", "Operations"] as const;
type Tab = (typeof TABS)[number];
const TAB_LABELS: Record<Tab, string> = {
  Executive: "Executive View",
  Marketing: "Marketing & Funnel",
  Operations: "Operations & Diagnostics",
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
  bg: "#F8FAFC",
  surface: "#FFFFFF",
  ink: "#0F172A",
  inkSoft: "#64748B",
  inkFaint: "#94A3B8",
  line: "#E2E8F0",
  track: "#F1F5F9",
  accent: "#059669",
  accentSoft: "#D1FAE5",
  up: "#059669",
  upSoft: "#D1FAE5",
  down: "#DC2626",
  downSoft: "#FEE2E2",
  amberSoft: "#FEF3C7",
  amberInk: "#92400E",
  blue: "#2563EB",
  blueSoft: "#DBEAFE",
  teal: "#0D9488",
  indigo: "#4F46E5",
  indigoSoft: "#E0E7FF",
};
const DONUT_COLORS = ["#059669", "#2563EB", "#0D9488", "#4F46E5", "#94A3B8", "#CBD5E1"];
const SHADOW = "0 1px 3px rgba(15,23,42,0.04), 0 1px 2px -1px rgba(15,23,42,0.03)";
const SHADOW_HOVER = "0 12px 24px -8px rgba(5,150,105,0.18)";
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
    // Deliberately not clearing `data` here — keeping the previous chart on
    // screen while a new range loads avoids re-mounting ResponsiveContainer,
    // which caused a "narrow then widen" flash on earlier tab switches.
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
}: {
  options: readonly { value: T; label: ReactNode }[];
  value: T;
  onChange: (v: T) => void;
  size?: "sm" | "md";
}) {
  const pad = size === "sm" ? "px-3 py-1.5 text-xs" : "px-4 py-2 text-xs";
  return (
    <div className="inline-flex p-1 rounded-2xl gap-0.5" style={{ background: "#E2E8F0B3" }}>
      {options.map((o) => {
        const active = o.value === value;
        return (
          <button
            key={o.value}
            onClick={() => onChange(o.value)}
            className={`rounded-xl font-semibold transition-all ${pad}`}
            style={{
              color: active ? COLORS.ink : COLORS.inkSoft,
              background: active ? COLORS.surface : "transparent",
              boxShadow: active ? "0 1px 2px rgba(15,23,42,0.06), 0 1px 6px rgba(15,23,42,0.05)" : "none",
              border: active ? `1px solid ${COLORS.line}` : "1px solid transparent",
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
      style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW }}
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
      style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW }}
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
      style={{ background: COLORS.ink, color: "#fff", boxShadow: SHADOW_HOVER }}
    >
      <p className="mb-1" style={{ color: "#94A3B8" }}>
        {label}
      </p>
      {payload.map((p: any, i: number) => (
        <p key={i} className="font-semibold" style={{ fontVariantNumeric: "tabular-nums" }}>
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
};

function SourcePanel({
  activeEndpoint,
  excel,
  onExcelFileSelect,
}: {
  activeEndpoint: string;
  excel: ExcelState;
  onExcelFileSelect: (file: File) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
      {/* GA4 */}
      <div
        className="ob-ga4-card rounded-2xl p-4 hover:-translate-y-0.5"
        style={{
          background: `linear-gradient(to bottom, #fff, ${COLORS.accentSoft}66)`,
          border: `1px solid #A7F3D0`,
          boxShadow: SHADOW,
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

      {/* Excel / CSV — real file picker, green hover state to match GA4's active look */}
      <div
        onClick={() => inputRef.current?.click()}
        className="ob-excel-card rounded-2xl p-4 cursor-pointer hover:-translate-y-0.5"
        style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW }}
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
          style={{ color: excel.connected ? COLORS.accent : COLORS.inkFaint }}
        >
          <Upload size={13} />
          <span>
            {excel.connected
              ? excel.rows !== null
                ? `${excel.rows} rows parsed from ${excel.filename ?? "file"} — ready to blend`
                : "File selected — ready to parse"
              : "+ Upload Offline Dataset (.xlsx/.csv)"}
          </span>
        </div>
      </div>

      {/* ERP */}
      <div
        className="rounded-2xl p-4 opacity-80"
        style={{ background: "#F8FAFC", border: `1px solid ${COLORS.line}` }}
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

// Common mismatches between GA4's country names and the map's Natural Earth
// names. Not exhaustive — most countries match directly.
const COUNTRY_ALIASES: Record<string, string> = {
  "united states": "united states of america",
  russia: "russian federation",
  "south korea": "republic of korea",
  "north korea": "dem. rep. korea",
  czechia: "czech republic",
  vietnam: "viet nam",
};

// Approximate [lon, lat] centroids for labeling countries that actually show
// up in GA4 traffic data. Not exhaustive — covers common markets. Countries
// without a listed centroid still get colored on the map, just no label pin.
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
    if (value === null || max === 0) return "#E9EEF1";
    const intensity = 0.28 + 0.72 * (value / max);
    return `rgba(5, 150, 105, ${intensity})`;
  }

  if (rows.length === 0) {
    return <p className="text-sm" style={{ color: COLORS.inkFaint }}>No data in this range.</p>;
  }

  // Label the top 5 countries by traffic that have a known centroid — a pill
  // per country keeps it readable; more than ~5 and labels start to overlap.
  const MARKER_COLORS = [COLORS.accent, COLORS.blue, COLORS.indigo, COLORS.teal, "#B45309"];
  const labeled = [...rows]
    .sort((a, b) => b.activeUsers - a.activeUsers)
    .map((r) => ({ ...r, centroid: COUNTRY_CENTROIDS[r.country.trim().toLowerCase()] }))
    .filter((r): r is typeof r & { centroid: [number, number] } => !!r.centroid)
    .slice(0, 5);

  return (
    <div style={{ position: "relative" }}>
      <ComposableMap
        projection="geoEqualEarth"
        // Centered/zoomed toward North America rather than a full-world view —
        // this dashboard's traffic is concentrated there, so a full globe just
        // wastes space on oceans and continents with no data.
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
                    default: { fill: colorFor(value), stroke: "#fff", strokeWidth: 0.5, outline: "none" },
                    hover: {
                      fill: value !== null ? "#047857" : colorFor(value),
                      stroke: "#fff",
                      strokeWidth: 0.5,
                      outline: "none",
                    },
                    pressed: { fill: "#065F46", stroke: "#fff", strokeWidth: 0.5, outline: "none" },
                  }}
                />
              );
            })
          }
        </Geographies>

        {/* Pass 1: ping rings + solid dots for every labeled marker, drawn
            before any pill. SVG paints later elements on top, so if pills
            were interleaved per-marker, one country's pill could visually
            cover a neighboring country's pulse ring. Two passes avoids that. */}
        {labeled.map((r, i) => {
          const dotColor = MARKER_COLORS[i % MARKER_COLORS.length];
          return (
            <Marker key={`dot-${r.country}`} coordinates={r.centroid}>
              {/* Ping ring on every labeled marker — any country with real
                  traffic shows as "live". Two layers: a white base so the
                  ring stays visible even over a same-hue green landmass
                  (previously invisible over Canada's dark-green fill), plus a
                  colored outline on top for per-country identity. */}
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

        {/* Pass 2: the pill labels, always on top of every dot/ping. Sized
            larger than desktop strictly needs, since this SVG scales down
            to fit a phone-width card — text this size on a ~340px-wide
            mobile container renders at roughly half these viewBox pixels. */}
        {labeled.map((r, i) => {
          const dotColor = MARKER_COLORS[i % MARKER_COLORS.length];
          const label = `${r.country} (${r.activeUsers.toLocaleString()})`;
          // Rough monospace-independent width estimate so the pill fits the text.
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
                  fill="#fff"
                  stroke={dotColor}
                  strokeWidth={1.5}
                  style={{ filter: "drop-shadow(0 1px 2px rgba(15,23,42,0.15))" }}
                />
                <circle cx={14} cy={13} r={3.5} fill={dotColor} />
                <text
                  x={23}
                  y={18}
                  style={{ fontSize: 14, fontWeight: 700, fill: COLORS.ink, fontFamily: FONT_STACK }}
                >
                  {label}
                </text>
              </g>
            </Marker>
          );
        })}
      </ComposableMap>
      {hover && (
        <div
          className="text-xs font-semibold px-2.5 py-1 rounded-lg"
          style={{
            position: "absolute",
            top: 8,
            left: 8,
            background: COLORS.ink,
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

function LiveInsightsCard({ input }: { input: InsightsInput | null }) {
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
        background: "linear-gradient(160deg, #1E293B, #334155)",
        color: "#fff",
        boxShadow: SHADOW,
      }}
    >
      <div>
        <div
          className="flex items-center justify-between pb-3 mb-3"
          style={{ borderBottom: "1px solid rgba(255,255,255,0.1)" }}
        >
          <div className="flex items-center gap-2 text-xs font-bold" style={{ color: "#34D399" }}>
            <Sparkles size={14} strokeWidth={2} />
            Live Insights
          </div>
          <span
            className="text-[10px] font-semibold px-2 py-0.5 rounded-full"
            style={{ background: "rgba(255,255,255,0.08)", color: "#94A3B8" }}
          >
            From live data
          </span>
        </div>

        {!input ? (
          <p className="text-xs" style={{ color: "#94A3B8" }}>
            Loading…
          </p>
        ) : (
          <ul className="space-y-2.5 text-xs leading-relaxed" style={{ color: "#CBD5E1" }}>
            {bullets.map((b, i) => (
              <li key={i} className="flex items-start gap-2">
                <span style={{ color: "#34D399" }} className="font-bold">
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
          style={{ background: "rgba(255,255,255,0.08)", color: "#fff" }}
        >
          {copied ? <CheckCircle2 size={13} /> : <Copy size={13} />}
          {copied ? "Copied" : "Copy Insights"}
        </button>
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

function ExecutiveView({
  range,
  onData,
}: {
  range: Range;
  onData?: (data: ExecutiveData, clientMs?: number) => void;
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

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-4">
        <Card
          title={`Live Sessions Trend · ${rangeLabel(range)}`}
          icon={TrendingUp}
          badge={activeUsersDelta !== null ? <DeltaBadge value={activeUsersDelta} /> : undefined}
          className="lg:col-span-2"
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

        <LiveInsightsCard
          input={{
            range: rangeLabel(range),
            activeUsers: data.summary.activeUsers,
            sessions: data.summary.sessions,
            engagementRate: data.summary.engagementRate,
            topChannel: data.summary.topChannel,
            topChannelSharePct: topChannelShare,
            topCountry: data.geoCountries?.[0]?.country ?? null,
            topCountrySharePct:
              data.geoCountries && data.geoCountries.length > 0 && totalGeoUsers > 0
                ? (data.geoCountries[0].activeUsers / totalGeoUsers) * 100
                : null,
          }}
        />
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

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <Card title="Geographical Telemetry Map" icon={Globe2} className="lg:col-span-2">
          <CountryChoropleth rows={data.geoCountries ?? []} />
        </Card>

        <BarListChart
          title={`Active Metro Cities · ${rangeLabel(range)}`}
          icon={MapPin}
          rows={(data.geoCities ?? []).map((g) => ({ label: g.city, value: g.activeUsers }))}
          barColor={COLORS.accent}
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
    // Reuses COLORS.inkFaint — the same grey the Traffic Source donut below
    // uses for its smaller slices — instead of introducing a new color.
    steps.push({ label: "Key events", value: funnel.keyEvents, color: COLORS.inkFaint });
  }
  const max = steps[0]?.value || 1;

  // Key events being exactly 0 usually means no GA4 event is marked as a
  // "Key event" yet (Admin → Events), not that conversion genuinely failed —
  // the query still succeeds and returns 0 either way, so we can't tell the
  // two cases apart from the number alone. Rather than show a permanently
  // uninformative "0.0% completion," fall back to engaged-session rate,
  // which is always a real, non-zero signal, and say so plainly.
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
          // No bar at all for a genuine zero — a minimum-width grey sliver
          // reads as "some small amount happened" when actually nothing did.
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

function MarketingView({ range }: { range: Range }) {
  const { data, error, isFetching } = useGa4<MarketingData>("/api/ga4/marketing", range);

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
        <div className="lg:col-span-2">
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
          // 2nd slice is grey rather than teal, so it doesn't read as a second
          // shade of green next to the top (accent) slice.
          colors={[COLORS.accent, COLORS.inkFaint, COLORS.blue, COLORS.indigo, COLORS.teal, "#CBD5E1"]}
        />
        <BarListChart
          title={`Landing Page Acquisition · ${rangeLabel(range)}`}
          icon={Flag}
          barColor={COLORS.teal}
          rows={data.landingPages.map((p) => ({ label: p.page, value: p.sessions }))}
        />
      </div>
    </div>
  );
}

/* ============== Operations ============== */

function OperationsView({
  range,
  onData,
}: {
  range: Range;
  onData?: (data: OperationsData, clientMs?: number) => void;
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
          <ShareDonut title={`Device Category · ${rangeLabel(range)}`} icon={Smartphone} rows={deviceRows} />
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

      {/* 3. Bottom — general page traffic on the left (moved here from
          Marketing, since 404s are an operational signal), diagnostic
          route logs on the right, presented as routine monitoring data. */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <BarListChart
          title={`Top Content Pages · ${rangeLabel(range)}`}
          icon={FileText}
          barColor={COLORS.blue}
          rows={(data.pages ?? []).map((p) => ({ label: p.title, value: p.views }))}
        />

        <Card title="Diagnostic Route Logs" icon={FileWarning} badge={<span className="text-[11px] font-medium" style={{ color: COLORS.inkFaint }}>Live Endpoint Inspector</span>}>
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
                        className="truncate text-sm font-medium"
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
        </Card>
      </div>
    </div>
  );
}

/* ============== Page ============== */

export default function Home() {
  const [tab, setTab] = useState<Tab>("Executive");
  const [range, setRange] = useState<Range>(DEFAULT_RANGE);
  const [showPayload, setShowPayload] = useState(false);
  const [demoLoading, setDemoLoading] = useState(false);
  const [demoError, setDemoError] = useState<string | null>(null);

  const [excel, setExcel] = useState<ExcelState>({
    connected: false,
    filename: null,
    rows: null,
    source: null,
  });

  const [latestData, setLatestData] = useState<{
    executive?: ExecutiveData;
    operations?: OperationsData;
  }>({});
  const [lastMs, setLastMs] = useState<number | undefined>(undefined);

  function handleExecutiveData(data: ExecutiveData, ms?: number) {
    setLatestData((prev) => ({ ...prev, executive: data }));
    if (ms !== undefined) setLastMs(ms);
  }

  function handleOperationsData(data: OperationsData, ms?: number) {
    setLatestData((prev) => ({ ...prev, operations: data }));
    if (ms !== undefined) setLastMs(ms);
  }

  function handleExcelFileSelect(file: File) {
    setExcel({ connected: true, filename: file.name, rows: null, source: "upload" });
  }

  async function handleLoadDemoExcel() {
    setDemoLoading(true);
    setDemoError(null);
    try {
      const res = await fetch("/demo-baseline.csv");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const text = await res.text();
      const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
      setExcel({
        connected: true,
        filename: "demo-baseline.csv",
        rows: parsed.data.length,
        source: "demo",
      });
    } catch (e) {
      setDemoError("Could not load the demo dataset.");
    } finally {
      setDemoLoading(false);
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
                Auto-Dashboard Engine · Live GA4 Web Telemetry API + Excel Baseline Blending
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <div
              className="flex items-center gap-2 text-xs font-medium px-3.5 py-2 rounded-xl"
              style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}`, boxShadow: SHADOW, fontFamily: MONO_STACK }}
              title="Round-trip time for the most recently loaded GA4 API response"
            >
              <span className="w-1.5 h-1.5 rounded-full ob-live-dot" style={{ background: COLORS.up }} />
              <Zap size={13} style={{ color: "#F59E0B" }} />
              <span style={{ color: COLORS.inkSoft }}>{lastMs !== undefined ? `${lastMs}ms` : "…"}</span>
            </div>

            <button
              onClick={handleLoadDemoExcel}
              disabled={demoLoading}
              className="ob-header-btn text-xs font-bold px-3.5 py-2 rounded-xl flex items-center gap-2 transition active:scale-95 disabled:opacity-60"
              style={{ background: COLORS.accentSoft, color: COLORS.accent, border: "1px solid transparent" }}
            >
              <FileSpreadsheet size={14} />
              {demoLoading ? "Loading…" : "Blend Demo Excel"}
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

        <SourcePanel
          activeEndpoint={`/api/ga4/${tab.toLowerCase()} (${lastMs ?? "…"}ms)`}
          excel={excel}
          onExcelFileSelect={handleExcelFileSelect}
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
            />
          </div>
          <SegmentedControl options={RANGES} value={range} onChange={setRange} size="sm" />
        </div>

        <div className={tab === "Executive" ? "ob-fade-in" : "hidden"}>
          <ExecutiveView range={range} onData={handleExecutiveData} />
        </div>
        <div className={tab === "Marketing" ? "ob-fade-in" : "hidden"}>
          <MarketingView range={range} />
        </div>
        <div className={tab === "Operations" ? "ob-fade-in" : "hidden"}>
          <OperationsView range={range} onData={handleOperationsData} />
        </div>
      </main>

      {showPayload && <RawPayloadModal payload={payload} onClose={() => setShowPayload(false)} />}
    </div>
  );
}
