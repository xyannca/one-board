// app/api/ai-narrative/_lib/anomalies.ts
//
// Pure rule-based anomaly detection. Every signal here is produced by
// arithmetic on the input series — no AI, no guessing. The AI Narrative
// route later receives this output and is only allowed to rephrase these
// pre-computed signals into sentences; it cannot invent new ones or judge
// what "looks off" on its own.
//
// This is importable from both server code (the AI narrative route) and
// client code (page.tsx adapters), since it's pure JS/TS with no
// server-only APIs.

import type { AiAnomaly } from "../../../../types/ai-narrative";

export interface TrendPoint {
  label: string; // e.g. "Jul" or "07/31"
  value: number;
}

// Re-exported under its historical name so nothing else in this file has
// to change — the shape now lives in types/ai-narrative.ts as the single
// source of truth.
export type AnomalySignal = AiAnomaly;

// Marks which single-period steps (comparing point i-1 to point i) are
// already covered by a reported sustained-trend run, so the spike/drop
// pass below can skip them instead of restating the same move twice.
export interface TrendRun {
  startIndex: number; // index of the run's first point
  endIndex: number; // index of the run's last point
  direction: "up" | "down";
}

// Finds every maximal run of `minStreak` or more consecutive same-direction
// periods and emits ONE signal per run covering its FULL extent. A 6-month
// upward run is reported once, as a 6-month run with its full cumulative
// change — not chopped into a 3-month alert plus leftover single-step spikes.
// Exported so callers that need the actual detected run — not just the
// resulting sentence — can read it directly (e.g. a trend chart's start
// point) instead of re-deriving it or falling back to an arbitrary index.
// Same "compute once, reuse everywhere" rule the color-status system uses.
export function detectSustainedTrend(
  series: TrendPoint[],
  metricLabel: string,
  minStreak: number,
  valueIsPercentage: boolean
): { signals: AnomalySignal[]; runs: TrendRun[] } {
  const signals: AnomalySignal[] = [];
  const runs: TrendRun[] = [];

  if (series.length < 2) return { signals, runs };

  let runStartStep = 1; // first step index (1-based) of the current run
  let runDir: "up" | "down" | null = null;

  function closeRun(lastStep: number) {
    if (runDir === null) return;
    const stepCount = lastStep - runStartStep + 1;
    const pointCount = stepCount + 1;
    if (pointCount < minStreak) return;

    const startIdx = runStartStep - 1;
    const endIdx = lastStep;
    const first = series[startIdx];
    const last = series[endIdx];

    // Rate-like metrics (already a %, e.g. attrition/compliance) report the
    // change in percentage points — relative % change on a value that's
    // already a percentage reads as misleadingly large (0.7% -> 3.2% is
    // "+2.5pp", not "+357%").
    const changeClause = valueIsPercentage
      ? `, a cumulative ${Math.abs(last.value - first.value).toFixed(1)} percentage point change.`
      : first.value !== 0
      ? `, a cumulative ${Math.abs(((last.value - first.value) / Math.abs(first.value)) * 100).toFixed(0)}% change.`
      : `.`;

    signals.push({
      id: `${metricLabel}-trend-${first.label}-${last.label}`,
      type: "sustained_trend",
      label: `${metricLabel} trending ${runDir === "up" ? "upward" : "downward"} since ${first.label}`,
      detail:
        `${metricLabel} has moved ${runDir === "up" ? "up" : "down"} for ${pointCount} consecutive periods, ` +
        `from ${first.value} in ${first.label} to ${last.value} in ${last.label}` + changeClause,
      severity: "info",
    });


    runs.push({ startIndex: startIdx, endIndex: endIdx, direction: runDir });
  }

  for (let step = 1; step < series.length; step++) {
    const diff = series[step].value - series[step - 1].value;
    const dir: "up" | "down" | null = diff > 0 ? "up" : diff < 0 ? "down" : null;

    if (step === 1) {
      runStartStep = step;
      runDir = dir;
      continue;
    }

    if (dir === runDir && dir !== null) {
      continue; // still the same run — keep extending, nothing to close yet
    }

    // direction changed (or hit a flat point) — close the run that just
    // ended, then start tracking a new one from here.
    closeRun(step - 1);
    runStartStep = step;
    runDir = dir;
  }
  closeRun(series.length - 1); // close whatever run reaches the final point

  return { signals, runs };
}

// Flags a single-period move larger than `thresholdPct`, skipping any step
// already covered by a sustained-trend run for this metric — the run
// already tells that story more completely, so restating one step inside
// it as a separate spike would just repeat the same fact.
function detectSpikesAndDrops(
  series: TrendPoint[],
  metricLabel: string,
  thresholdPct: number,
  coveredRuns: TrendRun[]
): AnomalySignal[] {
  const signals: AnomalySignal[] = [];
  for (let i = 1; i < series.length; i++) {
    const isCovered = coveredRuns.some((r) => i - 1 >= r.startIndex && i <= r.endIndex);
    if (isCovered) continue;

    const prev = series[i - 1].value;
    const curr = series[i].value;
    if (prev === 0) continue; // avoid divide-by-zero; can't compute a % change from zero
    const pctChange = ((curr - prev) / prev) * 100;
    if (Math.abs(pctChange) >= thresholdPct) {
      const direction = pctChange > 0 ? "spike" : "drop";
      signals.push({
        id: `${metricLabel}-${direction}-${series[i].label}`,
        type: direction,
        label: `${metricLabel} ${direction} in ${series[i].label}`,
        detail: `${metricLabel} moved from ${prev} to ${curr} between ${series[i - 1].label} and ${series[i].label}, a ${Math.abs(pctChange).toFixed(0)}% ${direction}.`,
        severity: Math.abs(pctChange) >= thresholdPct * 2 ? "warning" : "info",
      });
    }
  }
  return signals;
}

// Flags the latest value breaching a stated target/threshold — the target
// itself must come from the caller (e.g. HRDataset.companyTargets), never
// invented here.
function detectTargetBreach(
  series: TrendPoint[],
  metricLabel: string,
  target: { max?: number; min?: number }
): AnomalySignal[] {
  if (series.length === 0) return [];
  const latest = series[series.length - 1];
  const signals: AnomalySignal[] = [];

  if (target.max !== undefined && latest.value > target.max) {
    signals.push({
      id: `${metricLabel}-breach-max-${latest.label}`,
      type: "target_breach",
      label: `${metricLabel} above internal target`,
      detail: `${metricLabel} is ${latest.value} in ${latest.label}, above the internal target of ${target.max}.`,
      severity: "warning",
    });
  }
  if (target.min !== undefined && latest.value < target.min) {
    signals.push({
      id: `${metricLabel}-breach-min-${latest.label}`,
      type: "target_breach",
      label: `${metricLabel} below internal target`,
      detail: `${metricLabel} is ${latest.value} in ${latest.label}, below the internal target of ${target.min}.`,
      severity: "warning",
    });
  }
  return signals;
}

export function detectAnomalies(
  series: TrendPoint[],
  metricLabel: string,
  options?: {
    spikeThresholdPct?: number;
    trendMinStreak?: number;
    target?: { max?: number; min?: number };
    valueIsPercentage?: boolean;
  }
): AnomalySignal[] {
  if (series.length < 2) return [];

  const spikeThresholdPct = options?.spikeThresholdPct ?? 25;
  const trendMinStreak = options?.trendMinStreak ?? 3;
  const valueIsPercentage = options?.valueIsPercentage ?? false;

  const { signals: trendSignals, runs } = detectSustainedTrend(series, metricLabel, trendMinStreak, valueIsPercentage);

  const spikeSignals = detectSpikesAndDrops(series, metricLabel, spikeThresholdPct, runs);

  const signals: AnomalySignal[] = [...trendSignals, ...spikeSignals];

  if (options?.target) {
    signals.push(...detectTargetBreach(series, metricLabel, options.target));
  }

  return signals;
}