
// app/api/hr/_lib/hrFacts.ts
//
// Turns the static sample HR dataset (public/sample-hr-dataset.json) into
// the same Fact[] shape the GA4 routes return, so the front end and any
// future LLM narrative layer can treat GA4 facts and HR facts identically.
//
// All values here are computed from the dataset at request time — nothing
// is hardcoded, and nothing is sourced from an external "industry
// benchmark." The only static numbers used are `companyTargets` in the
// dataset itself, which are explicitly labeled as internal targets, not
// industry figures.

export interface MonthlySnapshot {
  month: string;
  headcount: number;
  hires: number;
  terminations: number;
  attritionRate: number;
  complianceRate: number;
}

export interface HRDataset {
  meta: { label: string; generatedAt: string; note: string };
  companyTargets: { attritionTargetMax: number; complianceTargetMin: number };
  monthlySnapshot: MonthlySnapshot[];
  employees: unknown[];
  topHireSource: string;
  topHireSourceShare: number;
}

export interface Fact {
  id: string;
  metric: string;
  value: number | string;
  unit?: string;
  delta?: { value: number; direction: "up" | "down"; period: string };
  context?: string;
}

// Same percent-change convention as ga4/_lib/comparison.ts: null when the
// prior value is zero and current isn't, so the UI can show "—" instead of
// a misleading Infinity.
function percentDelta(current: number, previous: number): number | null {
  if (previous === 0) return current === 0 ? 0 : null;
  return ((current - previous) / previous) * 100;
}

function toDeltaField(current: number, previous: number, period: string) {
  const pct = percentDelta(current, previous);
  if (pct === null) return undefined;
  const direction: "up" | "down" = pct >= 0 ? "up" : "down";
  return { value: Number(pct.toFixed(1)), direction, period };
}

// For metrics that are ALREADY a rate/percentage (attritionRate,
// complianceRate), a "percent change of a percent" is misleading —
// e.g. 2.5% -> 3.2% is a 28% relative jump, but reads as if attrition
// rose 28 percentage points. Report the raw percentage-point (pp)
// difference instead, which is how HR reporting conventionally states
// rate-over-rate movement.
function toPointDeltaField(current: number, previous: number, period: string) {
  const diff = current - previous;
  const direction: "up" | "down" = diff >= 0 ? "up" : "down";
  return { value: Number(Math.abs(diff).toFixed(1)), direction, period };
}

export function computeHRFacts(dataset: HRDataset): Fact[] {
  const snapshot = dataset.monthlySnapshot;
  const latest = snapshot[snapshot.length - 1];
  const prior = snapshot[snapshot.length - 2];

  const facts: Fact[] = [
    {
      id: "headcount",
      metric: "Total Headcount",
      value: latest.headcount,
      delta: prior ? toDeltaField(latest.headcount, prior.headcount, "MoM") : undefined,
    },
    {
      id: "attrition",
      metric: "Trailing 12-Month Attrition",
      value: latest.attritionRate,
      unit: "%",
      delta: prior ? toPointDeltaField(latest.attritionRate, prior.attritionRate, "MoM, pp") : undefined,
      context:
        latest.attritionRate <= dataset.companyTargets.attritionTargetMax
          ? `Within internal target (<${dataset.companyTargets.attritionTargetMax}%)`
          : `Above internal target (<${dataset.companyTargets.attritionTargetMax}%)`,
    },
    {
      id: "compliance",
      metric: "Training Compliance",
      value: latest.complianceRate,
      unit: "%",
      delta: prior ? toPointDeltaField(latest.complianceRate, prior.complianceRate, "MoM, pp") : undefined,
      context:
        latest.complianceRate >= dataset.companyTargets.complianceTargetMin
          ? `Meets internal target (>${dataset.companyTargets.complianceTargetMin}%)`
          : `Below internal target (>${dataset.companyTargets.complianceTargetMin}%)`,
    },
    {
      id: "top_hire_source",
      metric: "Top Hire Source",
      value: dataset.topHireSource,
      context: `${dataset.topHireSourceShare}% of all recorded hires`,
    },
  ];

  return facts;
}

// NOTE: this deliberately does not compute or return any cross-metric
// "correlation" facts (the relatedFactIds pattern from the OneBoard
// upgrade plan). The HR dataset is monthly and low-volume, too coarse for
// a meaningful correlation coefficient — forcing one here would be exactly
// the kind of unsupported claim the Rules Engine / LLM Narrative split is
// meant to prevent. Add it only if a real statistical test is implemented.
