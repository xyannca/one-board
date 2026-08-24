# OneBoard Executive Briefing — Design Spec (McKinsey-style, one-page)

Status: Finalized for HR tab (Aug 2026). Apply this same structure to Executive, Marketing, and Operations tabs.

## Core principle (non-negotiable)

Every number must trace to a real computed value. Never invent dimensions, causal
explanations, or benchmarks that aren't in the underlying dataset. If a drill-down
table isn't backed by real data, don't add one — a shorter honest report beats a
denser fabricated one.

## Color system

Three status colors, used consistently and ONLY for status meaning — never decorative:

- **Red** `#b03a2e` — metric has crossed its target/threshold (critical)
- **Amber** `#a06c00` — metric is trending the wrong direction but still within
  target/threshold (watch)
- **Green** `#0f7a44` — on track / no action needed
- **Black/gray** (`#10151f` text, `#5a6472` secondary) — everything structural:
  headers, borders, numbered circles, labels. No blue or other accent colors —
  they compete with the status colors and dilute the "at a glance" read.

**Hard rule:** color is computed ONCE per data point and reused everywhere that data
point appears — KPI card, watch item, drill-down table row, status pill. Never let
two components describing the same metric disagree on color (this was the biggest
bug found during review — e.g. an "Attrition" KPI card and its own Watch Item
initially had different colors for the same number).

Icons: use CSS-drawn triangle (warning) and circle (checkmark) badges, not emoji.
Emoji glyphs render unreliably/faintly depending on font availability — CSS shapes
are guaranteed and match the reference styling better anyway.

## Page structure, top to bottom

1. **Header**: Real OneBoard logo mark (rounded square, brand green, "O") +
   "OneBoard" wordmark (bold) + "Executive Briefing" (light gray, same line).
   Title line: `[Tab] Performance` + INTERNAL tag (amber outline pill).
   Right-aligned meta block: Period / Audience / Generated — three fields only,
   no Scope or Status field (redundant with title + summary). Audience stays
   "Executive Leadership" across all tabs (it describes the reader, not the topic —
   don't vary it per tab). Period and the title DO vary per tab.

2. **Executive Summary**: Black left border (4px), no background fill, label
   "EXECUTIVE SUMMARY" in black (not blue), no "30-second brief" tag. Body text
   leads with the real conclusion as the first sentence (bold the key clause),
   not a generic label — this is the actual "action title" principle from the
   Pyramid Principle, just placed inside the summary box rather than as a giant
   page headline.

3. **KPI strip** (3-4 cards): icon + label, bold large value, delta line, status
   pill. Delta line rules:
   - If the metric has a real target, state it inline: `▲ 0.7pp (< 8% Target)` —
     don't make the reader look elsewhere to know if a number is good or bad.
   - Color/icon severity must match the pill below it (see hard rule above).
   - Small/noise-level changes (e.g. -0.3% MoM on a 293 headcount) should NOT get
     warning color just because they're negative — direction (arrow) and severity
     (color) are separate signals. Only color-flag changes that are actually
     significant.

4. **Chart + Watch Items** (two columns): chart shows only real data points —
   if you only have two readings (start/end of period), show two bars, don't
   draw a fabricated smooth trend line through invented intermediate months.
   Watch items: short (1-2 sentences), color-matched to the KPI they describe,
   CSS badge icon, not paragraphs.

5. **Drill-down table** (add only if real dimensional data exists — see below):
   this is what separates a "summary" from an "analysis." A single 94.1%
   compliance number states a fact; a table breaking it down by department/device/
   browser explains why. Same color system applies to the table's status column.

6. **Recommended Next Steps** (2 cards): white background, light gray border only
   (no color fill), black numbered circles, visible gap between cards (~26px).
   Each recommendation must be a decision/threshold-based action grounded in real
   numbers already shown — never a fabricated causal story (e.g. NOT "due to
   conflicting client schedules" unless that's an actual tracked field).

7. **Footer**: one-line discipline statement — every figure traces to live data,
   no unobserved figures or causal claims added during export.

## Second surface using this pattern: AiNarrativeCard (on-screen, not PDF)

Not yet started. `AiNarrativeCard` is the on-dashboard summary card (the one
with the "Generate AI Summary" button — separate component from
`ExecutiveBriefingPDF`, rendered inline in `ExecutiveView`/`MarketingView`/
`OperationsView`/`HRView`, not exported as a file). Current version is too
long/repetitive and doesn't read as a quick "live insight" — same complaint
that drove the PDF redesign.

Direction: give it the same structure as the finalized one-pager
(Executive Summary → KPIs → Watch Items → Recommendations), same status/
color-consistency rules, just without icons (it's a compact on-screen card,
not a printed report). Same `classifyStatus \u2192 alertsToRisks` pattern
should apply — don't re-derive a separate coloring scheme for this
component. Read the current `AiNarrativeCard.tsx` before changing it (its
props are `facts`, `anomalies`, `context`, `result`, `onGenerated` — shape
partially visible from call sites in page.tsx, but the component body
itself hasn't been reviewed yet).

## Per-tab drill-down data status (verify before building)

- **Marketing & Operations**: CONFIRMED real dimensional data exists — Device
  Category (Desktop/Mobile), Browser Matrix, Operating System, Top Content Pages
  are all live GA4 data already surfaced in the dashboard (see Operations tab
  screenshot). Build the drill-down table for these tabs using these real fields.
- **HR**: NOT YET CONFIRMED whether the current synthetic HR dataset has a
  department (or other) dimension to drill into. Check the data generation logic
  before adding a table — do not add a department breakdown unless it's real.
  If no such dimension exists, the HR one-pager stays at KPI + trend + watch items
  (current finalized version), no table.
- **Executive** (cross-functional): scope not yet decided — likely needs its own
  pass since it spans HR + Marketing + Operations rather than one dataset.

## Known technical constraint for implementation

This spec was designed and visually verified as static HTML/CSS (rendered via
wkhtmltoimage, iterated by actually looking at output, not guessing from code).
`@react-pdf/renderer`'s styling is more limited than browser CSS — flexbox support
is partial, and `::before`/`::after` pseudo-elements (used for the CSS triangle/
circle badges) aren't supported. When translating to `ExecutiveBriefingPDF.tsx`,
these will need reimplementing (e.g. badges as actual small SVG or View-based
shapes) — expect this translation step to require real adaptation, not a direct
copy-paste, and confirm the rendered PDF still matches visually before considering
it done.