# Dashboard specs — layout and metric definitions

What `scripts/build_dashboards.py` produces, and the exact logic behind each number. Read this when modifying the script, explaining a metric, or rebuilding a section by hand.

## Shared definitions

- **Open opp** — Stage is one of the config open stages (not won, not lost).
- **Weighted Value** — Value × stage probability from config; won = 100%, lost = 0. The probability table is a client decision — surface it, don't bury it.
- **Forecast buckets** — from `Forecast Category` mapped via config: Pipeline, Best Case, Commit, Closed.
- **Strong Best Case** — Best Case AND `Include Forecast` = 1 (when the flag is in use; otherwise Strong Best Case = all Best Case and the "Remaining Best Case" column is 0).
- **Bookings / Won Value** — sum of Value where Stage = won stage, Close Date in period.
- **Total Pipeline** — Won + Commit + Best Case + Pipeline value in scope.
- **Gap to Target** — Target − metric. Negative = ahead of plan.
- **Pipe Coverage (weighted)** — weighted open pipeline ÷ (Target − Won). If Won ≥ Target, shown as "target met".
- **ASP** — Bookings ÷ # won opps.
- **Attainment** — Bookings ÷ Quota. **Forecast Attainment** — (Won + Commit) ÷ Quota.
- **Large-deal mix** — share of a rep's open pipeline in the config `large_bands`.
- **Deal band** — Value bucketed by config `deal_bands` thresholds (ascending `min`).
- **Fiscal periods** — quarters/months derived from Close Date using `fiscal_year_start_month`; FY labelled by the calendar year it starts in.

## Tab: README

Engagement name, as-of date, data sources, row counts, what was skipped (missing optional columns), and an **Assumptions to confirm** list — every config value not explicitly confirmed by the client (probabilities, band thresholds, mapped fields). This list is the honesty layer; keep it.

## Tab 1 — Quarterly Sales Forecast (current fiscal quarter)

1. **KPI scorecard** (4 × 3 grid): Won; Commit + Won; Strong Best Case + Commit + Won; Weighted Pipeline — each with % of quarter target and Gap to Target. Then Pipe Coverage (weighted) and Pipe Coverage (strong best case). Plus one KPI per configured deal-quality flag (flagged count / eligible count).
2. **Forecast Confidence by Month** — rows = months of the quarter; columns = Won Value, # Won, Commit Value, # Commit, Strong Best Case, Remaining Best Case, # Best Case, Pipeline Value, # Pipeline, Total Pipeline, Weighted Pipeline. Totals row.
3. **Weighted Forecast — Stage × Month** — weighted value, columns = Won + each open stage (funnel order), grand total.
4. **Pipeline by Stage × Forecast Category** — unweighted value matrix.
5. **Closed Won (this quarter)** — Owner, Type, Opp link/name, ICP Grade, Close Date, Deal Cycle days (Close − Created), Value.
6. **Moved to Lost (last 7 days from as-of)** — Owner, Type, name, ICP Grade, Lead Source, Loss Reason, Close Date, Value.
7. **Deal-quality exception tables** — one per configured flag: eligible opps (≥ min stage, ≥ min value, open) where the flag = 0. Owner, Forecast Category, Stage, name, Next Step, Close Date, Value. These tables drive the coaching conversation; the KPI alone doesn't.
8. **Deal Review** — all open opps closing this quarter, sorted Value desc: Owner, Forecast Category, Stage, name, Next Step, Type, Deal Band, ICP Grade, flag columns.

## Tab 2 — FY Sales Forecast & Pipeline

1. **KPI scorecard**: FY Bookings (% of FY target, gap), Commit + Won, Strong BC + Commit + Won, Weighted Pipeline, Total Open Pipeline (no date bound), # Open Opps.
2. **Forecast Confidence by Quarter** — same columns as tab 1's monthly table, rows = FY quarters.
3. **Open Pipeline by Deal Band × Quarter** — Value + # opps per band, grand totals.
4. **Open Pipeline by Product × Quarter** — from line items (or an opp-level Product column); skipped with a note if no product data.
5. **Weighted Forecast — Stage × Quarter** and **Total Pipeline — Stage × Quarter** matrices.
6. **Pipeline created by month** (if Created Date present) — new pipeline value by created-month, with ICP-grade split when grades exist.

## Tab 3 — Seller Performance

1. **Forecast by Owner** — Won, # Won, Commit, # Commit, Strong BC, Remaining BC, # BC, Pipeline, # Pipeline, Total Pipeline, Weighted Pipeline. FY scope. Totals row.
2. **Current-quarter scorecard** — per rep: Attainment, Quota, Bookings, ASP, Forecast Attainment, Forecast (Won+Commit), Weighted Pipeline, Weighted Pipe Coverage, Large-deal Value, % large-deal. Quota from targets.csv; blank quota → "awaiting target", never 0.
3. **Trailing + current scorecard** — last 2 quarters' Attainment/Quota/Bookings/ASP next to current-quarter forecast metrics and pipe created. This is the "is this rep trending up or down" view.
4. **Open Pipeline — Owner × Stage** — per open stage: Value, Weighted, # opps.
5. **Open Pipeline — Owner × Deal Band** — Value + # opps per band.
6. **Deal Review** — open opps FY scope, as tab 1.

## Styling conventions (the script implements these)

Dark header band per section title; bold column headers on light fill; currency `#,##0` with the config symbol; percents `0%`; coverage `0.0"x"`; KPI value cells amber when the input target is missing; totals rows bold with a top border; freeze panes below each tab's title; column widths sized to content. Tab order: README, dashboards 1–3, then raw tabs last.
