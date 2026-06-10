---
name: gtm-dashboard-builder
description: Build CEO-ready GTM dashboards — pipeline health, sales forecast, and seller scorecards — as a multi-tab Excel workbook populated from a client's Salesforce data (via a connected Salesforce MCP, or a CSV export fallback). Use whenever the user wants pipeline-health dashboards, forecast dashboards, seller/rep scorecards, attainment reporting, weighted-pipeline views, pipe-coverage analysis, or anything that sounds like "give the CEO/CRO visibility into the pipeline and how each seller is performing." Trigger on phrases like "pipeline dashboard", "forecast dashboard", "seller scorecard", "rep scorecard", "pipeline health", "build dashboards from Salesforce/our CRM", "quota attainment report" — even if the user doesn't say "dashboard". Do NOT trigger for one-off SOQL queries, CRM data cleanup, or designing the meetings where dashboards get reviewed (that's gtm-cadence-builder).
---

# GTM Dashboard Builder

Build three connected dashboards as one Excel workbook, populated with the client's real CRM data:

1. **Quarterly Sales Forecast** — KPI scorecard (Won / Commit+Won / Strong Best Case rollup / Weighted pipeline, each with gap-to-target), forecast confidence by month, weighted forecast by stage, closed-won and moved-to-lost detail, deal review.
2. **FY Sales Forecast & Pipeline** — full-year KPI scorecard, forecast confidence by quarter, open pipeline by deal-size band, by product, and by stage across quarters.
3. **Seller Performance** — per-rep forecast rollup, quarterly scorecard (attainment, ASP, forecast attainment, weighted pipe coverage, large-deal mix), trailing + current quarter scorecard, open pipeline by rep × stage and rep × deal band, deal review.

Every label — stage names, deal-size bands, account grades, products, fiscal calendar — comes from a per-client config. Nothing is hard-coded to any one company's CRM.

## Workflow

### Step 1 — Intake: learn the client's CRM vocabulary

The dashboards are only as credible as the mapping behind them. A CEO will dismiss the whole workbook over one wrong stage name. Before touching data, establish (use AskUserQuestion if the user is present; otherwise pull from engagement notes):

1. **Stages** — exact open-stage names in order, win probability per stage, and the closed-won / closed-lost stage names.
2. **Value field** — what the org treats as deal value: standard `Amount`, a custom ARR field, or a line-item rollup. Never assume. Ask, or verify against the org (Step 2).
3. **Forecast categories** — the picklist values for Pipeline / Best Case / Commit / Closed, and whether a separate "include in forecast" flag marks the strong best case.
4. **Deal-size bands** — names and value thresholds (e.g. Small / Mid / Large / Strategic), and which bands count as "large" for the large-deal-mix metric.
5. **Account/ICP grade** — does the org grade accounts (A/B/C/D or similar)? Which field? If they don't, omit it — don't invent one.
6. **Fiscal calendar** — fiscal-year start month and the FY label they use.
7. **Targets** — AOP (company plan) per quarter/FY and quota per rep per quarter. These are never in Salesforce opportunity data; the client supplies them.
8. **Deal-quality flags** (optional) — checkbox fields the org uses for deal discipline (e.g. ROI analysis done, value map built) and the eligibility rule (min stage, min value).

Record the answers in `config.json` (schema below, worked example in `assets/config-example.json`).

### Step 2 — Get the data

Two paths, in order of preference:

**A. Salesforce MCP connected** — read `references/salesforce-pull.md` and follow it. Key discipline: *describe the Opportunity object first* to verify custom field API names before querying. Guessed field names that happen to return nulls produce dashboards that look fine and are silently wrong — the worst failure mode for a tool whose whole point is forecast truth.

**B. No connector** — give the client/user the export spec in `references/salesforce-pull.md` (§ CSV fallback): a flat Opportunity report and optionally an OpportunityLineItem report. Any CRM that exports CSV works — the script only needs the column names below.

Either way, end this step with up to three CSVs in a working folder:

- `opportunities.csv` (required) — one row per opportunity
- `line_items.csv` (optional) — one row per opportunity product line
- `targets.csv` (required for attainment/coverage) — Period, Owner, Target rows supplied by the client

Exact column contracts are in `references/salesforce-pull.md`. Missing optional columns are fine — the script degrades gracefully and notes what it skipped.

### Step 3 — Write `config.json`

```json
{
  "client_name": "Client Co",
  "currency_symbol": "$",
  "fiscal_year_start_month": 1,
  "fiscal_year": 2026,
  "fy_label": "FY26",
  "value_label": "ARR",
  "stages": [
    {"name": "Discovery", "probability": 0.10},
    {"name": "Qualification", "probability": 0.25},
    {"name": "Proposal", "probability": 0.50},
    {"name": "Negotiation", "probability": 0.75}
  ],
  "won_stage": "Closed Won",
  "lost_stage": "Closed Lost",
  "forecast_categories": {
    "pipeline": "Pipeline", "best_case": "Best Case",
    "commit": "Commit", "closed": "Closed"
  },
  "use_include_forecast_flag": true,
  "deal_bands": [
    {"name": "Small", "min": 0},
    {"name": "Mid", "min": 25000},
    {"name": "Large", "min": 75000},
    {"name": "Strategic", "min": 150000}
  ],
  "large_bands": ["Large", "Strategic"],
  "icp_grades": ["A", "B", "C", "D", "Ungraded"],
  "deal_quality_flags": [
    {"column": "ROI Analysis", "label": "ROI analysis", "min_stage_index": 2, "min_value": 30000}
  ]
}
```

Notes that matter: `stages` lists *open* stages in funnel order — won/lost are separate keys. `probability` drives Weighted Pipeline; confirm the weighting table with the client rather than inventing one. `min_stage_index` is a 0-based index into `stages`. Set `icp_grades` to `[]` and `deal_quality_flags` to `[]` to drop those features cleanly. The fiscal year is labelled by the calendar year in which it starts.

### Step 4 — Build the workbook

```bash
python3 scripts/build_dashboards.py \
  --config config.json \
  --opportunities opportunities.csv \
  --targets targets.csv \
  --line-items line_items.csv \
  --as-of 2026-06-10 \
  --output "Client Co — GTM Dashboards.xlsx"
```

`--as-of` sets "today" — it picks the current fiscal quarter for tab 1 and the scorecards, and the moved-to-lost window. `--line-items` and `--targets` are optional flags; omit what you don't have. The script needs `pandas` and `openpyxl` (`pip install pandas openpyxl --break-system-packages` if missing).

The script also writes the raw data as tabs (Raw Opportunities, Raw Line Items, Targets, Config) so every dashboard number can be traced to source rows — that traceability is what makes a CEO trust the sheet.

### Step 5 — Verify, then deliver

Before handing anything over, reconcile — don't eyeball:

1. Sum of `Value` on Raw Opportunities for closed-won-in-FY equals the Bookings KPI on tab 2.
2. Per-rep Total Pipeline on Seller Performance sums to the org-level Total Pipeline.
3. Weighted pipeline ≤ unweighted pipeline everywhere.
4. No "(blank)" owners or stages — blanks mean dirty CRM data; list the offending opportunity IDs for the client instead of hiding them.
5. Every config value the client never explicitly confirmed (probabilities, band thresholds, field mappings) is flagged in the workbook README tab as "assumed — confirm".

Deliver the .xlsx, state which mappings were verified vs assumed, and end with a clear next step (e.g. "confirm the stage probabilities and I'll re-run").

If the engagement is recurring, offer to re-run the pull + build on a schedule so the dashboards stay current.

## Reference files

- `references/salesforce-pull.md` — SOQL templates, field-verification steps, MCP guidance, CSV export fallback, exact CSV column contracts. Read whenever you're getting data.
- `references/dashboard-specs.md` — the precise layout and metric definitions of all three dashboards. Read if you need to modify the script's output or explain a metric's logic.
- `assets/config-example.json` — a complete worked config.
- `assets/sample-data/` — small synthetic dataset (opportunities, line items, targets + config) for demoing the workbook before a client connects their CRM.

## Principles

- **Never present a guessed field mapping as fact.** Standard Salesforce fields (Name, StageName, Type, CloseDate, Owner, ForecastCategory, CreatedDate) are reliable; everything custom must be verified per org or flagged.
- **Targets never come from opportunity data.** AOP and quota are supplied inputs; if the client hasn't provided them, build the workbook anyway and leave the attainment cells visibly marked as awaiting targets — don't fabricate.
- **Generic output is worthless.** The deliverable must use the client's own stage names, bands, grades, and products throughout. If intake answers are missing, ask — don't fill with defaults silently.
