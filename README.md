# gtm-dashboard-builder

A Claude skill that builds CEO-ready GTM dashboards — pipeline health, sales forecast, and seller scorecards — as a multi-tab Excel workbook populated from a client's Salesforce data.

Built and maintained by [Ben Abbate](https://www.linkedin.com/in/benabbate) at Win Room Studio. One of a set of open-source skills that encode a GTM operating system for scaling fintech SaaS companies — alongside [icp-builder](https://github.com/benabbate-sales/icp-builder), [gtm-cadence-builder](https://github.com/benabbate-sales/gtm-cadence-builder) and [basho-outreach](https://github.com/benabbate-sales/basho-outreach-skill).

## What it produces

One workbook, three dashboards:

1. **Quarterly Sales Forecast** — KPI scorecard (Won / Commit+Won / Strong Best Case rollup / Weighted pipeline, each with gap-to-target), forecast confidence by month, weighted forecast by stage, closed-won and moved-to-lost detail, deal review.
2. **FY Sales Forecast & Pipeline** — full-year KPIs, forecast confidence by quarter, open pipeline by deal-size band, product and stage across quarters.
3. **Seller Performance** — per-rep forecast rollup, quarterly scorecard (attainment, ASP, forecast attainment, weighted pipe coverage, large-deal mix), trailing + current quarter view, rep × stage and rep × band pipeline, deal review.

Plus a README tab listing every assumption to confirm with the client, and raw-data tabs so each number traces to source rows.

## How it works

1. **Intake** — the skill walks through the client's CRM vocabulary: stages and win probabilities, value field (Amount vs custom ARR), forecast categories, deal-size bands, ICP grades, fiscal calendar, targets. Everything lands in a per-client `config.json` — nothing is hard-coded to any one company's CRM.
2. **Data** — pulled live through a connected Salesforce MCP (with field-verification steps and SOQL templates in `references/salesforce-pull.md`), or from a plain CSV export from any CRM.
3. **Build** — `scripts/build_dashboards.py` (pandas + openpyxl) generates the styled workbook from the CSVs and config.
4. **Verify** — the skill reconciles workbook totals against raw data before anything is delivered.

Two rules the skill enforces: targets and quotas are supplied inputs, never invented — missing targets render as "awaiting target", not zero. And no guessed Salesforce field mapping is ever presented as fact; unverified mappings are flagged on the README tab.

## Install

Download `gtm-dashboard-builder.skill` from the latest release and add it via Settings → Capabilities → Skills in the Claude app. Or copy this folder into your skills directory if you run Claude Code.

## Try it without a CRM

`assets/sample-data/` contains a synthetic dataset (120 opportunities, line items, targets, config). Build a demo workbook with:

```bash
python3 scripts/build_dashboards.py \
  --config assets/sample-data/config.json \
  --opportunities assets/sample-data/opportunities.csv \
  --line-items assets/sample-data/line_items.csv \
  --targets assets/sample-data/targets.csv \
  --as-of 2026-06-10 --output demo.xlsx
```

Requires Python 3.10+ with `pandas` and `openpyxl`.

## Repo layout

```
SKILL.md                      — the skill definition Claude follows
references/salesforce-pull.md — SOQL templates, field verification, CSV contracts
references/dashboard-specs.md — layout and metric definitions for all three dashboards
scripts/build_dashboards.py   — workbook generator
assets/config-example.json    — worked per-client config
assets/sample-data/           — synthetic demo dataset
evals/evals.json              — test prompts used to evaluate the skill
```

Questions or want this installed against your own pipeline? Open an issue or find me on LinkedIn.
