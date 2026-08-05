# Client profile — EXAMPLE, `gtm-dashboard-builder`

**Copy this to `client-profile.md` and fill it in with the client. This file is the intake, not a
set of defaults.** The skill reads the profile at run time and generates `config.json` from it; the
config is a build artefact, and if the two disagree the profile wins.

Every field here exists because a dashboard number depends on it. **A CEO dismisses the whole
workbook over one wrong stage name**, so an unconfirmed answer is worth less than an admitted gap:
mark anything unverified and it gets flagged in the workbook README rather than passed off as fact.

---

## 1. The engagement

- **Client name** (as it should appear on the workbook):
- **Currency symbol:**
- **As-of date convention** — what "today" means for a run:
- **Who reads this** — CEO, CRO, board, the sales team:

## 2. Fiscal calendar

- **Fiscal year start month:**
- **FY label the client actually uses** (their words, not a convention):
- **Are quarters calendar-aligned?**

## 3. Value

- **What the org treats as deal value** — the standard amount field, a custom ARR field, or a
  line-item rollup. **Never assume; ask or verify against the org.**
- **The label to print** (ARR, TCV, ACV, Bookings):

## 4. Stages

Open stages **in funnel order**, with the win probability the client agrees. Closed-won and
closed-lost are named separately.

| # | Stage name (exact CRM spelling) | Probability |
|---|---|---|
| 1 | | |
| 2 | | |
| 3 | | |
| 4 | | |

- **Closed-won stage name:**  · **Closed-lost stage name:**
- **Probability table confirmed by the client?** yes / no — *if no, it is flagged as assumed in the
  workbook. It drives Weighted Pipeline, so it is never quietly invented.*

## 5. Forecast categories

- **Pipeline:**  · **Best Case:**  · **Commit:**  · **Closed:**
- **Is there a separate "include in forecast" flag marking the strong best case?** yes / no — and
  the field name:

## 6. Deal-size bands

| Band name | Minimum value |
|---|---|
| | |

- **Which bands count as "large"** for the large-deal-mix metric:

## 7. Account grading

- **Does the org grade accounts?** yes / no. If no, **omit the feature — do not invent a grading
  scheme.**
- **Field name:**  · **Grade values in use:**

## 8. Targets — supplied, never derived

**These are never in opportunity data.** The client supplies them; without them the attainment cells
are left visibly awaiting targets rather than filled with a guess.

- **Company plan per quarter / per FY:**
- **Quota per rep per quarter:**
- **Received?** yes / no / partial:

## 9. Deal-quality flags (optional)

Checkbox fields the org uses for deal discipline, with the eligibility rule for each — minimum
stage, minimum value.

| Field / column | Label | Min stage | Min value |
|---|---|---|---|
| | | | |

## 10. Data source

- **Salesforce MCP connected?** yes / no. If no, the CSV export spec goes to the client.
- **Custom field API names verified against the org?** — *describe the object first; a guessed
  field name that returns nulls produces a dashboard that looks fine and is silently wrong.*
- **Known data-quality issues** — blank owners, blank stages, duplicate opps:

## 11. Unconfirmed — carried into the workbook README

Anything above that the client has not explicitly confirmed. **This list is a deliverable**, not a
private note: it becomes the "assumed — confirm" section of the README tab.

-

## 12. Open questions for the client

-
