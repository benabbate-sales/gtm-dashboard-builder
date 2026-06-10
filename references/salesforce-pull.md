# Getting the data — Salesforce MCP or CSV fallback

The build script consumes flat CSVs. This file covers how to produce them from a live Salesforce connection, and the export spec when there's no connector.

## Path A — Salesforce MCP (or similar CRM connector)

### 1. Verify fields before querying

Custom field API names differ per org. Querying a wrong-but-valid field returns nulls, and null-driven dashboards look plausible while being wrong. So, first:

- Use the connector's describe/metadata tool (often `describe_object`, `sobject_describe`, or a `runSoql` on `FieldDefinition`) on **Opportunity**.
- Confirm the API names for: the value field (`Amount` vs `ARR__c` vs similar), forecast category (`ForecastCategory` / `ForecastCategoryName` vs a custom picklist), any include-in-forecast flag, ICP/account grade, deal-quality checkboxes.
- Confirm the actual `StageName` picklist values and, if available, their default probabilities (`OpportunityStage` object: `SELECT MasterLabel, DefaultProbability, IsClosed, IsWon FROM OpportunityStage WHERE IsActive = true`). Use these to pre-fill the config, then confirm with the client.

If the connector has no describe tool, query one record with the candidate fields and inspect what comes back; a `MALFORMED_QUERY`/invalid-field error is your verification signal.

### 2. Pull opportunities

Template — replace bracketed custom fields with the verified names, or delete the line if the org doesn't have the concept:

```sql
SELECT Id, Name, Account.Name, Owner.Name, StageName, Type,
       CloseDate, CreatedDate, NextStep, LeadSource,
       ForecastCategoryName,
       Amount,                      -- or [ARR__c]
       [Include_Forecast__c],
       [ICP_Grade__c],
       [Region__c],
       [Loss_Reason__c],
       [ROI_Analysis__c], [Value_Map__c]
FROM Opportunity
WHERE CloseDate >= [FY_START] AND CloseDate <= [FY_END_PLUS_LOOKAHEAD]
```

Pull the whole fiscal year plus any open deals beyond it (a second query with `IsClosed = false AND CloseDate > [FY_END]` if you want the unbounded-pipeline KPI). Most MCP query tools cap rows (~2,000); page with `ORDER BY Id` + `WHERE Id > last_id` if needed.

### 3. Pull line items (only if doing product-level views)

```sql
SELECT OpportunityId, Product2.Name, Product2.Family, TotalPrice, Quantity
FROM OpportunityLineItem
WHERE Opportunity.CloseDate >= [FY_START]
```

Skip this entirely if the org doesn't use Products — the script drops the by-product table and says so.

### 4. Write the CSVs

Map the query results to the column contracts below. Build `Opp Link` as `<instance_url>/<Id>` (get the instance URL from the connector's auth info or ask).

## Path B — CSV fallback (no connector)

Ask the client to run a Salesforce report (or any CRM export) with the columns below and send the CSV(s). For Salesforce: Reports → New Report → Opportunities → add columns → Export → Details Only, CSV. Rename headers to match the contract (or rename in the file after).

## CSV column contracts

### opportunities.csv (required)

| Column | Required | Notes |
|---|---|---|
| Opp Id | yes | any unique id |
| Opportunity Name | yes | |
| Account Name | no | |
| Opp Owner | yes | rep full name — rollup dimension |
| Stage | yes | must match config stage / won / lost names exactly |
| Type | no | New Business / Cross-Sell / Upsell / Renewal etc. |
| Forecast Category | yes | must match config `forecast_categories` values |
| Include Forecast | no | 0/1 — only if `use_include_forecast_flag` is true |
| Value | yes | numeric deal value (ARR or Amount — per config `value_label`) |
| Close Date | yes | ISO or unambiguous date |
| Created Date | no | enables pipe-created metrics |
| ICP Grade | no | must match config `icp_grades` values |
| Region | no | |
| Next Step | no | shown on deal-review tables |
| Loss Reason | no | moved-to-lost table |
| Lead Source | no | moved-to-lost table |
| Opp Link | no | hyperlink to the CRM record |
| + one column per configured deal-quality flag | no | 0/1, header = the flag's `column` |

### line_items.csv (optional)

Opp Id, Product, Product Family (optional), Product Value, Quantity (optional).

### targets.csv (required for attainment & coverage)

| Period | Owner | Target |
|---|---|---|
| FY2026 | | 4000000 |
| FY2026 Q2 | | 1000000 |
| FY2026 Q2 | Jane Doe | 250000 |

Blank Owner = org-level AOP. Period format: `FY<year>` for the year, `FY<year> Q<n>` for quarters (fiscal year = calendar year it starts in). These numbers come from the client's plan — never from the CRM. If the client hasn't supplied them yet, build without `--targets`; attainment cells will show "awaiting target".

## Sanity checks on the pulled data

- Row count vs what the client expects ("we have ~300 open opps" vs 14 rows = wrong filter).
- Distinct Stage values ⊆ config stages + won + lost. Anything else → fix config or data before building.
- Value column: no nulls on open deals; flag zero-value open deals to the client.
- Owners: a handful of reps, not 40 variants of the same names.
