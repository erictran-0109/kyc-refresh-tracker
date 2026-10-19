# KYC Refresh & Quality Control Tracker

An Excel workbook that simulates the daily work of a KYC Refresh / QC Analyst on a Capital Markets client book: risk-rating clients, scheduling periodic refreshes, tracking documents, escalating red flags, and reviewing QC results.

All data is synthetic. Client names are fictional, and thresholds and refresh cycles are illustrative assumptions, not any bank's actual policy.

## Business problem

Banks must re-review client KYC files on a schedule that depends on each client's risk. A refresh team needs to know, every morning:

1. Which clients are high risk, and why?
2. Which files are overdue or coming due?
3. Which documents are missing or expired, and which clients need outreach?
4. Which cases must be escalated to AML Compliance or Sanctions?
5. Where is the team making mistakes, and what should change?

## Workbook structure

| Sheet | Purpose |
|---|---|
| Client_Master | 200 institutional clients (funds, corporations, banks, pension funds, trusts) across Canada, US, UK, Europe, APAC and offshore jurisdictions |
| Change_Log | Audit trail of every data correction (field, old value, new value, date, reason) |
| Risk_Scoring | Weighted risk model and Low / Medium / High rating per client, plus escalation routing |
| Refresh_Queue | Next refresh date, days until due, queue status and outreach flag per client |
| Document_Tracker | 800 client documents with status: Missing, Expired, Expiring Soon (60 days) or OK |
| Escalation_Log | 22 escalated cases with trigger, receiving team, date, status and analyst note |
| QC_Review | 80 maker-checker (4-eye) reviews with pivot tables by analyst and error category |
| Dashboard | Summary of key metrics |

## Method

**1. Data quality check.** Before any review, checked completeness and consistency. Found 9 issues: 4 clients with a region that did not match their country, 3 missing last-review dates and 2 blank UBO fields. Region errors were corrected and logged in Change_Log.

**2. Risk scoring.**

| Factor | Points |
|---|---|
| Jurisdiction: Canada / US / UK | 0 |
| Jurisdiction: Europe / APAC | 1 |
| Jurisdiction: Offshore (Cayman, BVI, Bermuda) | 3 |
| Entity: Bank / Pension Fund | 0 |
| Entity: Corporation | 1 |
| Entity: Investment Fund / Trust | 2 |
| Complex ownership structure | 2 |
| UBO not identified (or blank) | 3 |
| PEP connection | 3 |
| Adverse media: Minor / Material | 1 / 4 |

Total 0–4 = Low, 5–7 = Medium, 8+ = High. Any sanctions potential match is High regardless of score. Missing information is scored as risk, not treated as safe.

**3. Refresh schedule.** High every 12 months, Medium every 24, Low every 36. Status is Overdue, Due in 30 days, Not Due, or Missing Date. One As_Of_Date cell drives every date calculation in the workbook.

**4. Document tracking.** Four required documents per client: Certificate of Incorporation, UBO Declaration, Authorized Signer ID and Tax Form (W-8/W-9). Documents expiring within 60 days are flagged early because institutional outreach takes weeks. Issues are rolled up by client into a Front Office outreach list.

**5. Escalation.** Sanctions potential match goes to Global Economic Sanctions. Material adverse media or a PEP connection goes to AML Compliance. Each case is logged with an analyst note covering what was found, why it matters and what happens next.

**6. QC 4-eye review.** Pivot tables summarize the QC pass/fail rate by analyst (maker) and by error category.

## Key findings

| Area | Result |
|---|---|
| Risk ratings | 15 High, 45 Medium, 140 Low |
| Refresh queue | 60 overdue, including 7 High-risk; 6 due within 30 days; 3 missing review dates |
| Documents | 110 of 800 need action (51 missing, 40 expired, 19 expiring soon) |
| Outreach | 92 clients need Front Office outreach, 10 of them High-risk |
| Escalations | 22 cases: 6 to Sanctions, 16 to AML Compliance |
| QC | 15% error rate overall; one analyst at 30% |
| Top QC error | Expired ID accepted: 7 of 12 errors (58%) |

## Recommendations

1. **Work the queue in risk order.** Clear the 7 overdue High-risk files first, then the 3 records missing review dates, then the 6 due within 30 days.
2. **Automate expiry checks.** The most common QC error was accepting an expired ID. The Document_Tracker status formula already flags expired documents, so running it before submitting a file would prevent most of these errors.
3. **Targeted coaching.** One analyst's error rate is twice the team average. Coaching on document expiry and screening evidence would address most of their errors.
4. **Model improvement.** A PEP connection adds only 3 points, so some PEP clients score Low even though they still require escalation and senior approval. A next version should set a minimum rating of Medium for any PEP.

## Limitations

- Synthetic data; real KYC files hold far more fields and document types.
- Weights, thresholds and refresh cycles are illustrative, not regulatory requirements.
- Sanctions and adverse media results are given as inputs; the workbook does not perform screening itself.

## Excel skills used

VLOOKUP, nested IF, IF with OR, COUNTIF, COUNTIFS, EDATE, date arithmetic, absolute references, cross-sheet references, data validation, filters, paste values, pivot tables (count and % of row total), audit change log.
