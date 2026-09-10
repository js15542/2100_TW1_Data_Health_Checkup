# TW1 Data Health Checkup - Remediation Plan (DRAFT, not yet enabled)

Team session 2-4 5 | MASY1-GC 2100 Advanced Business Analytics | Case: TechPoint SaaS Solutions
Status: proposal for Part 4. The matching script `tw1_data_cleaning.py` runs in dry-run mode only until the team approves this plan; then run it with `--apply`.

## 1. What the audit found (from `TW1_Audit_Summary.md`)

| # | Issue | Size | Where |
|---|-------|------|-------|
| 1 | File has 2,510 rows, IT claimed 2,500; 9 exact duplicates + 1 conflicting duplicate ID (1002: `No` vs `1`) | 10 rows | Dup_* sheets |
| 2 | Industry casing: 12 spellings for 6 industries | 251 rows (10%) | Schema_Categorical |
| 3 | Churn_Label mixes `Yes` / `No` / `1`; only 7 of the 50 `1` rows have a cancel date | 50 rows | Schema_Categorical, Logic_LabelDateConflict |
| 4 | Contract_Type = `Unknown` | 50 rows | Schema_Categorical |
| 5 | Support_Tickets blank | 125 rows (5%) | Schema_Missing |
| 6 | Negative Last_Login_Days_Ago (-5) | 50 rows | Logic_Negatives |
| 7 | Negative Total_Users (-10), negative Monthly_Revenue (-500) | 1 row each | Logic_Negatives |
| 8 | Whale: Monthly_Revenue = $1,000,000 (ID 1051); Total_Users = 50,000 (ID 1061) | 1 row each | Outlier_* |
| 9 | Monthly_Revenue: 378 cells display as `$x.xx`, the rest as plain numbers (display only, values are numeric) | 378 cells | Schema_Storage |
| 10 | Date_Cancelled is a direct proxy for churn (leakage) | 1 column | P3_LeakageCrosstab |
| 11 | Last_Login_Days_Ago is not time-anchored: 470 of 493 cancelled customers show a "last login" after their cancellation date | 470 rows | Consist_LoginAfterCancel |
| 12 | No cancellations recorded before 2022 although customers joined from 2020: pre-2022 churn history is missing or purged | window | Summary (consistency) |
| 13 | Company_Name is `Company_<ID-1001>` for every row: an identifier, not a feature | 1 column | Summary (consistency) |

Sanity check that ties issues 7 and 8 together: typical accounts pay about $50 per user per month (see `Consist_RevPerUser`). The whale implies $58,824 per user and the 50,000-user row implies $0.02 per user, so each is a single mistyped field about 1,000x off, not a real strategic account. Support_Tickets blanks are spread evenly across churn status (5.4% vs 4.9%) and industries (3.6% to 6.9%), so they can be treated as missing at random.

## 2. Top 3 critical errors and campaign impact (Part 4 draft)

1. **Churn label integrity (issues 3 + 1).** The target variable itself is unreliable for 50 rows, and the one conflicting duplicate is exactly a `No` vs `1` disagreement. If `1` is blindly recoded to `Yes`, 43 customers who show no cancellation are treated as churned: the model learns the wrong pattern, and the campaign spends retention budget on accounts that never left while the real churn rate is overstated. If the rows are dropped instead, the model loses signal. Recommended: resolve against Date_Cancelled and flag.
2. **Extreme outliers (issue 8).** One $1,000,000 record lifts average monthly revenue from about $983 to $1,381 (+40%); one 50,000-user record doubles the average user count (19 to 39). Any "high-value account" threshold, revenue-at-risk estimate, or budget allocation built on these averages is wrong, and a model will over-weight revenue and users as churn drivers.
3. **Segment fragmentation (issue 2).** Lower-case industry variants create six phantom segments of about 40 accounts each. Segment-level churn rates and the "which industry to target" decision are computed on a 10% smaller base per real industry, and a model treats `retail` and `Retail` as different customers, diluting the industry effect the campaign wants to act on.

Runner-up: Last_Login_Days_Ago as a whole (issues 6 + 11). Engagement recency is the most actionable early-warning signal for a retention campaign, but 50 rows are negative and 470 cancelled customers show logins after they left, so the field is not measured at a consistent snapshot date. Used as-is it would rank departed accounts as "recently active" and push genuine at-risk accounts down the outreach list. The fix is upstream (re-extract the field as of one snapshot date), not in this table.

## 3. Cleaning strategy (rule by rule, mirrors `tw1_data_cleaning.py`)

| Rule | Action | Rationale / alternative |
|------|--------|-------------------------|
| R1 | Drop 9 exact duplicate rows | No information lost. |
| R2 | Keep one row per Customer_ID (first occurrence), log the conflict | Alternative: ask IT which record is current. |
| R3 | Trim + Title Case Industry | Reversible, no judgement needed. |
| R4 | Churn_Label `1` -> `Yes` if Date_Cancelled present, else `No`; add `Churn_Label_Flag` | Alternative: drop the 50 rows (2%) if the team prefers zero assumptions. Confirm with IT if possible. |
| R5 | Negative values -> missing + flag column | Do not guess a corrected value (e.g. abs()); a sign error and a data-entry error look the same. The flag lets the model use "value was bad" as information. |
| R6 | Revenue >= $100k and Users >= 10k -> missing + flag; keep in a review list | Exclude from baselines and modelling until Sales confirms the records. The $/user ratio shows the whale's revenue (17 users) and the 50,000-user row's user count ($1,133 revenue) are the mistyped fields; a corrected value of about $1,000 and 23 users is plausible, but correcting requires source confirmation, so the default is to blank and flag. |
| R7 | Blank Support_Tickets -> 0 + `Support_Tickets_Missing` flag | Assumption: no row = no ticket logged. Switch `SUPPORT_TICKET_IMPUTATION = "median"` if the team believes blanks are lost records. |
| R8 | Keep Contract_Type `Unknown` as its own category + flag | 2% of rows; an explicit category avoids silent bias from dropping them. |
| R9 | Drop Date_Cancelled before modelling; keep Churn_Year for reporting only | Leakage. Also drop Company_Name (identifier). Last_Login_Days_Ago should be re-extracted as of one snapshot date; until then set `DROP_LAST_LOGIN = True` or treat the field as unreliable. |
| Display | Apply one number format to Monthly_Revenue in Excel | Cosmetic; no value change. |

Expected result after cleaning: 2,500 unique customers, 6 industries, Churn_Label strictly Yes/No (491 Yes / 2,009 No), flag columns preserved so nothing is silently lost.

## 4. Open decisions for the team

- Treat Churn_Label `1` as "resolve by date" (default) or "drop"?
- Whale and 50k-user records: exclude from the modelling table (default) or cap at a fence (e.g. 99th percentile)?
- Support_Tickets blanks: zero (default) or median?
- Last_Login_Days_Ago: keep with a caveat, or drop until IT re-extracts it at a single snapshot date? (470 cancelled customers show logins after cancellation.)
- Churn trend: report 2022 vs 2023 as-is, or caveat that no pre-2022 cancellations exist in the extract?

## 5. Draft inputs for the remaining memo sections (for teammates to edit)

**Data Quality Grade (draft): C.** The extract is usable and the core fields are numeric and mostly complete, but the target label is inconsistent, 10% of a key segment field is mis-cased, single-row anomalies distort every average, and the main engagement field is not measured at a consistent point in time; nothing can be modelled until the fixes above are applied and one field is re-extracted.

**Churn observation:** cancellations rose from 145 in 2022 to 346 in 2023 (about 2.4x; 8.8% -> 14.7% of the active base), so the data supports the VP's concern, with the caveat that the extract contains no cancellations before 2022, so the 2022 base may already exclude earlier churners. Month-to-Month contracts churn at about 34% versus about 15% for 2-Year contracts, and Education shows the highest industry churn (about 24%).

**Data augmentation ideas (Part 5, Data Proximity Framework):**
- Zero-party: onboarding survey on intended use case and success criteria; quarterly NPS / renewal-intent question.
- First-party: product telemetry (weekly active users, feature adoption depth, seats used vs. purchased), support ticket sentiment and time-to-resolution, billing events (late payments, downgrades).
- Third-party: firmographic growth signals (headcount change, funding rounds, layoffs), industry-level macro indicators, technographic data showing adoption of a competing tool.
