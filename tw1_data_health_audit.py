"""
TW1 - Data Health Checkup: Part 1 audit script (read-only).

Course: MASY1-GC 2100 Advanced Business Analytics, Team session 2-4 5
Case:   TechPoint SaaS Solutions - B2B SaaS Churn Data

What this script does
---------------------
1. Loads the raw Excel extract WITHOUT modifying it.
2. Runs the four Part 1 checks:
     - Volume & Redundancy      (row count vs. IT's claim, duplicate rows / IDs)
     - Schema & Format          (cell storage types, currency formatting, casing,
                                 mixed label encodings, missing values)
     - Logical Errors           (impossible negative values, label/date conflicts)
     - Extreme Outliers         (whale revenue, 50,000 users, IQR screen)
3. Pre-computes the numbers Parts 2 and 3 need:
     - churn rate 2022 vs. 2023, whale impact on the revenue baseline,
       segment snapshot, and the Date_Cancelled leakage evidence.
4. Writes:
     - TW1_Audit_Findings.xlsx   one sheet per check, with the offending rows
     - TW1_Audit_Summary.md      plain-English summary for the memo

Usage
-----
    python tw1_data_health_audit.py [path/to/"B2B SaaS Churn Data.xlsx"] [output_dir]

Dependencies: pandas, openpyxl
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import pandas as pd

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
DEFAULT_INPUT = Path("B2B SaaS Churn Data.xlsx")
IT_CLAIMED_ROWS = 2500
NUMERIC_COLS = ["Total_Users", "Monthly_Revenue", "Support_Tickets", "Last_Login_Days_Ago"]
CATEGORICAL_COLS = ["Industry", "Contract_Type", "Churn_Label"]
DATE_COLS = ["Join_Date", "Date_Cancelled"]
# Hard thresholds for "extreme" anomalies named in the assignment brief.
WHALE_REVENUE = 100_000        # anything >= this is an obvious whale (brief cites $1,000,000)
EXTREME_USERS = 10_000         # anything >= this is an obvious anomaly (brief cites 50,000)
IQR_MULTIPLIER = 3.0           # conservative fence for the statistical screen


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def pct(n: float, d: float) -> str:
    return f"{100 * n / d:.1f}%" if d else "n/a"


def normalize_churn(v) -> str | None:
    """Map the mixed Yes/No/1/0 encodings to Yes/No for analysis only."""
    s = str(v).strip().lower()
    if s in {"yes", "y", "1", "true"}:
        return "Yes"
    if s in {"no", "n", "0", "false"}:
        return "No"
    return None


# --------------------------------------------------------------------------- #
# Load
# --------------------------------------------------------------------------- #
def load(path: Path) -> tuple[pd.DataFrame, openpyxl.worksheet.worksheet.Worksheet]:
    df = pd.read_excel(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    return df, wb.active


# --------------------------------------------------------------------------- #
# Check 1: Volume & Redundancy
# --------------------------------------------------------------------------- #
def check_volume(df: pd.DataFrame) -> tuple[dict, dict[str, pd.DataFrame]]:
    exact_dups = df[df.duplicated(keep=False)].sort_values("Customer_ID")
    id_dups = df[df.duplicated("Customer_ID", keep=False)].sort_values("Customer_ID")
    # IDs that repeat but whose rows are not byte-for-byte identical (conflicting records)
    conflicting_ids = (
        id_dups.groupby("Customer_ID")
        .filter(lambda g: len(g.drop_duplicates()) > 1)
        .sort_values("Customer_ID")
    )
    summary = {
        "rows_in_file": len(df),
        "rows_claimed_by_IT": IT_CLAIMED_ROWS,
        "row_difference": len(df) - IT_CLAIMED_ROWS,
        "exact_duplicate_rows_to_remove": int(df.duplicated().sum()),
        "duplicate_customer_id_rows_to_remove": int(df.duplicated("Customer_ID").sum()),
        "customer_ids_with_conflicting_records": int(conflicting_ids["Customer_ID"].nunique()),
        "unique_customer_ids": int(df["Customer_ID"].nunique()),
        "rows_after_dedup_on_customer_id": int(len(df.drop_duplicates("Customer_ID"))),
    }
    tables = {
        "Dup_ExactRows": exact_dups,
        "Dup_CustomerID": id_dups,
        "Dup_ConflictingID": conflicting_ids,
    }
    return summary, tables


# --------------------------------------------------------------------------- #
# Check 2: Schema & Format
# --------------------------------------------------------------------------- #
def check_schema(df: pd.DataFrame, ws) -> tuple[dict, dict[str, pd.DataFrame]]:
    headers = [c.value for c in ws[1]]
    col_idx = {h: i for i, h in enumerate(headers)}

    # 2a. Excel-level storage: is any numeric column stored as text, and how is it formatted?
    storage_rows = []
    for col in NUMERIC_COLS:
        i = col_idx[col]
        types, formats = {}, {}
        for row in ws.iter_rows(min_row=2, min_col=i + 1, max_col=i + 1):
            c = row[0]
            types[c.data_type] = types.get(c.data_type, 0) + 1
            formats[c.number_format] = formats.get(c.number_format, 0) + 1
        storage_rows.append(
            {
                "column": col,
                "pandas_dtype": str(df[col].dtype),
                "excel_cell_types": str(types),              # 'n' numeric, 's' string
                "text_stored_cells": types.get("s", 0),
                "excel_number_formats": str(formats),
                "currency_formatted_cells": sum(v for k, v in formats.items() if "$" in k),
            }
        )
    storage = pd.DataFrame(storage_rows)

    # 2b. Categorical casing / encoding inconsistencies
    cat_rows = []
    for col in CATEGORICAL_COLS:
        vc = df[col].astype(str).value_counts()
        canon = df[col].astype(str).str.strip().str.lower().value_counts()
        for raw, n in vc.items():
            cat_rows.append(
                {
                    "column": col,
                    "raw_value": raw,
                    "count": int(n),
                    "canonical_value": raw.strip().lower(),
                    "variants_sharing_canonical": int((df[col].astype(str).str.strip().str.lower() == raw.strip().lower()).sum()),
                }
            )
    cat = pd.DataFrame(cat_rows).sort_values(["column", "canonical_value", "count"], ascending=[True, True, False])

    industry_lower = df["Industry"].str.strip().str.lower()
    industry_case_issue = int((df["Industry"] != df["Industry"].str.strip().str.title()).sum())
    is_numeric_label = df["Churn_Label"].map(lambda v: normalize_churn(v) is not None and str(v).strip().lower() not in {"yes", "no"})
    churn_numeric = int(is_numeric_label.sum())
    # How trustworthy is the numeric encoding? Compare against Date_Cancelled.
    numeric_with_date = int((is_numeric_label & df["Date_Cancelled"].notna()).sum())
    yes_with_date = int(((df["Churn_Label"].astype(str).str.strip().str.lower() == "yes") & df["Date_Cancelled"].notna()).sum())
    yes_total = int((df["Churn_Label"].astype(str).str.strip().str.lower() == "yes").sum())

    # 2c. Missing values
    missing = pd.DataFrame(
        {
            "column": df.columns,
            "missing_count": df.isna().sum().values,
            "missing_pct": [pct(m, len(df)) for m in df.isna().sum().values],
        }
    )

    summary = {
        "numeric_columns_stored_as_text_in_excel": int(storage["text_stored_cells"].sum()),
        "monthly_revenue_cells_with_currency_format": int(storage.loc[storage.column == "Monthly_Revenue", "currency_formatted_cells"].iloc[0]),
        "industry_rows_with_non_title_case": industry_case_issue,
        "industry_distinct_raw_values": int(df["Industry"].nunique()),
        "industry_distinct_canonical_values": int(industry_lower.nunique()),
        "churn_label_raw_values": str(df["Churn_Label"].astype(str).value_counts().to_dict()),
        "churn_label_rows_encoded_as_number": churn_numeric,
        "numeric_label_rows_with_date_cancelled": numeric_with_date,
        "yes_label_rows_with_date_cancelled": f"{yes_with_date} of {yes_total}",
        "contract_type_unknown_rows": int((df["Contract_Type"] == "Unknown").sum()),
        "support_tickets_missing": int(df["Support_Tickets"].isna().sum()),
        "date_cancelled_missing": int(df["Date_Cancelled"].isna().sum()),
    }
    return summary, {"Schema_Storage": storage, "Schema_Categorical": cat, "Schema_Missing": missing}


# --------------------------------------------------------------------------- #
# Check 3: Logical Errors & Impossible Values
# --------------------------------------------------------------------------- #
def check_logic(df: pd.DataFrame) -> tuple[dict, dict[str, pd.DataFrame]]:
    neg_frames = []
    summary = {}
    for col in NUMERIC_COLS:
        bad = df[df[col] < 0]
        summary[f"negative_{col}"] = len(bad)
        if len(bad):
            neg_frames.append(bad.assign(issue=f"negative {col}"))
    negatives = pd.concat(neg_frames) if neg_frames else df.iloc[0:0]

    churn = df["Churn_Label"].map(normalize_churn)
    has_cancel = df["Date_Cancelled"].notna()
    conflict_a = df[(churn == "Yes") & ~has_cancel]          # churned but no cancel date
    conflict_b = df[(churn == "No") & has_cancel]            # not churned but has cancel date
    cancel_before_join = df[has_cancel & (df["Date_Cancelled"] < df["Join_Date"])]

    summary.update(
        {
            "churn_yes_without_date_cancelled": len(conflict_a),
            "churn_no_with_date_cancelled": len(conflict_b),
            "date_cancelled_before_join_date": len(cancel_before_join),
            "join_date_range": f"{df['Join_Date'].min().date()} to {df['Join_Date'].max().date()}",
            "date_cancelled_range": f"{df['Date_Cancelled'].min().date()} to {df['Date_Cancelled'].max().date()}",
        }
    )
    tables = {
        "Logic_Negatives": negatives,
        "Logic_LabelDateConflict": pd.concat([conflict_a.assign(issue="Yes without date"), conflict_b.assign(issue="No with date")]),
        "Logic_CancelBeforeJoin": cancel_before_join,
    }
    return summary, tables


# --------------------------------------------------------------------------- #
# Check 4: Extreme Outliers
# --------------------------------------------------------------------------- #
def check_outliers(df: pd.DataFrame) -> tuple[dict, dict[str, pd.DataFrame]]:
    whales = df[df["Monthly_Revenue"] >= WHALE_REVENUE]
    big_users = df[df["Total_Users"] >= EXTREME_USERS]

    iqr_rows = []
    for col in NUMERIC_COLS:
        s = df[col].dropna()
        q1, q3 = s.quantile([0.25, 0.75])
        iqr = q3 - q1
        lo, hi = q1 - IQR_MULTIPLIER * iqr, q3 + IQR_MULTIPLIER * iqr
        flagged = df[(df[col] < lo) | (df[col] > hi)]
        iqr_rows.append(
            {
                "column": col,
                "median": float(s.median()),
                "mean": float(s.mean()),
                "min": float(s.min()),
                "max": float(s.max()),
                "iqr_low_fence": float(lo),
                "iqr_high_fence": float(hi),
                "rows_outside_fence": len(flagged),
            }
        )
    iqr = pd.DataFrame(iqr_rows)

    # Revenue baseline with and without the whale(s): the Part 2 question.
    rev = df["Monthly_Revenue"]
    rev_ex = rev[rev < WHALE_REVENUE]
    rev_ex_clean = rev_ex[rev_ex >= 0]
    baseline = pd.DataFrame(
        [
            {"scenario": "All rows (raw)", "n": len(rev), "mean": rev.mean(), "median": rev.median()},
            {"scenario": f"Excluding revenue >= {WHALE_REVENUE:,}", "n": len(rev_ex), "mean": rev_ex.mean(), "median": rev_ex.median()},
            {"scenario": "Excluding whale AND negative revenue", "n": len(rev_ex_clean), "mean": rev_ex_clean.mean(), "median": rev_ex_clean.median()},
        ]
    )
    summary = {
        "whale_rows": len(whales),
        "whale_customer_ids": str(whales["Customer_ID"].tolist()),
        "extreme_user_rows": len(big_users),
        "extreme_user_customer_ids": str(big_users["Customer_ID"].tolist()),
        "revenue_mean_raw": round(rev.mean(), 2),
        "revenue_mean_ex_whale": round(rev_ex.mean(), 2),
        "revenue_median_raw": round(rev.median(), 2),
        "whale_inflates_mean_by": round(rev.mean() - rev_ex.mean(), 2),
        "users_mean_raw": round(df["Total_Users"].mean(), 2),
        "users_mean_ex_extreme": round(df.loc[df["Total_Users"] < EXTREME_USERS, "Total_Users"].mean(), 2),
        "users_median_raw": round(df["Total_Users"].median(), 2),
    }
    tables = {
        "Outlier_Whales": whales,
        "Outlier_Users": big_users,
        "Outlier_IQR_Screen": iqr,
        "Outlier_RevenueBaseline": baseline,
    }
    return summary, tables


# --------------------------------------------------------------------------- #
# Check 5: Cross-field consistency (sanity checks the single-column tests miss)
# --------------------------------------------------------------------------- #
REV_PER_USER_LOW, REV_PER_USER_HIGH = 20, 150      # plausible $/user/month band


def check_consistency(df: pd.DataFrame) -> tuple[dict, dict[str, pd.DataFrame]]:
    churn = df["Churn_Label"].map(normalize_churn)

    # 5a. Revenue per user: the dataset prices at roughly $50 per user per month,
    #     so a ratio 1,000x off identifies which field of an outlier row is wrong.
    valid = (df["Total_Users"] > 0) & (df["Monthly_Revenue"] > 0)
    rpu = (df["Monthly_Revenue"] / df["Total_Users"]).where(valid)
    rpu_typ = rpu[(df["Monthly_Revenue"] < WHALE_REVENUE) & (df["Total_Users"] < EXTREME_USERS)]
    rpu_bad = df.assign(Revenue_per_User=rpu.round(2))[(rpu < REV_PER_USER_LOW) | (rpu > REV_PER_USER_HIGH)]

    # 5b. Engagement field vs. cancellation: a customer that cancelled in Jan 2022
    #     cannot have logged in within the last 30 days of an extract that runs to
    #     Dec 2023. Last_Login_Days_Ago is therefore not anchored to one snapshot date.
    extract_date = df["Date_Cancelled"].max()
    days_since_cancel = (extract_date - df["Date_Cancelled"]).dt.days
    login_after_cancel = df[(days_since_cancel > df["Last_Login_Days_Ago"].abs()) & df["Date_Cancelled"].notna()]

    # 5c. Missingness pattern for Support_Tickets: random or concentrated?
    miss = df["Support_Tickets"].isna()
    miss_by_churn = pd.crosstab(churn.fillna("Unparsed"), miss, normalize="index").rename(columns={True: "missing_share", False: "present_share"}).round(3).reset_index()
    miss_by_industry = pd.crosstab(df["Industry"].str.strip().str.title(), miss, normalize="index").rename(columns={True: "missing_share", False: "present_share"}).round(3).reset_index()

    # 5d. Cancellation history window
    cancel_years = df["Date_Cancelled"].dt.year.dropna().astype(int).value_counts().sort_index()
    join_years = df["Join_Date"].dt.year.value_counts().sort_index()

    summary = {
        "revenue_per_user_typical_median": round(float(rpu_typ.median()), 2),
        "revenue_per_user_typical_range": f"{rpu_typ.min():.1f} to {rpu_typ.max():.1f}",
        "rows_outside_revenue_per_user_band": len(rpu_bad),
        "whale_revenue_per_user": ", ".join(f"{x:,.0f}" for x in rpu[df["Monthly_Revenue"] >= WHALE_REVENUE]),
        "extreme_users_revenue_per_user": ", ".join(f"{x:,.2f}" for x in rpu[df["Total_Users"] >= EXTREME_USERS]),
        "implied_extract_date": str(extract_date.date()),
        "churned_rows_with_login_after_cancellation": len(login_after_cancel),
        "churned_rows_total": int(df["Date_Cancelled"].notna().sum()),
        "support_missing_share_churned_vs_retained": f"{miss[churn == 'Yes'].mean():.1%} vs {miss[churn == 'No'].mean():.1%}",
        "support_missing_share_by_industry_range": f"{miss.groupby(df['Industry'].str.strip().str.title()).mean().min():.1%} to {miss.groupby(df['Industry'].str.strip().str.title()).mean().max():.1%}",
        "join_years": str(join_years.to_dict()),
        "cancellation_years": str(cancel_years.to_dict()),
        "company_name_is_derived_from_id": bool((df["Company_Name"] == "Company_" + (df["Customer_ID"] - df["Customer_ID"].min()).astype(str)).all()),
    }
    tables = {
        "Consist_RevPerUser": rpu_bad,
        "Consist_LoginAfterCancel": login_after_cancel.assign(days_since_cancel=days_since_cancel[login_after_cancel.index]),
        "Consist_SupportMissing": pd.concat([miss_by_churn.rename(columns={"Churn_Label": "group"}), miss_by_industry.rename(columns={"Industry": "group"})]),
    }
    return summary, tables


# --------------------------------------------------------------------------- #
# Part 2 / Part 3 support: churn trend, segments, leakage evidence
# --------------------------------------------------------------------------- #
def part2_part3_support(df: pd.DataFrame) -> tuple[dict, dict[str, pd.DataFrame]]:
    d = df.drop_duplicates("Customer_ID").copy()      # light dedup so counts are per customer
    d["churn"] = d["Churn_Label"].map(normalize_churn)
    d["industry"] = d["Industry"].str.strip().str.title()
    d["cancel_year"] = d["Date_Cancelled"].dt.year

    # Churn spike: churn events by year of cancellation, against the customer base
    # that had already joined by the start of that year (all joins are 2020-2022).
    rows = []
    for yr in (2022, 2023):
        churned = int((d["cancel_year"] == yr).sum())
        at_risk = int((d["Join_Date"] < pd.Timestamp(yr, 1, 1)).sum())
        # Customers already cancelled before the year begins are not at risk.
        already_gone = int((d["cancel_year"] < yr).sum())
        base = at_risk - already_gone
        rows.append({"year": yr, "churn_events": churned, "customers_at_start_of_year": base, "churn_rate": churned / base if base else None})
    churn_year = pd.DataFrame(rows)
    churn_year["churn_rate_pct"] = (churn_year["churn_rate"] * 100).round(1)

    overall_churn = d["churn"].value_counts()
    by_industry = (
        d.groupby("industry")["churn"]
        .agg(customers="size", churned=lambda s: (s == "Yes").sum())
        .assign(churn_rate_pct=lambda t: (100 * t.churned / t.customers).round(1))
        .sort_values("customers", ascending=False)
    )
    by_contract = (
        d.groupby("Contract_Type")["churn"]
        .agg(customers="size", churned=lambda s: (s == "Yes").sum())
        .assign(churn_rate_pct=lambda t: (100 * t.churned / t.customers).round(1))
        .sort_values("churn_rate_pct", ascending=False)
    )
    by_ind_contract = (
        d.pivot_table(index="industry", columns="Contract_Type", values="churn", aggfunc=lambda s: round(100 * (s == "Yes").mean(), 1))
    )

    # Leakage evidence: Date_Cancelled is populated if and only if the customer churned.
    leak = pd.crosstab(d["churn"].fillna("Unparsed"), d["Date_Cancelled"].notna().map({True: "Date_Cancelled present", False: "Date_Cancelled blank"}))

    summary = {
        "churn_2022_events": int(churn_year.loc[0, "churn_events"]),
        "churn_2022_rate_pct": float(churn_year.loc[0, "churn_rate_pct"]),
        "churn_2023_events": int(churn_year.loc[1, "churn_events"]),
        "churn_2023_rate_pct": float(churn_year.loc[1, "churn_rate_pct"]),
        "overall_churn_rate_pct": round(100 * (d["churn"] == "Yes").mean(), 1),
        "unique_customers": len(d),
        "churn_2022_share_all_pct": round(100 * churn_year.loc[0, "churn_events"] / len(d), 1),
        "churn_2023_share_all_pct": round(100 * churn_year.loc[1, "churn_events"] / len(d), 1),
        "largest_industry": by_industry.index[0],
        "largest_industry_share_pct": round(100 * by_industry["customers"].iloc[0] / len(d), 1),
        "highest_churn_industry": by_industry["churn_rate_pct"].idxmax(),
        "highest_churn_industry_rate_pct": float(by_industry["churn_rate_pct"].max()),
        "highest_churn_contract": by_contract.index[0],
        "highest_churn_contract_rate_pct": float(by_contract["churn_rate_pct"].iloc[0]),
        "lowest_churn_contract": by_contract.index[-1],
        "lowest_churn_contract_rate_pct": float(by_contract["churn_rate_pct"].iloc[-1]),
        "leakage_churn_yes_with_date": int(((d["churn"] == "Yes") & d["Date_Cancelled"].notna()).sum()),
        "leakage_churn_no_with_date": int(((d["churn"] == "No") & d["Date_Cancelled"].notna()).sum()),
    }
    tables = {
        "P2_ChurnByYear": churn_year,
        "P2_ChurnByIndustry": by_industry.reset_index(),
        "P2_ChurnByContract": by_contract.reset_index(),
        "P2_Industry_x_Contract": by_ind_contract.reset_index(),
        "P3_LeakageCrosstab": leak.reset_index(),
    }
    return summary, tables


# --------------------------------------------------------------------------- #
# Reporting
# --------------------------------------------------------------------------- #
def write_outputs(out_dir: Path, summaries: dict[str, dict], tables: dict[str, pd.DataFrame], df: pd.DataFrame) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)

    # Excel workbook: summary first, then one sheet per finding table
    xlsx = out_dir / "TW1_Audit_Findings.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        flat = [{"section": sec, "metric": k, "value": v} for sec, s in summaries.items() for k, v in s.items()]
        pd.DataFrame(flat).to_excel(xw, sheet_name="Summary", index=False)
        for name, t in tables.items():
            t.to_excel(xw, sheet_name=name[:31], index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)

    # Markdown summary
    v, s, l, o, c, p = (summaries[k] for k in ("volume", "schema", "logic", "outliers", "consistency", "part2_3"))
    md = f"""# TW1 Data Health Checkup - Audit Summary (auto-generated)

Source file: B2B SaaS Churn Data.xlsx | Generated by tw1_data_health_audit.py
All figures below are computed from the raw extract; nothing has been cleaned or removed.

## Part 1 - High-Level Data Health Audit

### 1. Volume & Redundancy
- IT claimed **{v['rows_claimed_by_IT']:,}** records; the file contains **{v['rows_in_file']:,}** rows ({v['row_difference']:+d}).
- **{v['exact_duplicate_rows_to_remove']}** rows are exact duplicates of another row.
- **{v['duplicate_customer_id_rows_to_remove']}** rows repeat an existing Customer_ID; {v['customer_ids_with_conflicting_records']} Customer_ID(s) appear with conflicting field values (same ID, different data).
- Unique customers: **{v['unique_customer_ids']:,}**. After de-duplicating on Customer_ID the table holds {v['rows_after_dedup_on_customer_id']:,} rows.

### 2. Schema & Format Inconsistencies
- Numeric columns stored as text in Excel: **{s['numeric_columns_stored_as_text_in_excel']}** cells. Monthly_Revenue is numeric, but only **{s['monthly_revenue_cells_with_currency_format']}** cells carry a `$` currency display format while the rest are plain numbers, so the column looks mixed in Excel and would break if anyone re-typed it as text.
- Industry: **{s['industry_distinct_raw_values']}** raw spellings collapse to **{s['industry_distinct_canonical_values']}** real industries; **{s['industry_rows_with_non_title_case']}** rows use lower-case variants (e.g. `retail` vs `Retail`). Any pivot or model treats these as separate segments.
- Churn_Label mixes text and numbers: {s['churn_label_raw_values']} - **{s['churn_label_rows_encoded_as_number']}** rows are encoded as a number instead of Yes/No. The numeric code is NOT a safe synonym for "Yes": every `Yes` row has a Date_Cancelled ({s['yes_label_rows_with_date_cancelled']}), but only **{s['numeric_label_rows_with_date_cancelled']}** of the {s['churn_label_rows_encoded_as_number']} numeric rows do. These rows need to be resolved against Date_Cancelled (or confirmed with IT), not blindly recoded.
- Contract_Type contains **{s['contract_type_unknown_rows']}** rows labelled `Unknown`.
- Missing values: Support_Tickets is blank in **{s['support_tickets_missing']}** rows; Date_Cancelled is blank in {s['date_cancelled_missing']:,} rows (expected - blank means the customer has not cancelled).

### 3. Logical Errors & Impossible Values
- Negative Total_Users: **{l['negative_Total_Users']}** row(s); negative Monthly_Revenue: **{l['negative_Monthly_Revenue']}** row(s); negative Last_Login_Days_Ago: **{l['negative_Last_Login_Days_Ago']}** rows; negative Support_Tickets: {l['negative_Support_Tickets']}.
- Churn label vs. cancel date: {l['churn_yes_without_date_cancelled']} churned rows without a Date_Cancelled; {l['churn_no_with_date_cancelled']} non-churned rows with a Date_Cancelled; {l['date_cancelled_before_join_date']} rows cancelled before they joined.
- Join_Date spans {l['join_date_range']}; Date_Cancelled spans {l['date_cancelled_range']}.

### 4. Extreme Outliers
- Monthly_Revenue whale(s) at or above ${WHALE_REVENUE:,}: **{o['whale_rows']}** row(s) (Customer_ID {o['whale_customer_ids']}). Raw mean = **${o['revenue_mean_raw']:,.2f}** vs. **${o['revenue_mean_ex_whale']:,.2f}** without the whale (median ${o['revenue_median_raw']:,.2f}). The single record inflates the average by about ${o['whale_inflates_mean_by']:,.2f}.
- Total_Users at or above {EXTREME_USERS:,}: **{o['extreme_user_rows']}** row(s) (Customer_ID {o['extreme_user_customer_ids']}). Raw mean = {o['users_mean_raw']} users vs. {o['users_mean_ex_extreme']} without it (median {o['users_median_raw']}).
- See sheet `Outlier_IQR_Screen` for a 3x IQR fence on every numeric column.

### 5. Cross-field consistency (beyond the four required checks)
- Pricing sanity: typical accounts pay about **${c['revenue_per_user_typical_median']} per user per month** (range {c['revenue_per_user_typical_range']}). The whale row implies ${c['whale_revenue_per_user']} per user and the 50,000-user row implies ${c['extreme_users_revenue_per_user']} per user, i.e. both are roughly 1,000x off the pricing pattern. This confirms they are data-entry errors (an extra "000"), not real strategic accounts, and identifies which field is wrong in each row. {c['rows_outside_revenue_per_user_band']} rows fall outside the ${REV_PER_USER_LOW}-${REV_PER_USER_HIGH} per-user band (sheet `Consist_RevPerUser`).
- Engagement field is not time-anchored: the extract runs to {c['implied_extract_date']}, yet **{c['churned_rows_with_login_after_cancellation']} of {c['churned_rows_total']}** cancelled customers show a Last_Login_Days_Ago that would place their last login AFTER they cancelled (every value is 0-30 days regardless of when the account left). Last_Login_Days_Ago was evidently measured at a different point in time per customer, or generated without reference to cancellation, so it cannot be used as-is as an early-warning signal.
- Support_Tickets blanks look random rather than systematic: missing share is {c['support_missing_share_churned_vs_retained']} for churned vs. retained customers and {c['support_missing_share_by_industry_range']} across industries. This supports simple imputation plus a missing flag instead of dropping rows.
- History window: customers joined in {c['join_years']} but cancellations exist only for {c['cancellation_years']}. Either churn before 2022 was purged from the extract (survivorship bias in the 2022 base) or the file only tracks two years of exits; confirm with IT before quoting a multi-year trend.
- Company_Name is derived from Customer_ID (`Company_<ID-1001>`, derived for all rows = {c['company_name_is_derived_from_id']}); it carries no information and should be treated as an identifier, not a feature.

## Part 2 support - Descriptive & Diagnostic numbers (per unique customer)
- Churn events by cancellation year: **2022 = {p['churn_2022_events']}** ({p['churn_2022_rate_pct']}% of customers active at the start of 2022) vs. **2023 = {p['churn_2023_events']}** ({p['churn_2023_rate_pct']}% of customers active at the start of 2023). Churn events roughly {p['churn_2023_events'] / p['churn_2022_events']:.1f}x year over year; the data supports the VP's concern. Overall churn = {p['overall_churn_rate_pct']}% of unique customers.
- Simple Excel-style view (share of all {p['unique_customers']:,} customers): 2022 = {p['churn_2022_share_all_pct']}%, 2023 = {p['churn_2023_share_all_pct']}%.
- Largest segment: **{p['largest_industry']}** ({p['largest_industry_share_pct']}% of customers). Highest churn industry: {p['highest_churn_industry']} ({p['highest_churn_industry_rate_pct']}%).
- Contract type: **{p['highest_churn_contract']}** churns most ({p['highest_churn_contract_rate_pct']}%), **{p['lowest_churn_contract']}** least ({p['lowest_churn_contract_rate_pct']}%). Full tables in sheets `P2_*`.
- Note: churn rates by year use the cleaned Yes/No label and de-duplicated customers; the "rate" denominator is customers who had joined before Jan 1 of that year and had not yet cancelled.

## Part 3 support - Data Leakage evidence
- Date_Cancelled is populated for **{p['leakage_churn_yes_with_date']}** churned customers and for **{p['leakage_churn_no_with_date']}** retained customers: it is a one-to-one proxy for Churn_Label and is only known AFTER the customer has left. It must be dropped (or used only to build the label / churn-year) before modelling. See sheet `P3_LeakageCrosstab`.
- Last_Login_Days_Ago is measured at extract time, not at the decision point, so it is a softer leakage risk and should be time-anchored in a production pipeline.

## Files
- `TW1_Audit_Findings.xlsx` - Summary sheet plus every offending row for each check.
- `tw1_data_health_audit.py` - this audit (read-only).
- `tw1_data_cleaning.py` - remediation script, disabled by default (dry-run) until the team agrees on Part 4.
"""
    (out_dir / "TW1_Audit_Summary.md").write_text(md, encoding="utf-8")
    print(md)
    print(f"\nWritten: {xlsx}\nWritten: {out_dir / 'TW1_Audit_Summary.md'}")


def main() -> None:
    inp = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_INPUT
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else inp.parent
    df, ws = load(inp)

    summaries, tables = {}, {}
    for key, fn in (("volume", check_volume), ("schema", check_schema), ("logic", check_logic), ("outliers", check_outliers), ("consistency", check_consistency), ("part2_3", part2_part3_support)):
        s, t = fn(df, ws) if key == "schema" else fn(df)
        summaries[key] = s
        tables.update(t)
    write_outputs(out, summaries, tables, df)


if __name__ == "__main__":
    main()
