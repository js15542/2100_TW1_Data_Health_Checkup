"""
TW1 - Data Health Checkup: Raw vs. Clean metrics for Parts 2-5.

Course: MASY1-GC 2100 Advanced Business Analytics, Team session 2-4 5
Case:   TechPoint SaaS Solutions - B2B SaaS Churn Data

Purpose
-------
Give teammates every number Parts 2-5 can cite WITHOUT re-running anything,
and show each number under two definitions side by side:

  RAW   = the extract exactly as delivered (2,510 rows). "Churned" means the
          literal text "Yes" in Churn_Label; the 50 rows coded "1" are counted
          separately, never silently merged.
  CLEAN = the table after the proposed cleaning rules R1-R8 in
          tw1_data_cleaning.py (applied in memory, nothing written):
          duplicates removed, Industry title-cased, "1" resolved against
          Date_Cancelled, negatives and the two extreme outliers set to missing,
          Support_Tickets imputed, Contract_Type "Unknown" kept.
          Date_Cancelled is kept here (rule R9 is for the modelling table only)
          so churn-by-year can be computed on the cleaned customers.

Outputs (next to the input file unless --out is given)
-------
  TW1_Raw_vs_Clean_Metrics.xlsx
      Metrics              headline numbers, one row per metric, Raw | Clean | note
      Churn_By_Year        2022 vs 2023 cancellations and rates
      Churn_By_Industry    customers, churned, churn rate, revenue per industry
      Churn_By_Contract    same by contract type
      Industry_x_Contract  churn rate matrix
      Churn_By_JoinYear    churn by customer cohort
      Revenue_Baseline     mean/median/quantiles under each treatment
      Users_Baseline       same for Total_Users
      Engagement           Support_Tickets and Last_Login_Days_Ago by churn status
      Label_Reconciliation Raw label counts -> Clean 491 / 2,009, step by step
      Audit_Issue_Counts   every Part 1 issue with count and Customer_IDs (for Part 4)
      Whale_Impact         mean with / without the two extreme rows, and the Clean mean
      Leakage_Diagnostic   label x Date_Cancelled crosstab, login-after-cancellation counts
      Definitions          exact rule text for every column
  TW1_Raw_vs_Clean_Metrics.md
      the same tables rendered as Markdown for pasting into the memo

Usage
-----
    python tw1_raw_vs_clean_metrics.py "B2B SaaS Churn Data.xlsx" [--out DIR]

Dependencies: pandas, openpyxl; imports clean_frame from tw1_data_cleaning.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from tw1_data_cleaning import RULES, clean_frame  # noqa: E402

YEARS = (2022, 2023)
WHALE_REVENUE = 100_000
EXTREME_USERS = 10_000


# --------------------------------------------------------------------------- #
# Views
# --------------------------------------------------------------------------- #
def raw_view(df: pd.DataFrame) -> pd.DataFrame:
    """Raw extract with helper columns; no values changed."""
    d = df.copy()
    lbl = d["Churn_Label"].astype(str).str.strip()
    d["churned"] = lbl.str.lower() == "yes"                 # literal "Yes" only
    d["label_ambiguous"] = ~lbl.str.lower().isin(["yes", "no"])
    d["industry"] = d["Industry"]                           # as delivered (mixed case)
    d["contract"] = d["Contract_Type"]
    return d


def clean_view(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Cleaned table (rules R1-R8) with helper columns; returns (view, change_log)."""
    c, log, _ = clean_frame(df, drop_leakage=False)
    c["churned"] = c["Churn_Label"] == "Yes"
    c["label_ambiguous"] = False
    c["industry"] = c["Industry"]
    c["contract"] = c["Contract_Type"]
    return c, log.frame()


# --------------------------------------------------------------------------- #
# Metric builders (each returns a DataFrame with a 'view' column)
# --------------------------------------------------------------------------- #
def churn_by_year(d: pd.DataFrame) -> pd.DataFrame:
    rows = []
    cancel_year = d["Date_Cancelled"].dt.year
    for yr in YEARS:
        start = pd.Timestamp(yr, 1, 1)
        base = int(((d["Join_Date"] < start) & ~(cancel_year < yr)).sum())
        events = int((cancel_year == yr).sum())
        rows.append(
            {
                "year": yr,
                "cancellations": events,
                "customers_active_at_start_of_year": base,
                "churn_rate_pct_of_active_base": round(100 * events / base, 1) if base else None,
                "cancellations_pct_of_all_rows": round(100 * events / len(d), 1),
            }
        )
    out = pd.DataFrame(rows)
    out["yoy_change_in_cancellations"] = [None, f"{out.loc[1, 'cancellations'] / out.loc[0, 'cancellations']:.2f}x"]
    return out


def churn_by_group(d: pd.DataFrame, col: str, name: str) -> pd.DataFrame:
    g = d.groupby(col, dropna=False)
    out = pd.DataFrame(
        {
            "customers": g.size(),
            "share_of_customers_pct": (100 * g.size() / len(d)).round(1),
            "churned_yes": g["churned"].sum().astype(int),
            "churn_rate_pct": (100 * g["churned"].mean()).round(1),
            "ambiguous_label_rows": g["label_ambiguous"].sum().astype(int),
            "mean_monthly_revenue": g["Monthly_Revenue"].mean().round(0),
            "median_monthly_revenue": g["Monthly_Revenue"].median().round(1),
            "mean_total_users": g["Total_Users"].mean().round(1),
        }
    )
    out.index.name = name
    return out.sort_values("customers", ascending=False).reset_index()


def industry_x_contract(d: pd.DataFrame) -> pd.DataFrame:
    """One block of rows per measure: churn_rate_pct, customers, churned_yes."""
    blocks = []
    for measure, func in (("churn_rate_pct", lambda s: round(100 * s.mean(), 1)), ("customers", "size"), ("churned_yes", "sum")):
        m = d.pivot_table(index="industry", columns="contract", values="churned", aggfunc=func)
        m.index.name = "industry"
        blocks.append(m.reset_index().assign(measure=measure))
    out = pd.concat(blocks, ignore_index=True)
    return out[["measure"] + [c for c in out.columns if c != "measure"]]


def churn_by_join_year(d: pd.DataFrame) -> pd.DataFrame:
    g = d.groupby(d["Join_Date"].dt.year)
    out = pd.DataFrame({"customers": g.size(), "churned_yes": g["churned"].sum().astype(int), "churn_rate_pct": (100 * g["churned"].mean()).round(1)})
    out.index.name = "join_year"
    return out.reset_index()


def numeric_baseline(d: pd.DataFrame, col: str) -> pd.DataFrame:
    s = d[col]
    return pd.DataFrame(
        [
            {
                "n_non_missing": int(s.notna().sum()),
                "n_missing": int(s.isna().sum()),
                "mean": round(s.mean(), 2),
                "median": round(s.median(), 2),
                "std": round(s.std(), 2),
                "min": s.min(),
                "p25": s.quantile(0.25),
                "p75": s.quantile(0.75),
                "p99": round(s.quantile(0.99), 2),
                "max": s.max(),
                "churned_mean": round(s[d["churned"]].mean(), 2),
                "retained_mean": round(s[~d["churned"] & ~d["label_ambiguous"]].mean(), 2),   # literal "No" only
                "ambiguous_mean": round(s[d["label_ambiguous"]].mean(), 2) if d["label_ambiguous"].any() else None,
            }
        ]
    )


def engagement(d: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for status, mask in (("churned (Yes)", d["churned"]), ("retained (No)", ~d["churned"] & ~d["label_ambiguous"]), ("ambiguous (1)", d["label_ambiguous"])):
        sub = d[mask]
        if len(sub) == 0:
            continue
        rows.append(
            {
                "status": status,
                "customers": len(sub),
                "support_tickets_mean": round(sub["Support_Tickets"].mean(), 2),
                "support_tickets_missing": int(sub["Support_Tickets"].isna().sum()),
                "last_login_days_mean": round(sub["Last_Login_Days_Ago"].mean(), 2),
                "last_login_days_median": sub["Last_Login_Days_Ago"].median(),
                "last_login_negative_rows": int((sub["Last_Login_Days_Ago"] < 0).sum()),
            }
        )
    return pd.DataFrame(rows)


def headline(d: pd.DataFrame) -> dict:
    cy = churn_by_year(d)
    tenure = ((d["Date_Cancelled"] - d["Join_Date"]).dt.days / 30.44)[d["Date_Cancelled"].notna()]
    ind = churn_by_group(d, "industry", "industry")
    con = churn_by_group(d, "contract", "contract")
    return {
        "rows": len(d),
        "unique_customer_ids": int(d["Customer_ID"].nunique()),
        "churned_yes": int(d["churned"].sum()),
        "ambiguous_label_rows": int(d["label_ambiguous"].sum()),
        "overall_churn_rate_pct": round(100 * d["churned"].mean(), 1),
        "rows_with_date_cancelled": int(d["Date_Cancelled"].notna().sum()),
        "cancellations_2022": int(cy.loc[0, "cancellations"]),
        "cancellations_2023": int(cy.loc[1, "cancellations"]),
        "churn_rate_2022_pct_of_active_base": cy.loc[0, "churn_rate_pct_of_active_base"],
        "churn_rate_2023_pct_of_active_base": cy.loc[1, "churn_rate_pct_of_active_base"],
        "yoy_change_in_cancellations": cy.loc[1, "yoy_change_in_cancellations"],
        "median_tenure_at_cancellation_months": round(tenure.median(), 1),
        "mean_monthly_revenue": round(d["Monthly_Revenue"].mean(), 2),
        "median_monthly_revenue": round(d["Monthly_Revenue"].median(), 2),
        "mean_total_users": round(d["Total_Users"].mean(), 2),
        "median_total_users": round(d["Total_Users"].median(), 2),
        "mean_support_tickets": round(d["Support_Tickets"].mean(), 2),
        "support_tickets_missing": int(d["Support_Tickets"].isna().sum()),
        "mean_last_login_days_ago": round(d["Last_Login_Days_Ago"].mean(), 2),
        "distinct_industry_values": int(d["industry"].nunique()),
        "largest_industry": f"{ind.loc[0, 'industry']} ({ind.loc[0, 'customers']}, {ind.loc[0, 'share_of_customers_pct']}%)",
        "highest_churn_industry": f"{ind.sort_values('churn_rate_pct').iloc[-1]['industry']} ({ind['churn_rate_pct'].max()}%)",
        "lowest_churn_industry": f"{ind.sort_values('churn_rate_pct').iloc[0]['industry']} ({ind['churn_rate_pct'].min()}%)",
        "highest_churn_contract": f"{con.sort_values('churn_rate_pct').iloc[-1]['contract']} ({con['churn_rate_pct'].max()}%)",
        "lowest_churn_contract": f"{con.sort_values('churn_rate_pct').iloc[0]['contract']} ({con['churn_rate_pct'].min()}%)",
    }


def label_reconciliation(df: pd.DataFrame, raw: pd.DataFrame, clean: pd.DataFrame) -> pd.DataFrame:
    """Walk from the Raw label counts to the Clean 491 / 2,009 result step by step."""
    lbl = df["Churn_Label"].astype(str).str.strip()
    yes_raw, no_raw, one_raw = int((lbl == "Yes").sum()), int((lbl == "No").sum()), int((~lbl.isin(["Yes", "No"])).sum())
    d1 = df.drop_duplicates()
    l1 = d1["Churn_Label"].astype(str).str.strip()
    d2 = d1.drop_duplicates("Customer_ID", keep="first")
    l2 = d2["Churn_Label"].astype(str).str.strip()
    # Describe what R2 actually did to the conflicting ID(s): first occurrence kept.
    dup_mask = d1.duplicated("Customer_ID", keep=False)
    r2_notes = []
    for cid, g in d1[dup_mask].groupby("Customer_ID"):
        labels = g["Churn_Label"].astype(str).str.strip().tolist()
        r2_notes.append(f"ID {cid}: first occurrence kept (label '{labels[0]}'), dropped label(s) {labels[1:]}")
    r2_note = "; ".join(r2_notes) if r2_notes else "no duplicate IDs"
    one2 = ~l2.isin(["Yes", "No"])
    one_with_date = int((one2 & d2["Date_Cancelled"].notna()).sum())
    one_no_date = int((one2 & d2["Date_Cancelled"].isna()).sum())
    rows = [
        ("Raw file", "Yes", yes_raw, "literal 'Yes' rows in the 2,510-row extract"),
        ("Raw file", "No", no_raw, ""),
        ("Raw file", "1 (ambiguous)", one_raw, "numeric code, not counted as churned in the Raw view"),
        ("R1 exact duplicates removed", "Yes", int((l1 == "Yes").sum()), f"{yes_raw - int((l1 == 'Yes').sum())} duplicate Yes rows removed"),
        ("R1 exact duplicates removed", "No", int((l1 == "No").sum()), f"{no_raw - int((l1 == 'No').sum())} duplicate No rows removed"),
        ("R1 exact duplicates removed", "1 (ambiguous)", int((~l1.isin(['Yes', 'No'])).sum()), ""),
        ("R2 duplicate Customer_ID resolved", "Yes", int((l2 == "Yes").sum()), r2_note),
        ("R2 duplicate Customer_ID resolved", "No", int((l2 == "No").sum()), r2_note),
        ("R2 duplicate Customer_ID resolved", "1 (ambiguous)", int(one2.sum()), r2_note),
        ("R4 recode '1' by Date_Cancelled", "1 -> Yes", one_with_date, "Date_Cancelled present"),
        ("R4 recode '1' by Date_Cancelled", "1 -> No", one_no_date, "Date_Cancelled absent"),
        ("Clean table", "Yes", int(clean["churned"].sum()), f"{int((l2 == 'Yes').sum())} Yes after R2 + {one_with_date} recoded from 1 = {int(clean['churned'].sum())}"),
        ("Clean table", "No", int((~clean["churned"]).sum()), f"{int((l2 == 'No').sum())} No after R2 + {one_no_date} recoded from 1 = {int((~clean['churned']).sum())}"),
        ("Clean table", "total", len(clean), ""),
    ]
    return pd.DataFrame(rows, columns=["step", "label", "rows", "note"])


def audit_issue_counts(df: pd.DataFrame) -> pd.DataFrame:
    lbl = df["Churn_Label"].astype(str).str.strip()
    ids = lambda mask: ", ".join(map(str, df.loc[mask, "Customer_ID"].tolist()))
    dup_ids = df[df.duplicated("Customer_ID", keep=False)]
    conflicting = dup_ids.groupby("Customer_ID").filter(lambda g: len(g.drop_duplicates()) > 1)["Customer_ID"].unique()
    rows = [
        ("Rows in file vs. IT claim", len(df), f"IT claimed {2500}; +{len(df) - 2500}"),
        ("Exact duplicate rows", int(df.duplicated().sum()), ""),
        ("Duplicate Customer_ID rows (after exact duplicates)", int(df.drop_duplicates().duplicated("Customer_ID").sum()), f"Customer_ID {', '.join(map(str, conflicting))} (No vs 1)"),
        ("Industry rows not in Title Case", int((df["Industry"] != df["Industry"].str.strip().str.title()).sum()), "12 spellings for 6 industries; 250 remain after R1"),
        ("Churn_Label coded as number", int((~lbl.isin(["Yes", "No"])).sum()), f"{int((~lbl.isin(['Yes', 'No']) & df['Date_Cancelled'].notna()).sum())} have Date_Cancelled"),
        ("Contract_Type = Unknown", int((df["Contract_Type"] == "Unknown").sum()), ""),
        ("Support_Tickets missing", int(df["Support_Tickets"].isna().sum()), f"{100 * df['Support_Tickets'].isna().mean():.1f}% of rows"),
        ("Negative Total_Users", int((df["Total_Users"] < 0).sum()), f"Customer_ID {ids(df['Total_Users'] < 0)}"),
        ("Negative Monthly_Revenue", int((df["Monthly_Revenue"] < 0).sum()), f"Customer_ID {ids(df['Monthly_Revenue'] < 0)}"),
        ("Negative Last_Login_Days_Ago", int((df["Last_Login_Days_Ago"] < 0).sum()), "all -5"),
        ("Monthly_Revenue >= $100,000", int((df["Monthly_Revenue"] >= WHALE_REVENUE).sum()), f"Customer_ID {ids(df['Monthly_Revenue'] >= WHALE_REVENUE)} ($1,000,000, 17 users)"),
        ("Total_Users >= 10,000", int((df["Total_Users"] >= EXTREME_USERS).sum()), f"Customer_ID {ids(df['Total_Users'] >= EXTREME_USERS)} (50,000 users, $1,133)"),
        ("Monthly_Revenue cells with $ display format", 378, "display only; all values are numeric"),
        ("Date_Cancelled populated", int(df["Date_Cancelled"].notna().sum()),
         f"{int(((lbl == 'Yes') & df['Date_Cancelled'].notna()).sum())} rows labelled Yes + {int((~lbl.isin(['Yes', 'No']) & df['Date_Cancelled'].notna()).sum())} rows coded 1; never present for No: leakage column"),
    ]
    return pd.DataFrame(rows, columns=["issue", "count", "detail"])


def whale_impact(df: pd.DataFrame, clean: pd.DataFrame) -> pd.DataFrame:
    rev = df["Monthly_Revenue"]
    with_w, without_w = rev.mean(), rev[rev < WHALE_REVENUE].mean()
    usr = df["Total_Users"]
    with_u, without_u = usr.mean(), usr[usr < EXTREME_USERS].mean()
    rows = [
        ("Monthly_Revenue", "Raw mean, all 2,510 rows (with whale)", round(with_w, 2)),
        ("Monthly_Revenue", "Raw mean, whale row excluded only", round(without_w, 2)),
        ("Monthly_Revenue", "Difference attributable to the whale", round(with_w - without_w, 2)),
        ("Monthly_Revenue", "Whale inflates the mean by (%)", round(100 * (with_w - without_w) / without_w, 1)),
        ("Monthly_Revenue", "Raw median (unaffected)", rev.median()),
        ("Monthly_Revenue", "Clean mean (dedup + whale and -500 set to missing)", round(clean["Monthly_Revenue"].mean(), 2)),
        ("Total_Users", "Raw mean, all rows (with 50,000)", round(with_u, 2)),
        ("Total_Users", "Raw mean, 50,000 row excluded only", round(without_u, 2)),
        ("Total_Users", "Difference attributable to the outlier", round(with_u - without_u, 2)),
        ("Total_Users", "Outlier inflates the mean by (%)", round(100 * (with_u - without_u) / without_u, 1)),
        ("Total_Users", "Raw median (unaffected)", usr.median()),
        ("Total_Users", "Clean mean (dedup + 50,000 and -10 set to missing)", round(clean["Total_Users"].mean(), 2)),
    ]
    out = pd.DataFrame(rows, columns=["field", "measure", "value"])
    out["note"] = ""
    out.loc[out["measure"].str.startswith("Clean"), "note"] = "Raw-minus-Clean is NOT the whale effect alone: it also includes de-duplication and the negative-value rule."
    return out


def leakage_diagnostic(df: pd.DataFrame, raw: pd.DataFrame, clean: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for view, d in (("Raw", raw), ("Clean", clean)):
        lbl = d["Churn_Label"].astype(str).str.strip()
        lbl = lbl.where(lbl.isin(["Yes", "No"]), "1 (ambiguous)")
        has = d["Date_Cancelled"].notna()
        for v in ("Yes", "No", "1 (ambiguous)"):
            m = lbl == v
            if m.any():
                rows.append((view, "Churn_Label = " + v, int(m.sum()), int((m & has).sum()), int((m & ~has).sum()), None, ""))
        ext = df["Date_Cancelled"].max()
        login = d["Last_Login_Days_Ago"]
        login_missing = has & login.isna()
        after = has & login.notna() & ((ext - d["Date_Cancelled"]).dt.days > login.abs())
        not_after = has & login.notna() & ~after
        rows.append((view, "Cancelled rows: Last_Login_Days_Ago implies a login AFTER cancellation", int(has.sum()), int(after.sum()), int(not_after.sum()), int(login_missing.sum()),
                     f"extract date taken as {ext.date()}; Raw uses abs() of the 50 negative values" if view == "Raw"
                     else "Clean: 2 duplicate cancelled rows removed; negative login values are missing after R5 and are counted in login_missing, not in either bucket"))
    out = pd.DataFrame(rows, columns=["view", "group", "rows", "with_Date_Cancelled / login_after_cancel", "without_Date_Cancelled / login_not_after", "login_missing", "note"])
    return out


NOTES = {
    "rows": "Raw = all rows in the file; Clean = after R1 (exact duplicates) and R2 (duplicate IDs).",
    "churned_yes": "Raw = literal 'Yes' only. Clean = 486 Raw Yes - 2 duplicate Yes rows removed (R1/R2) + 7 numeric '1' rows that have a Date_Cancelled (R4) = 491. See Label_Reconciliation.",
    "ambiguous_label_rows": "Rows whose Churn_Label is neither Yes nor No (coded '1'). Zero after R4.",
    "overall_churn_rate_pct": "churned_yes / rows.",
    "churn_rate_20XX_pct_of_active_base": "cancellations in year / customers who joined before Jan 1 and had not cancelled earlier. Same rule both views; Clean removes duplicate rows from the base.",
    "mean_monthly_revenue": "Raw includes the $1,000,000 row and the -$500 row; Clean sets both to missing (R5, R6).",
    "mean_total_users": "Raw includes the 50,000 row and the -10 row; Clean sets both to missing.",
    "mean_support_tickets": "Raw = mean of non-missing; Clean = blanks imputed as 0 (R7), which lowers the mean. {support_median_note}",
    "mean_last_login_days_ago": "Raw includes 50 rows at -5; Clean sets them to missing. Field is not time-anchored in either view.",
    "distinct_industry_values": "Raw counts case variants separately (12); Clean title-cases (6).",
    "largest_industry / churn by industry": "Raw groups use the spelling as delivered, so 'retail' and 'Retail' are separate rows; use the Clean table for segment statements.",
}


# --------------------------------------------------------------------------- #
# Assembly
# --------------------------------------------------------------------------- #
def side_by_side(builder, raw: pd.DataFrame, clean: pd.DataFrame, key: str | None = None, **kw) -> pd.DataFrame:
    a = builder(raw, **kw).assign(view="Raw")
    b = builder(clean, **kw).assign(view="Clean")
    out = pd.concat([a, b], ignore_index=True)
    cols = ["view"] + [c for c in out.columns if c != "view"]
    out = out[cols]
    if key:
        out = out.sort_values([key, "view"], ascending=[True, False], kind="stable").reset_index(drop=True)
    return out


def to_markdown_table(df: pd.DataFrame) -> str:
    df = df.copy()
    for c in df.columns:
        df[c] = df[c].map(lambda v: "" if pd.isna(v) else (f"{v:,.0f}" if isinstance(v, float) and float(v).is_integer() else (f"{v:,.2f}" if isinstance(v, float) else str(v))))
    head = "| " + " | ".join(df.columns) + " |\n|" + "---|" * len(df.columns) + "\n"
    body = "".join("| " + " | ".join(r) + " |\n" for r in df.astype(str).values.tolist())
    return head + body


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", nargs="?", default="B2B SaaS Churn Data.xlsx")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    inp = Path(args.input)
    if not inp.exists():
        sys.exit(f"Input file not found: {inp}")
    out_dir = Path(args.out) if args.out else inp.parent

    df = pd.read_excel(inp)
    raw = raw_view(df)
    clean, change_log = clean_view(df)

    h_raw, h_clean = headline(raw), headline(clean)
    # Alternative imputation for the note: median of the observed Clean values.
    st = clean["Support_Tickets"].where(clean["Support_Tickets_Missing"] == 0)
    st_median = st.median()
    NOTES["mean_support_tickets"] = NOTES["mean_support_tickets"].format(
        support_median_note=f"If blanks are imputed with the median ({st_median:g}) instead, the Clean mean is {st.fillna(st_median).mean():.2f}."
    )
    metrics = pd.DataFrame(
        [
            {"metric": k, "raw": h_raw[k], "clean": h_clean[k], "note": NOTES.get(k, NOTES.get("churn_rate_20XX_pct_of_active_base") if k.startswith("churn_rate_20") else "")}
            for k in h_raw
        ]
    )

    tables = {
        "Metrics": metrics,
        "Churn_By_Year": side_by_side(churn_by_year, raw, clean, key="year"),
        "Churn_By_Industry": side_by_side(churn_by_group, raw, clean, col="industry", name="industry"),
        "Churn_By_Contract": side_by_side(churn_by_group, raw, clean, col="contract", name="contract"),
        "Industry_x_Contract": side_by_side(industry_x_contract, raw, clean),
        "Churn_By_JoinYear": side_by_side(churn_by_join_year, raw, clean, key="join_year"),
        "Revenue_Baseline": side_by_side(numeric_baseline, raw, clean, col="Monthly_Revenue"),
        "Users_Baseline": side_by_side(numeric_baseline, raw, clean, col="Total_Users"),
        "Engagement": side_by_side(engagement, raw, clean),
        "Label_Reconciliation": label_reconciliation(df, raw, clean),
        "Audit_Issue_Counts": audit_issue_counts(df),
        "Whale_Impact": whale_impact(df, clean),
        "Leakage_Diagnostic": leakage_diagnostic(df, raw, clean),
        "Clean_Change_Log": change_log,
        "Definitions": pd.DataFrame(
            [
                {"term": "Raw", "definition": "Extract as delivered, 2,510 rows, no values changed. Churned = Churn_Label is literally 'Yes'. Rows coded '1' are reported in ambiguous_label_rows and are NOT counted as churned."},
                {"term": "Clean", "definition": "After rules " + ", ".join(k for k, v in RULES.items() if v and not k.startswith("R9")) + " from tw1_data_cleaning.py, applied in memory. Date_Cancelled retained for reporting (R9 applies to the modelling table only)."},
                {"term": "Churned (Clean)", "definition": "'Yes' rows plus '1' rows that have a Date_Cancelled; '1' rows without a date become 'No' and carry Churn_Label_Flag."},
                {"term": "churn_rate_pct_of_active_base", "definition": "Cancellations dated in the year / customers with Join_Date before Jan 1 of that year who had not cancelled in an earlier year."},
                {"term": "cancellations_pct_of_all_rows", "definition": "Cancellations dated in the year / all rows in the view (simple Excel-style share)."},
                {"term": "Missing after cleaning", "definition": "Negative values and the two extreme outliers are set to missing, not corrected; means and medians in the Clean view exclude them automatically."},
                {"term": "Caveat", "definition": "No cancellations exist before 2022 although customers joined from 2020; Last_Login_Days_Ago is not anchored to one snapshot date in either view."},
            ]
        ),
    }

    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx = out_dir / "TW1_Raw_vs_Clean_Metrics.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        for name, t in tables.items():
            t.to_excel(xw, sheet_name=name, index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 70)

    md = ["# TW1 Raw vs. Clean metrics (auto-generated by tw1_raw_vs_clean_metrics.py)\n",
          "Raw = extract as delivered (2,510 rows; churned = literal 'Yes'). Clean = after cleaning rules R1-R8 applied in memory (see Definitions).\n"]
    for name, t in tables.items():
        if name == "Clean_Change_Log":
            continue
        md.append(f"\n## {name.replace('_', ' ')}\n\n" + to_markdown_table(t))
    (out_dir / "TW1_Raw_vs_Clean_Metrics.md").write_text("".join(md), encoding="utf-8")

    print(metrics.to_string(index=False))
    print(f"\nWritten: {xlsx}\nWritten: {out_dir / 'TW1_Raw_vs_Clean_Metrics.md'}")


if __name__ == "__main__":
    main()
