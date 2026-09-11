"""
TW1 - Data Health Checkup: remediation / cleaning script (DISABLED BY DEFAULT).

Course: MASY1-GC 2100 Advanced Business Analytics, Team session 2-4 5
Case:   TechPoint SaaS Solutions - B2B SaaS Churn Data

STATUS: DRAFT - NOT YET ENABLED.
    The rules below are a proposed implementation of the Part 4 remediation
    strategy (see TW1_Remediation_Plan.md). Until the team has agreed on Part 4,
    the script only runs in DRY-RUN mode: it reports what it WOULD change and
    writes nothing. Pass --apply to actually write the cleaned file.
    Each rule can be switched on/off in the RULES dictionary so the team can
    adjust the strategy without rewriting code.

Pipeline (in order)
-------------------
 R1  drop_exact_duplicates       remove byte-identical rows
 R2  resolve_duplicate_ids       one row per Customer_ID; conflicting rows are
                                 logged and the first occurrence is kept
 R3  standardize_industry        trim + Title Case  ('retail' -> 'Retail')
 R4  standardize_churn_label     Yes/No text; numeric codes resolved against
                                 Date_Cancelled (present -> Yes, absent -> No)
                                 and flagged in a Churn_Label_Flag column
 R5  fix_negative_values         impossible negatives -> NaN + flag column
                                 (values are NOT guessed; they are treated as
                                 missing so the model can learn from the flag)
 R6  cap_extreme_outliers        whale revenue / 50,000 users -> NaN + flag
                                 (excluded from baselines; kept in a review list)
 R7  impute_support_tickets      blank Support_Tickets -> 0 + Support_Tickets_Missing
                                 flag (assumption: blank = no ticket logged;
                                 alternative = median; set in RULES)
 R8  contract_unknown            keep 'Unknown' as its own category + flag
 R9  drop_leakage_columns        remove Date_Cancelled, Company_Name (and optionally
                                 Last_Login_Days_Ago) from the modelling table;
                                 Churn_Year goes to a separate reporting sheet
 R10 write_outputs               cleaned xlsx/csv + change log

The pipeline is also importable: clean_frame(df, drop_leakage=False) applies
the rules in memory (no files written) so tw1_raw_vs_clean_metrics.py can
compare raw and cleaned metrics side by side.

Usage
-----
    python tw1_data_cleaning.py "B2B SaaS Churn Data.xlsx"            # dry run, no files written
    python tw1_data_cleaning.py "B2B SaaS Churn Data.xlsx" --apply    # writes cleaned outputs

Dependencies: pandas, openpyxl
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Rule switches - the team edits THIS block after agreeing on Part 4.
# --------------------------------------------------------------------------- #
RULES = {
    "R1_drop_exact_duplicates": True,
    "R2_resolve_duplicate_ids": True,
    "R3_standardize_industry": True,
    "R4_standardize_churn_label": True,
    "R5_fix_negative_values": True,
    "R6_cap_extreme_outliers": True,
    "R7_impute_support_tickets": True,
    "R8_contract_unknown": True,
    "R9_drop_leakage_columns": True,
}
SUPPORT_TICKET_IMPUTATION = "zero"      # "zero" | "median" | "leave"
WHALE_REVENUE = 100_000                 # Monthly_Revenue at/above this is treated as an anomaly
EXTREME_USERS = 10_000                  # Total_Users at/above this is treated as an anomaly
DROP_LAST_LOGIN = False                 # set True if the team decides Last_Login_Days_Ago is also leakage
LEAKAGE_COLUMNS = ["Date_Cancelled"]         # outcome information, known only after churn
IDENTIFIER_COLUMNS = ["Company_Name"]        # identifiers carry no signal; Customer_ID is kept as join key only

NUMERIC_COLS = ["Total_Users", "Monthly_Revenue", "Support_Tickets", "Last_Login_Days_Ago"]


class ChangeLog:
    """Collects one line per rule so the memo can cite exactly what changed."""

    def __init__(self) -> None:
        self.rows: list[dict] = []

    def add(self, rule: str, action: str, rows_affected: int, detail: str = "") -> None:
        self.rows.append({"rule": rule, "action": action, "rows_affected": int(rows_affected), "detail": detail})
        print(f"[{rule}] {action}: {rows_affected} row(s) {detail}")

    def frame(self) -> pd.DataFrame:
        return pd.DataFrame(self.rows)


# --------------------------------------------------------------------------- #
# Rules
# --------------------------------------------------------------------------- #
def r1_drop_exact_duplicates(df: pd.DataFrame, log: ChangeLog) -> pd.DataFrame:
    n = int(df.duplicated().sum())
    log.add("R1", "drop exact duplicate rows", n)
    return df.drop_duplicates()


def r2_resolve_duplicate_ids(df: pd.DataFrame, log: ChangeLog) -> tuple[pd.DataFrame, pd.DataFrame]:
    conflicts = df[df.duplicated("Customer_ID", keep=False)].sort_values("Customer_ID")
    n = int(df.duplicated("Customer_ID").sum())
    log.add("R2", "drop repeated Customer_ID (keep first occurrence)", n, f"conflicting IDs logged: {conflicts['Customer_ID'].nunique()}")
    return df.drop_duplicates("Customer_ID", keep="first"), conflicts


def r3_standardize_industry(df: pd.DataFrame, log: ChangeLog) -> pd.DataFrame:
    fixed = df["Industry"].astype(str).str.strip().str.title()
    log.add("R3", "standardize Industry casing", int((fixed != df["Industry"]).sum()))
    return df.assign(Industry=fixed)


def r4_standardize_churn_label(df: pd.DataFrame, log: ChangeLog) -> pd.DataFrame:
    raw = df["Churn_Label"].astype(str).str.strip().str.lower()
    has_date = df["Date_Cancelled"].notna()
    label = pd.Series(np.where(raw == "yes", "Yes", np.where(raw == "no", "No", None)), index=df.index, dtype="object")
    ambiguous = label.isna()
    # Numeric / unrecognised codes: resolve against the cancellation date.
    label[ambiguous & has_date] = "Yes"
    label[ambiguous & ~has_date] = "No"
    # Flag records only that the label was recoded; it must NOT say whether a
    # cancellation date existed, otherwise the flag itself leaks the outcome.
    flag = np.where(ambiguous, "recoded_from_" + df["Churn_Label"].astype(str), "")
    log.add("R4", "recode non Yes/No Churn_Label values using Date_Cancelled", int(ambiguous.sum()),
            f"({int((ambiguous & has_date).sum())} -> Yes, {int((ambiguous & ~has_date).sum())} -> No)")
    return df.assign(Churn_Label=label, Churn_Label_Flag=flag)


def r5_fix_negative_values(df: pd.DataFrame, log: ChangeLog) -> pd.DataFrame:
    out = df.copy()
    for col in NUMERIC_COLS:
        neg = out[col] < 0
        if neg.any():
            out[f"{col}_Flag"] = np.where(neg, "negative_set_to_missing", out.get(f"{col}_Flag", ""))
            out.loc[neg, col] = np.nan
        log.add("R5", f"negative {col} -> missing + flag", int(neg.sum()))
    return out


def r6_cap_extreme_outliers(df: pd.DataFrame, log: ChangeLog) -> tuple[pd.DataFrame, pd.DataFrame]:
    out = df.copy()
    whale = out["Monthly_Revenue"] >= WHALE_REVENUE
    users = out["Total_Users"] >= EXTREME_USERS
    review = out[whale | users].copy()
    out["Monthly_Revenue_Flag"] = np.where(whale, "extreme_outlier_set_to_missing", out.get("Monthly_Revenue_Flag", ""))
    out["Total_Users_Flag"] = np.where(users, "extreme_outlier_set_to_missing", out.get("Total_Users_Flag", ""))
    out.loc[whale, "Monthly_Revenue"] = np.nan
    out.loc[users, "Total_Users"] = np.nan
    log.add("R6", f"Monthly_Revenue >= {WHALE_REVENUE:,} -> missing + flag (kept in review list)", int(whale.sum()))
    log.add("R6", f"Total_Users >= {EXTREME_USERS:,} -> missing + flag (kept in review list)", int(users.sum()))
    return out, review


def r7_impute_support_tickets(df: pd.DataFrame, log: ChangeLog) -> pd.DataFrame:
    out = df.copy()
    missing = out["Support_Tickets"].isna()
    out["Support_Tickets_Missing"] = missing.astype(int)
    if SUPPORT_TICKET_IMPUTATION == "zero":
        out.loc[missing, "Support_Tickets"] = 0
    elif SUPPORT_TICKET_IMPUTATION == "median":
        out.loc[missing, "Support_Tickets"] = out["Support_Tickets"].median()
    log.add("R7", f"impute blank Support_Tickets ({SUPPORT_TICKET_IMPUTATION}) + missing flag", int(missing.sum()))
    return out


def r8_contract_unknown(df: pd.DataFrame, log: ChangeLog) -> pd.DataFrame:
    unk = df["Contract_Type"].astype(str).str.strip().str.lower() == "unknown"
    log.add("R8", "keep Contract_Type 'Unknown' as explicit category + flag", int(unk.sum()))
    return df.assign(Contract_Type_Unknown=unk.astype(int))


def r9_drop_leakage_columns(df: pd.DataFrame, log: ChangeLog) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (modelling_table, reporting_frame).

    The modelling table must contain nothing derived from the outcome or from
    the customer's identity: Date_Cancelled (direct leak), Churn_Year (derived
    from Date_Cancelled), and Company_Name (identifier) are removed. Customer_ID
    stays only as a join key and must be excluded from model features.
    Churn_Year is written to a separate reporting frame keyed by Customer_ID.
    """
    cols = list(LEAKAGE_COLUMNS) + list(IDENTIFIER_COLUMNS) + (["Last_Login_Days_Ago"] if DROP_LAST_LOGIN else [])
    reporting = pd.DataFrame(
        {
            "Customer_ID": df["Customer_ID"],
            "Churn_Label": df["Churn_Label"],
            "Date_Cancelled": df["Date_Cancelled"],
            "Churn_Year": df["Date_Cancelled"].dt.year,
        }
    )
    present = [c for c in cols if c in df.columns]
    log.add("R9", f"drop leakage/identifier columns {present} from the modelling table (Churn_Year kept in a separate reporting frame)", len(df))
    return df.drop(columns=present), reporting


def clean_frame(df: pd.DataFrame, drop_leakage: bool = True) -> tuple[pd.DataFrame, ChangeLog, dict[str, pd.DataFrame]]:
    """Apply the enabled rules to a DataFrame in memory and return
    (cleaned_df, change_log, review_frames). Writes nothing.
    drop_leakage=False keeps Date_Cancelled so reporting scripts can
    compute churn-year metrics on the cleaned table."""
    log = ChangeLog()
    log.add("R0", "rows loaded", len(df))
    review_frames: dict[str, pd.DataFrame] = {}

    if RULES["R1_drop_exact_duplicates"]:
        df = r1_drop_exact_duplicates(df, log)
    if RULES["R2_resolve_duplicate_ids"]:
        df, review_frames["Review_DuplicateIDs"] = r2_resolve_duplicate_ids(df, log)
    if RULES["R3_standardize_industry"]:
        df = r3_standardize_industry(df, log)
    if RULES["R4_standardize_churn_label"]:
        df = r4_standardize_churn_label(df, log)
    if RULES["R5_fix_negative_values"]:
        df = r5_fix_negative_values(df, log)
    if RULES["R6_cap_extreme_outliers"]:
        df, review_frames["Review_Outliers"] = r6_cap_extreme_outliers(df, log)
    if RULES["R7_impute_support_tickets"]:
        df = r7_impute_support_tickets(df, log)
    if RULES["R8_contract_unknown"]:
        df = r8_contract_unknown(df, log)
    if RULES["R9_drop_leakage_columns"] and drop_leakage:
        df, review_frames["Reporting_ChurnDates"] = r9_drop_leakage_columns(df, log)

    log.add("R10", "rows in cleaned table", len(df), f"columns: {len(df.columns)}")
    return df, log, review_frames


def run(input_path: Path, output_dir: Path, apply: bool) -> None:
    df, log, review_frames = clean_frame(pd.read_excel(input_path))

    if not apply:
        print("\nDRY RUN - nothing written. Re-run with --apply once the team has approved the Part 4 strategy.")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    xlsx = output_dir / "B2B_SaaS_Churn_Data_CLEANED.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Cleaned", index=False)
        log.frame().to_excel(xw, sheet_name="Change_Log", index=False)
        for name, f in review_frames.items():
            f.to_excel(xw, sheet_name=name, index=False)
    df.to_csv(output_dir / "B2B_SaaS_Churn_Data_CLEANED.csv", index=False)
    print(f"\nWritten: {xlsx}\nWritten: {output_dir / 'B2B_SaaS_Churn_Data_CLEANED.csv'}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", nargs="?", default="B2B SaaS Churn Data.xlsx", help="raw Excel extract")
    ap.add_argument("--out", default=None, help="output directory (default: next to the input file)")
    ap.add_argument("--apply", action="store_true", help="actually write the cleaned files (default is dry run)")
    args = ap.parse_args()
    inp = Path(args.input)
    if not inp.exists():
        sys.exit(f"Input file not found: {inp}")
    run(inp, Path(args.out) if args.out else inp.parent, args.apply)


if __name__ == "__main__":
    main()
