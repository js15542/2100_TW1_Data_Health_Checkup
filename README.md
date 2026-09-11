# TW1 Data Health Checkup

MASY1-GC 2100 Advanced Business Analytics | Team session 2-4 5 | Case: TechPoint SaaS Solutions

Part 1 audit of `B2B SaaS Churn Data.xlsx`, plus the Raw vs. Clean numbers Parts 2-5 cite. Teammates can read everything in `outputs/` without running any code.

## Layout

```
.
├── README.md
├── requirements.txt
├── data/
│   ├── raw/       B2B SaaS Churn Data.xlsx               raw extract, never edited
│   └── cleaned/   B2B_SaaS_Churn_Data_CLEANED.xlsx/.csv  written by tw1_data_cleaning.py --apply
├── src/
│   ├── tw1_data_health_audit.py       Part 1 audit (read-only) + Part 2/3 support numbers
│   ├── tw1_data_cleaning.py           cleaning rules R1-R9; dry-run by default
│   └── tw1_raw_vs_clean_metrics.py    Parts 2-5 numbers, Raw and Clean side by side
├── outputs/
│   ├── TW1_Audit_Findings.xlsx        one sheet per check with the offending rows
│   ├── TW1_Audit_Summary.md           audit results in plain English
│   ├── TW1_Raw_vs_Clean_Metrics.xlsx  Metrics, Churn_By_*, Label_Reconciliation, Audit_Issue_Counts,
│   │                                  Whale_Impact, Leakage_Diagnostic, Definitions
│   └── TW1_Raw_vs_Clean_Metrics.md    same tables as Markdown
└── docs/
    └── TW1_Remediation_Plan.md        Part 4 proposal, rule rationale, open decisions, Part 5 draft ideas
```

## Run (from the project root, inside the virtual environment)

```bash
pip install -r requirements.txt
python src/tw1_data_health_audit.py          # rewrites outputs/TW1_Audit_*
python src/tw1_raw_vs_clean_metrics.py       # rewrites outputs/TW1_Raw_vs_Clean_Metrics.*
python src/tw1_data_cleaning.py              # dry run: prints what would change, writes nothing
python src/tw1_data_cleaning.py --apply      # writes data/cleaned/
```

Every script accepts an explicit input path as the first argument and `--out DIR` (or a second positional argument for the audit script) to override the default locations.

## Reading guide for teammates

- Business conclusions (churn trend, segments): use the **Clean** columns in `outputs/TW1_Raw_vs_Clean_Metrics.xlsx`.
- Data-quality statements and outlier impact: use the **Raw** columns and the `Whale_Impact` / `Audit_Issue_Counts` sheets.
- Row-level evidence for any issue: `outputs/TW1_Audit_Findings.xlsx`.
- Cleaning rules and the decisions still open for the team: `docs/TW1_Remediation_Plan.md`.
