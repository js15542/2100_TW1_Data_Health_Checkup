# TW1 Data Health Checkup - working files

Team session 2-4 5 | MASY1-GC 2100 | Raw data: `B2B SaaS Churn Data.xlsx` (do not edit the raw file)

## Where things live

- **Code project (run here):** `D:\0Coding\Python\1\2026Fall\2100_TW1_Data_Health_Checkup\`
  scripts, raw data copy, requirements, and generated outputs
- **Course folder (docs and results only):** `OneDrive\0graduate\2026Fall\MASY1-GC_2100_Advanced Business Analytics\2100_homework\Module 1\Session 2\`
  original data and assignment notes only

| File | Purpose | Status |
|------|---------|--------|
| `tw1_data_health_audit.py` | Part 1 audit (read-only) + Part 2/3 supporting numbers | done, re-runnable |
| `TW1_Audit_Findings.xlsx` | Output of the audit: Summary sheet + one sheet per check with the offending rows | generated |
| `TW1_Audit_Summary.md` | Plain-English audit results to paste into the memo | generated |
| `TW1_Remediation_Plan.md` | Part 4 proposal, cleaning rules, open decisions, draft inputs for Parts 2/3/5 | draft for team review |
| `tw1_data_cleaning.py` | Cleaning pipeline; dry-run by default, `--apply` writes cleaned data | NOT enabled until the team approves Part 4 |
| `requirements.txt` | pandas, openpyxl | |

## Run (from the code project folder, inside its virtual environment)

```bash
pip install -r requirements.txt
python tw1_data_health_audit.py "B2B SaaS Churn Data.xlsx"          # rewrites TW1_Audit_Findings.xlsx and TW1_Audit_Summary.md here
python tw1_data_cleaning.py "B2B SaaS Churn Data.xlsx"              # dry run: prints what would change, writes nothing
python tw1_data_cleaning.py "B2B SaaS Churn Data.xlsx" --apply      # writes B2B_SaaS_Churn_Data_CLEANED.xlsx/.csv
```

Both scripts take the input path as the first argument; outputs go next to the input unless a second argument / `--out` is given. Generated result files are committed here; the course folder keeps only the assignment notes.
