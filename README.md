# Insurance Renewal Tracker

Automates monthly monitoring of **policy renewals** for a motor insurance portfolio. The script
finds which expiring policies were renewed, measures the renewal rate across several
dimensions, and sends every sales manager a follow-up file with their own non-renewed policies.

## What it does

**1. Matches renewals.** For each expiring policy it looks for a new policy with the same
**plate number**, or the same **customer** as a fallback, that starts between 30 days before
and 60 days after the expiry date. Each new policy can renew only one old policy. Policies that
expired less than 30 days ago and are still unmatched are marked **Pending** and are not counted
as lost yet.

**2. Builds the renewal report** (Excel), with renewal rate by count and **premium retention**:

| Sheet | Breakdown |
|---|---|
| By channel | Agents, Regional, Key Accounts, SME, Partners |
| By client type | Individual / Corporate |
| By department | channel → department |
| By expiry month | monthly trend |
| Non-renewal reasons | why customers left (sold the car, price, competitor, …) |
| Policies | row-level data with the matched renewal policy |

**3. Splits follow-up files by manager** (`--split`). Each manager gets a workbook with only
their open policies, a **Status dropdown** (Positive / Negative / Wait) and a **protected sheet**
where only *Status* and *Comment* can be edited. This keeps the returned files consistent and
easy to consolidate.

## Quick start

```bash
pip install -r requirements.txt
python generate_sample_data.py
python renewal_tracker.py --as-of 2026-05-31 --split --password demo
```

```
As of 2026-05-31: 1500 expired policies | renewed 1020 | not renewed 380 | pending 100 | renewal rate 72.9%
Report saved to output/renewal_report.xlsx
7 manager follow-up files saved to output/follow_up
```

| channel | Expired | Renewed | Not renewed | Pending | Renewal rate | Premium retention |
|---|---:|---:|---:|---:|---:|---:|
| Agents | 491 | 353 | 138 | 38 | 71.9% | 72.0% |
| Key Accounts | 139 | 100 | 39 | 9 | 71.9% | 71.7% |
| Partners | 216 | 148 | 68 | 9 | 68.5% | 67.3% |
| Regional | 206 | 160 | 46 | 17 | 77.7% | 79.0% |
| SME | 348 | 259 | 89 | 27 | 74.4% | 75.4% |
| **Total** | **1400** | **1020** | **380** | **100** | **72.9%** | **73.1%** |

The protection password can also be passed via the `RENEWAL_FILE_PASSWORD` environment variable,
so it does not end up in your shell history or in the code.

## Tech

Python · pandas · openpyxl (data validation, sheet protection, styling)

> The data in this repository is randomly generated. No real company data is included.
