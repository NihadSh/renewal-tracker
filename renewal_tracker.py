"""Motor-insurance renewal tracker.

1. Matches expiring policies with new sales to find which ones were renewed
   (same plate number or same customer, starting close to the expiry date).
2. Builds an Excel report with renewal rates by channel, client type,
   department, expiry month and non-renewal reason - by count and by premium.
3. Optionally (--split) creates one follow-up workbook per manager with that
   manager's open policies, a Status dropdown (Positive / Negative / Wait) and
   a protected sheet so that only the Status and Comment columns are editable.

Usage:
    python renewal_tracker.py --as-of 2026-05-31
    python renewal_tracker.py --as-of 2026-05-31 --split --password secret
"""
import argparse
import os
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).parent
WINDOW_BEFORE, WINDOW_AFTER = 30, 60  # days around expiry that still count as a renewal
GRACE_DAYS = 30                       # unmatched for less than this -> still "Pending"


# --------------------------------------------------------------------------- matching
def match_renewals(expiring: pd.DataFrame, new_sales: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    """Add renewal_policy / renewal_premium / status columns to the expiring policies."""
    exp = expiring.copy()
    exp["renewal_policy"] = pd.NA
    exp["renewal_premium"] = pd.NA

    for key in ["plate_number", "customer_id"]:  # plate first, customer as a fallback
        todo = exp[exp["renewal_policy"].isna()]
        cand = todo[["policy_number", key, "end_date"]].merge(
            new_sales[[key, "policy_number", "start_date", "premium"]].rename(
                columns={"policy_number": "new_policy", "premium": "new_premium"}),
            on=key)
        delta = (cand["start_date"] - cand["end_date"]).dt.days
        cand = cand[delta.between(-WINDOW_BEFORE, WINDOW_AFTER)]
        cand = cand.assign(gap=delta.abs()).sort_values("gap").drop_duplicates("policy_number")
        # one new policy may renew only one old policy
        cand = cand[~cand["new_policy"].isin(exp["renewal_policy"].dropna())].drop_duplicates("new_policy")
        idx = exp["policy_number"].isin(cand["policy_number"])
        m = cand.set_index("policy_number")
        exp.loc[idx, "renewal_policy"] = exp.loc[idx, "policy_number"].map(m["new_policy"])
        exp.loc[idx, "renewal_premium"] = exp.loc[idx, "policy_number"].map(m["new_premium"])

    renewed = exp["renewal_policy"].notna()
    pending = ~renewed & (exp["end_date"] > as_of - pd.Timedelta(days=GRACE_DAYS))
    exp["status"] = "Not renewed"
    exp.loc[pending, "status"] = "Pending"
    exp.loc[renewed, "status"] = "Renewed"
    return exp


# --------------------------------------------------------------------------- summaries
def summarize(df: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    """Renewal rate by count and premium. Pending policies are excluded from the rate."""
    closed = df[df["status"] != "Pending"]
    g = closed.groupby(by, observed=True)
    out = pd.DataFrame({
        "Expired": g.size(),
        "Renewed": g.apply(lambda x: (x["status"] == "Renewed").sum(), include_groups=False),
        "Premium expired": g["premium"].sum(),
        "Premium renewed": g.apply(lambda x: x.loc[x["status"] == "Renewed", "premium"].sum(), include_groups=False),
    })
    out["Pending"] = df[df["status"] == "Pending"].groupby(by, observed=True).size().reindex(out.index).fillna(0).astype(int)
    out = out.reset_index()
    total = {c: ("Total" if c == by[0] else "") for c in by}
    total.update(out[["Expired", "Renewed", "Premium expired", "Premium renewed", "Pending"]].sum().to_dict())
    out = pd.concat([out, pd.DataFrame([total])], ignore_index=True)
    out["Not renewed"] = out["Expired"] - out["Renewed"]
    out["Renewal rate"] = out["Renewed"] / out["Expired"]
    out["Premium retention"] = out["Premium renewed"] / out["Premium expired"]
    return out[by + ["Expired", "Renewed", "Not renewed", "Pending", "Renewal rate",
                     "Premium expired", "Premium renewed", "Premium retention"]]


def reasons_summary(df: pd.DataFrame, reasons: pd.DataFrame) -> pd.DataFrame:
    lost = df[df["status"] == "Not renewed"].merge(reasons, on="policy_number", how="left")
    lost["reason"] = lost["reason"].fillna("No answer from customer")
    out = lost.groupby("reason").agg(Policies=("policy_number", "size"), Premium=("premium", "sum"))
    out["Share"] = out["Policies"] / out["Policies"].sum()
    return out.sort_values("Policies", ascending=False).reset_index().rename(columns={"reason": "Reason"})


# --------------------------------------------------------------------------- excel output
HEAD_FILL = PatternFill("solid", start_color="1F4E78")
HEAD_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", start_color="E2EFDA")


def write_sheet(wb: Workbook, title: str, df: pd.DataFrame) -> None:
    ws = wb.create_sheet(title)
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v) for v in row])
    for c, col in enumerate(df.columns, 1):
        h = ws.cell(1, c)
        h.fill, h.font, h.alignment = HEAD_FILL, HEAD_FONT, Alignment(horizontal="center", wrap_text=True)
        fmt = ("0.0%" if col in ("Renewal rate", "Premium retention", "Share")
               else "#,##0.00" if "remium" in col
               else "yyyy-mm-dd" if col.endswith("_date") else None)
        if fmt:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, c).number_format = fmt
        ws.column_dimensions[get_column_letter(c)].width = max(12, min(30, len(str(col)) + 4))
    if str(ws.cell(ws.max_row, 1).value) == "Total":
        for c in range(1, ws.max_column + 1):
            ws.cell(ws.max_row, c).fill = TOTAL_FILL
            ws.cell(ws.max_row, c).font = Font(bold=True)
    ws.freeze_panes = "B2"


def build_report(df: pd.DataFrame, reasons: pd.DataFrame, out: Path) -> None:
    df = df.assign(expiry_month=df["end_date"].dt.to_period("M").astype(str))
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "By channel", summarize(df, ["channel"]))
    write_sheet(wb, "By client type", summarize(df, ["client_type"]))
    write_sheet(wb, "By department", summarize(df, ["channel", "department"]))
    write_sheet(wb, "By expiry month", summarize(df, ["expiry_month"]))
    write_sheet(wb, "Non-renewal reasons", reasons_summary(df, reasons))
    write_sheet(wb, "Policies", df.drop(columns="expiry_month"))
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


def split_by_manager(df: pd.DataFrame, out_dir: Path, password: str | None) -> int:
    """One workbook per manager with open policies and an editable Status dropdown."""
    out_dir.mkdir(parents=True, exist_ok=True)
    open_ = df[df["status"] != "Renewed"]
    cols = ["policy_number", "client_type", "brand", "plate_number", "end_date", "premium", "status"]
    for manager, part in open_.groupby("manager"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Follow-up"
        ws.append(["Status", "Comment"] + cols)
        for row in part[cols].itertuples(index=False):
            ws.append(["Wait", None] + [v.to_pydatetime() if isinstance(v, pd.Timestamp) else v for v in row])
        for c in range(1, ws.max_column + 1):
            h = ws.cell(1, c)
            h.fill, h.font = HEAD_FILL, HEAD_FONT
            ws.column_dimensions[get_column_letter(c)].width = 18
        for r in range(2, ws.max_row + 1):
            ws.cell(r, 7).number_format = "yyyy-mm-dd"
            ws.cell(r, 8).number_format = "#,##0.00"

        dv = DataValidation(type="list", formula1='"Positive,Negative,Wait"', allow_blank=False,
                            showErrorMessage=True, errorTitle="Invalid value",
                            error="Please choose Positive, Negative or Wait")
        ws.add_data_validation(dv)
        dv.add(f"A2:A{ws.max_row}")

        if password:  # lock everything except Status and Comment
            for r in range(2, ws.max_row + 1):
                for c in (1, 2):
                    ws.cell(r, c).protection = Protection(locked=False)
            ws.protection.sheet = True
            ws.protection.password = password
            ws.protection.autoFilter = False
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "C2"
        safe = re.sub(r"[^\w\- ]", "_", str(manager)).strip()
        wb.save(out_dir / f"{safe}.xlsx")
    return open_["manager"].nunique()


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", type=Path, default=BASE / "data" / "policies.xlsx")
    ap.add_argument("--as-of", type=pd.Timestamp, default=pd.Timestamp.today().normalize())
    ap.add_argument("--out", type=Path, default=BASE / "output" / "renewal_report.xlsx")
    ap.add_argument("--split", action="store_true", help="also create one follow-up file per manager")
    ap.add_argument("--password", default=os.getenv("RENEWAL_FILE_PASSWORD"),
                    help="sheet protection password for the split files (or env RENEWAL_FILE_PASSWORD)")
    a = ap.parse_args()

    expiring = pd.read_excel(a.input, sheet_name="Expiring", parse_dates=["start_date", "end_date"])
    new_sales = pd.read_excel(a.input, sheet_name="NewSales", parse_dates=["start_date"])
    reasons = pd.read_excel(a.input, sheet_name="Reasons")
    expiring = expiring[expiring["end_date"] <= a.as_of]

    df = match_renewals(expiring, new_sales, a.as_of)
    build_report(df, reasons, a.out)

    counts = df["status"].value_counts()
    closed = counts.get("Renewed", 0) + counts.get("Not renewed", 0)
    print(f"As of {a.as_of:%Y-%m-%d}: {len(df)} expired policies | renewed {counts.get('Renewed', 0)} | "
          f"not renewed {counts.get('Not renewed', 0)} | pending {counts.get('Pending', 0)} | "
          f"renewal rate {counts.get('Renewed', 0) / closed:.1%}")
    print(f"Report saved to {a.out.resolve()}")

    if a.split:
        n = split_by_manager(df, a.out.parent / "follow_up", a.password)
        print(f"{n} manager follow-up files saved to {(a.out.parent / 'follow_up').resolve()}"
              + ("" if a.password else " (unprotected: no --password given)"))


if __name__ == "__main__":
    main()
