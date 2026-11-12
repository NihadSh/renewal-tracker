"""Generate synthetic motor-insurance policies for the renewal tracker demo.

Creates two sheets in data/policies.xlsx:
  * Expiring  - policies that expire in the analysed period
  * NewSales  - policies sold in the period (some of them are renewals)
All names, plates and amounts are random and fictional.
"""
from pathlib import Path

import numpy as np
import pandas as pd

CHANNELS = {
    "Agents": ["Agent Sales 1", "Agent Sales 2"],
    "Regional": ["Regional Sales North"],
    "Key Accounts": ["Key Accounts"],
    "SME": ["Small Business", "Corporate Sales"],
    "Partners": ["Car Dealers"],
}
MANAGERS = {d: f"Manager {chr(65 + i)}" for i, d in
            enumerate(d for ds in CHANNELS.values() for d in ds)}
BRANDS = ["Toyota", "Hyundai", "Kia", "Mercedes-Benz", "BMW", "Chevrolet", "Changan", "Lexus"]
REASONS = ["Customer declined", "Car sold", "Price too high", "Unhappy with claim handling",
           "Renewed with another insurer", None]


def main(out: Path, n: int = 1500, seed: int = 7) -> None:
    rng = np.random.default_rng(seed)
    out.parent.mkdir(parents=True, exist_ok=True)

    channel = rng.choice(list(CHANNELS), n, p=[0.35, 0.15, 0.1, 0.25, 0.15])
    dept = [rng.choice(CHANNELS[c]) for c in channel]
    end = pd.Timestamp("2026-01-01") + pd.to_timedelta(rng.integers(0, 151, n), unit="D")

    expiring = pd.DataFrame({
        "policy_number": [f"MTR-25-{i:06d}" for i in range(1, n + 1)],
        "customer_id": rng.integers(100_000, 999_999, n),
        "plate_number": [f"{rng.integers(10, 99)}-{chr(65 + rng.integers(0, 26))}{chr(65 + rng.integers(0, 26))}-{rng.integers(100, 999)}" for _ in range(n)],
        "client_type": rng.choice(["Individual", "Corporate"], n, p=[0.7, 0.3]),
        "channel": channel,
        "department": dept,
        "manager": [MANAGERS[d] for d in dept],
        "brand": rng.choice(BRANDS, n),
        "start_date": end - pd.DateOffset(years=1),
        "end_date": end,
        "premium": np.round(rng.lognormal(6.6, 0.5, n), 2),
    })

    # ~68% of expiring policies get renewed (same customer + same car).
    renewed = expiring.sample(frac=0.68, random_state=seed)
    renewals = pd.DataFrame({
        "policy_number": [f"MTR-26-{i:06d}" for i in range(1, len(renewed) + 1)],
        "customer_id": renewed["customer_id"].values,
        "plate_number": renewed["plate_number"].values,
        "start_date": renewed["end_date"].values + pd.to_timedelta(rng.integers(-5, 10, len(renewed)), unit="D"),
        "premium": np.round(renewed["premium"].values * rng.uniform(0.9, 1.15, len(renewed)), 2),
    })
    # Plus brand-new business that is not a renewal.
    k = 400
    fresh = pd.DataFrame({
        "policy_number": [f"MTR-26-{i:06d}" for i in range(len(renewed) + 1, len(renewed) + k + 1)],
        "customer_id": rng.integers(100_000, 999_999, k),
        "plate_number": [f"{rng.integers(10, 99)}-ZZ-{rng.integers(100, 999)}" for _ in range(k)],
        "start_date": pd.Timestamp("2026-01-01") + pd.to_timedelta(rng.integers(0, 151, k), unit="D"),
        "premium": np.round(rng.lognormal(6.6, 0.5, k), 2),
    })
    new_sales = pd.concat([renewals, fresh]).sample(frac=1, random_state=seed).reset_index(drop=True)

    # Reasons collected by sales staff for some non-renewed policies.
    lost = expiring[~expiring.index.isin(renewed.index)]
    reasons = pd.DataFrame({"policy_number": lost["policy_number"],
                            "reason": rng.choice(REASONS, len(lost))}).dropna()

    with pd.ExcelWriter(out) as xw:
        expiring.to_excel(xw, sheet_name="Expiring", index=False)
        new_sales.to_excel(xw, sheet_name="NewSales", index=False)
        reasons.to_excel(xw, sheet_name="Reasons", index=False)
    print(f"Sample data written to {out.resolve()}")


if __name__ == "__main__":
    main(Path(__file__).parent / "data" / "policies.xlsx")
