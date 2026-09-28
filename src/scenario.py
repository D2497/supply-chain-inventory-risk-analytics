"""
Step 6 - Disruption "what-if" simulator.

Question it answers:
  "If supplier X is delayed by N extra weeks (port strike, tariff hold,
   factory shutdown), which SKUs run out, when, and how much revenue is at risk?"

Run on its own:
    python src/scenario.py --supplier S08 --delay 4
"""
import argparse

import numpy as np
import pandas as pd

GROSS_MARKUP = 1.35   # assumed selling price = cost x 1.35 (illustrative)


def simulate_disruption(policy: pd.DataFrame, supplier_id: str, delay_weeks: float) -> pd.DataFrame:
    df = policy[policy["supplier_id"] == supplier_id].copy()
    if df.empty:
        raise ValueError(f"No SKUs found for supplier {supplier_id}")

    d = df["forecast_weekly_avg"].replace(0, np.nan)
    new_lt = df["lead_time_mean_wks"] + delay_weeks

    # Stock on hand + what is already on order (assume the open PO is also delayed)
    # calculate with full precision, round only for display (keeps the web
    # dashboard, which uses the same formula, in exact agreement)
    weeks_left = (df["on_hand_units"] + df["on_order_units"]) / d
    gap = (new_lt - weeks_left).clip(lower=0)
    df["new_lead_time_wks"] = new_lt.round(1)
    df["weeks_until_stockout"] = weeks_left.round(1)
    df["gap_weeks"] = gap.round(1)
    df["units_short"] = np.ceil(gap * d).fillna(0)
    df["revenue_at_risk_cad"] = (df["units_short"] * df["unit_cost_cad"] * GROSS_MARKUP).round(2)
    df["stockout_expected"] = gap > 0

    # Extra safety stock needed to stay protected if the delay becomes the new normal
    df["extra_buffer_units"] = np.ceil(d * delay_weeks).fillna(0)
    df["extra_buffer_cost_cad"] = (df["extra_buffer_units"] * df["unit_cost_cad"]).round(2)

    cols = ["sku", "product_name", "abc_class", "on_hand_units", "on_order_units",
            "forecast_weekly_avg", "lead_time_mean_wks", "new_lead_time_wks",
            "weeks_until_stockout", "gap_weeks", "stockout_expected", "units_short",
            "revenue_at_risk_cad", "extra_buffer_units", "extra_buffer_cost_cad"]
    return df[cols].sort_values("revenue_at_risk_cad", ascending=False).reset_index(drop=True)


def summarize(result: pd.DataFrame, supplier_id: str, delay: float) -> str:
    n_out = int(result["stockout_expected"].sum())
    return (f"Supplier {supplier_id} delayed +{delay:g} weeks -> "
            f"{n_out} of {len(result)} SKUs stock out, "
            f"CAD {result['revenue_at_risk_cad'].sum():,.0f} revenue at risk; "
            f"protecting against it needs CAD {result['extra_buffer_cost_cad'].sum():,.0f} "
            f"of extra buffer stock.")


if __name__ == "__main__":
    # Local import so `python src/scenario.py` works after main.py has run once
    from main import run_pipeline

    parser = argparse.ArgumentParser(description="Supplier disruption what-if")
    parser.add_argument("--supplier", default="S08", help="supplier id, e.g. S08")
    parser.add_argument("--delay", type=float, default=4, help="extra weeks of delay")
    args = parser.parse_args()

    results = run_pipeline(write_outputs=False, verbose=False)
    out = simulate_disruption(results["policy"], args.supplier, args.delay)
    print(summarize(out, args.supplier, args.delay))
    print(out.head(15).to_string(index=False))
