"""
Step 5 - Supplier scorecard and risk score.

Each supplier gets a 0-100 risk score (higher = riskier) built from five
weighted factors (weights are in config.RISK_WEIGHTS):

  late_delivery          1 - on-time delivery rate
  lead_time_variability  coefficient of variation of actual lead time
  quality                defect rate
  country_risk           1-5 route/geopolitical score (assumption, see config)
  spend_dependency       share of total spend with this supplier

Each factor is scaled to 0-1 against a "clearly bad" threshold so the
score means the same thing even if you add or remove suppliers.
"""
import numpy as np
import pandas as pd

import config

# value at which a factor counts as fully risky (score = 1.0)
BAD_THRESHOLD = {
    "late_delivery": 0.35,          # 35% of orders late
    "lead_time_variability": 0.30,  # lead time CV of 30%
    "quality": 0.025,               # 2.5% defective units
    "spend_dependency": 0.30,       # 30% of all spend with one supplier
}


def score_suppliers(performance: pd.DataFrame, lt_stats: pd.DataFrame,
                    exposure: pd.DataFrame) -> pd.DataFrame:
    df = performance.merge(lt_stats, on="supplier_id").merge(
        exposure[["supplier_id", "skus_supplied", "share_of_spend"]], on="supplier_id")

    df["lead_time_cv"] = (df["lead_time_std_wks"] / df["lead_time_mean_wks"]).round(3)
    df["country_risk_1to5"] = df["country"].map(config.COUNTRY_RISK)

    factors = pd.DataFrame({
        "late_delivery": (1 - df["on_time_rate"]) / BAD_THRESHOLD["late_delivery"],
        "lead_time_variability": df["lead_time_cv"] / BAD_THRESHOLD["lead_time_variability"],
        "quality": df["defect_rate"] / BAD_THRESHOLD["quality"],
        "country_risk": (df["country_risk_1to5"] - 1) / 4,
        "spend_dependency": df["share_of_spend"] / BAD_THRESHOLD["spend_dependency"],
    }).clip(0, 1)

    for f in factors:
        df[f"f_{f}"] = factors[f].round(3)
    df["risk_score"] = (sum(factors[f] * w for f, w in config.RISK_WEIGHTS.items()) * 100).round(1)
    df["risk_tier"] = pd.cut(df["risk_score"], bins=[-1, 35, 60, 101],
                             labels=["Low", "Medium", "High"]).astype(str)

    # main driver = the factor contributing the most points
    contrib = pd.DataFrame({f: factors[f] * w for f, w in config.RISK_WEIGHTS.items()})
    df["main_risk_driver"] = contrib.idxmax(axis=1).str.replace("_", " ")
    df["recommended_action"] = df.apply(recommend, axis=1)
    return df.sort_values("risk_score", ascending=False).reset_index(drop=True)


def recommend(row) -> str:
    if row["risk_tier"] == "Low" and row["share_of_spend"] >= 0.25:
        return ("Reliable, but holds a large share of spend: keep as primary and "
                "qualify a backup supplier for its A-class SKUs")
    if row["risk_tier"] == "Low":
        return "Maintain; candidate for volume consolidation / preferred-supplier status"
    actions = {
        "late delivery": "Quarterly business review with OTD target; add late-delivery clause to contract",
        "lead time variability": "Raise safety stock on its SKUs; book freight earlier; track via ASN",
        "quality": "Supplier corrective action request (SCAR); increase incoming inspection",
        "country risk": "Qualify a nearshore/domestic second source for critical SKUs",
        "spend dependency": "Dual-source top SKUs to reduce concentration risk",
    }
    prefix = "Priority: " if row["risk_tier"] == "High" else ""
    return prefix + actions[row["main_risk_driver"]]
