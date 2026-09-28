"""
Step 4 - Inventory optimization.

For every SKU this module calculates:
  * ABC class  - by annual spend (A = the few SKUs that matter most)
  * XYZ class  - by demand variability (X = stable, Z = erratic)
  * Safety stock, Reorder point (ROP) and Economic Order Quantity (EOQ)
  * Current status: Order now / Critical / Healthy / Overstock
  * Cost saving of EOQ versus the current "order every 4 weeks" habit
"""
from statistics import NormalDist

import numpy as np
import pandas as pd

import config


# ------------------------------------------------------------------
# Core formulas (kept as small pure functions so they can be unit-tested)
# ------------------------------------------------------------------
def z_score(service_level: float) -> float:
    """95% service level -> z = 1.645"""
    return NormalDist().inv_cdf(service_level)


def safety_stock(z, demand_mean, demand_std, lt_mean, lt_std) -> float:
    """
    Safety stock when BOTH demand and lead time vary:
        SS = z * sqrt( LT * sigma_d^2  +  d^2 * sigma_LT^2 )
    d, sigma_d in units/week; LT, sigma_LT in weeks.
    """
    return z * np.sqrt(lt_mean * demand_std ** 2 + demand_mean ** 2 * lt_std ** 2)


def reorder_point(demand_mean, lt_mean, ss) -> float:
    """ROP = demand during lead time + safety stock"""
    return demand_mean * lt_mean + ss


def eoq(annual_demand, order_cost, unit_holding_cost) -> float:
    """Economic Order Quantity = sqrt(2DS / H)"""
    if annual_demand <= 0 or unit_holding_cost <= 0:
        return 0.0
    return np.sqrt(2 * annual_demand * order_cost / unit_holding_cost)


def annual_cycle_cost(annual_demand, order_qty, order_cost, unit_holding_cost) -> float:
    """Ordering cost + cycle-stock holding cost for a given order quantity."""
    if order_qty <= 0:
        return 0.0
    return annual_demand / order_qty * order_cost + order_qty / 2 * unit_holding_cost


# ------------------------------------------------------------------
# Classification
# ------------------------------------------------------------------
def abc_xyz(demand: pd.DataFrame, products: pd.DataFrame) -> pd.DataFrame:
    last_52 = demand[demand["week_start"] > demand["week_start"].max() - pd.Timedelta(weeks=52)]
    stats = last_52.groupby("sku")["units"].agg(annual_units="sum", weekly_mean="mean",
                                                 weekly_std="std").reset_index()
    df = products.merge(stats, on="sku")
    df["annual_spend_cad"] = df["annual_units"] * df["unit_cost_cad"]

    # ABC: sort by spend, take running % of total
    df = df.sort_values("annual_spend_cad", ascending=False).reset_index(drop=True)
    df["cum_spend_pct"] = df["annual_spend_cad"].cumsum() / df["annual_spend_cad"].sum()
    # use the share BEFORE this SKU so the SKU that crosses 80% is still an A
    prev = df["cum_spend_pct"].shift(fill_value=0)
    df["abc_class"] = np.select([prev < config.ABC_A_CUTOFF, prev < config.ABC_B_CUTOFF],
                                ["A", "B"], default="C")

    # XYZ: coefficient of variation = std / mean
    df["demand_cv"] = (df["weekly_std"] / df["weekly_mean"]).round(3)
    df["xyz_class"] = np.select([df["demand_cv"] <= config.XYZ_X_MAX_CV,
                                 df["demand_cv"] <= config.XYZ_Y_MAX_CV],
                                ["X", "Y"], default="Z")
    df["segment"] = df["abc_class"] + df["xyz_class"]
    return df


SEGMENT_STRATEGY = {
    "AX": "Tight control, frequent review, low safety stock, consider VMI/automation",
    "AY": "Weekly review, forecast-driven replenishment, moderate safety stock",
    "AZ": "Highest attention: collaborate with customers, dual-source, buffer stock",
    "BX": "Automated reorder point, periodic review",
    "BY": "Standard reorder point policy, monthly review",
    "BZ": "Higher safety stock or make-to-order where possible",
    "CX": "Automate fully, order in larger batches",
    "CY": "Automate, larger batches, low review effort",
    "CZ": "Review if SKU is needed; consider dropping or ordering on demand",
}


# ------------------------------------------------------------------
# Policy + status
# ------------------------------------------------------------------
def build_policy(classified: pd.DataFrame, forecast_stats: pd.DataFrame,
                 supplier_lt: pd.DataFrame, inventory: pd.DataFrame) -> pd.DataFrame:
    z = z_score(config.SERVICE_LEVEL)
    df = (classified
          .merge(forecast_stats, on="sku")
          .merge(supplier_lt, on="supplier_id")
          .merge(inventory, on="sku"))

    d = df["forecast_weekly_avg"]
    sigma_d = df["forecast_rmse_units"]          # forecast error, not raw demand std
    lt, sigma_lt = df["lead_time_mean_wks"], df["lead_time_std_wks"]
    h = df["unit_cost_cad"] * config.HOLDING_COST_RATE
    annual_d = d * config.WEEKS_PER_YEAR

    df["safety_stock"] = np.ceil(safety_stock(z, d, sigma_d, lt, sigma_lt))
    df["reorder_point"] = np.ceil(reorder_point(d, lt, df["safety_stock"]))
    df["eoq"] = np.ceil([eoq(a, config.ORDER_COST, hc) for a, hc in zip(annual_d, h)])
    df["max_stock"] = df["reorder_point"] + df["eoq"]

    df["inventory_position"] = df["on_hand_units"] + df["on_order_units"]
    df["weeks_of_cover"] = (df["on_hand_units"] / d).round(1)

    df["status"] = np.select(
        [df["inventory_position"] < df["safety_stock"],
         df["inventory_position"] <= df["reorder_point"],
         df["inventory_position"] > df["max_stock"]],
        ["Critical - expedite", "Order now", "Overstock"],
        default="Healthy")

    needs_order = df["status"].isin(["Critical - expedite", "Order now"])
    df["suggested_order_qty"] = np.where(
        needs_order, np.maximum(df["eoq"], df["max_stock"] - df["inventory_position"]), 0)
    df["excess_value_cad"] = np.where(
        df["status"] == "Overstock",
        (df["inventory_position"] - df["max_stock"]) * df["unit_cost_cad"], 0).round(2)
    df["safety_stock_value_cad"] = (df["safety_stock"] * df["unit_cost_cad"]).round(2)

    # Current habit = order ~4 weeks of demand each time. Compare with EOQ.
    current_q = d * 4
    df["current_cycle_cost_cad"] = [annual_cycle_cost(a, q, config.ORDER_COST, hc)
                                    for a, q, hc in zip(annual_d, current_q, h)]
    df["eoq_cycle_cost_cad"] = [annual_cycle_cost(a, q, config.ORDER_COST, hc)
                                for a, q, hc in zip(annual_d, df["eoq"], h)]
    df["annual_saving_cad"] = (df["current_cycle_cost_cad"] - df["eoq_cycle_cost_cad"]).round(2)
    df["strategy"] = df["segment"].map(SEGMENT_STRATEGY)
    return df


def supplier_lead_time_stats(lead_times: pd.DataFrame) -> pd.DataFrame:
    """Mean and std-dev of actual lead time (weeks) per supplier."""
    return (lead_times.groupby("supplier_id")["lead_time_wks"]
            .agg(lead_time_mean_wks="mean", lead_time_std_wks="std")
            .round(2).reset_index())
