"""
Step 7 - Logistics & landed cost: offshore vs nearshore sourcing.

Question it answers:
  "For each product we buy overseas, what does it REALLY cost once it lands in
   our warehouse, and would a nearshore supplier (Mexico, USA, Canada) be cheaper
   or worth a small premium?"

Total annual landed cost of a sourcing option =
      annual units x landed unit cost
    + cost of carrying the safety stock that option's lead time needs
    + cost of carrying pipeline (in-transit) inventory

  landed unit cost = price x (1 + freight% + duty% + brokerage%)
  safety stock     = z x sqrt(LT x sigma_d^2 + d^2 x sigma_LT^2)    (same formula as inventory.py)
  pipeline units   = weekly demand x lead time   (we own goods from shipment: FOB origin)
  carrying cost    = units x landed unit cost x holding cost rate

Prices: the same item costs more from a higher-cost country. We estimate a base
price from today's cost and the current country's price index, then apply each
alternative country's index (see config.LANDED_COST).
"""
import numpy as np
import pandas as pd

import config
from inventory import z_score

NEARSHORE = ("Mexico", "USA", "Canada")


def country_breakdown(scorecard: pd.DataFrame) -> pd.DataFrame:
    """Cost to land CAD 100 of product (priced at China level) from each country."""
    rows = []
    for country, (pi, fr, du, br, region) in config.LANDED_COST.items():
        price = 100 * pi
        sup = scorecard[scorecard["country"] == country]
        rows.append({
            "country": country, "region": region,
            "price": round(price, 2), "freight": round(price * fr, 2),
            "duty": round(price * du, 2), "brokerage": round(price * br, 2),
            "landed_total": round(price * (1 + fr + du + br), 2),
            "avg_lead_time_wks": round(float(sup["lead_time_mean_wks"].mean()), 1) if len(sup) else np.nan,
            "on_time_rate": round(float(sup["on_time_rate"].mean()), 3) if len(sup) else np.nan,
        })
    return pd.DataFrame(rows).sort_values("landed_total").reset_index(drop=True)


def _option_cost(d, sigma_d, base_price, country, lt, lt_std, z):
    pi, fr, du, br, _ = config.LANDED_COST[country]
    unit = base_price * pi * (1 + fr + du + br)
    ss = z * np.sqrt(lt * sigma_d ** 2 + d ** 2 * lt_std ** 2)
    pipeline = d * lt
    carrying = (ss + pipeline) * unit * config.HOLDING_COST_RATE
    annual_units = d * config.WEEKS_PER_YEAR
    return {"unit": unit, "ss": ss, "pipeline": pipeline, "carrying": carrying,
            "inventory_value": (ss + pipeline) * unit, "total": annual_units * unit + carrying}


def compare(policy: pd.DataFrame, scorecard: pd.DataFrame) -> pd.DataFrame:
    z = z_score(config.SERVICE_LEVEL)
    sc = scorecard.set_index("supplier_id")
    nearshore = sc[sc["country"].isin(NEARSHORE)]
    rows = []
    for p in policy.itertuples():
        cur = sc.loc[p.supplier_id]
        if cur["country"] in NEARSHORE or cur["country"] not in config.LANDED_COST:
            continue
        base = p.unit_cost_cad / config.LANDED_COST[cur["country"]][0]
        d, sd = p.forecast_weekly_avg, p.forecast_rmse_units
        now = _option_cost(d, sd, base, cur["country"], cur["lead_time_mean_wks"], cur["lead_time_std_wks"], z)
        best_id, best = None, None
        for sid, alt in nearshore.iterrows():
            opt = _option_cost(d, sd, base, alt["country"], alt["lead_time_mean_wks"], alt["lead_time_std_wks"], z)
            if best is None or opt["total"] < best["total"]:
                best_id, best = sid, opt
        saving = now["total"] - best["total"]
        pct = saving / now["total"] if now["total"] else 0
        if saving > 0:
            rec = "Nearshore: lower total cost"
        elif pct >= -config.NEARSHORE_PREMIUM_OK:
            rec = "Nearshore candidate: small premium for resilience"
        else:
            rec = "Keep offshore"
        rows.append({
            "sku": p.sku, "product_name": p.product_name, "abc_class": p.abc_class,
            "current_supplier": cur["supplier_name"], "current_country": cur["country"],
            "nearshore_supplier": sc.loc[best_id, "supplier_name"], "nearshore_country": sc.loc[best_id, "country"],
            "annual_units": round(d * config.WEEKS_PER_YEAR),
            "current_landed_unit_cad": round(now["unit"], 2), "nearshore_landed_unit_cad": round(best["unit"], 2),
            "current_lead_time_wks": round(cur["lead_time_mean_wks"], 1),
            "nearshore_lead_time_wks": round(sc.loc[best_id, "lead_time_mean_wks"], 1),
            "current_total_cad": round(now["total"], 2), "nearshore_total_cad": round(best["total"], 2),
            "annual_saving_cad": round(saving, 2), "saving_pct": round(pct, 4),
            "inventory_released_cad": round(now["inventory_value"] - best["inventory_value"], 2),
            "recommendation": rec,
        })
    return pd.DataFrame(rows).sort_values("annual_saving_cad", ascending=False).reset_index(drop=True)


def summary(comp: pd.DataFrame) -> dict:
    move = comp[comp["recommendation"] != "Keep offshore"]
    cheaper = comp[comp["annual_saving_cad"] > 0]
    return {
        "offshore_skus": len(comp),
        "cheaper_nearshore": len(cheaper),
        "nearshore_candidates": len(move),
        "saving_if_cheaper_moved": round(float(cheaper["annual_saving_cad"].sum()), 0),
        "inventory_released_if_moved": round(float(move["inventory_released_cad"].sum()), 0),
        "avg_lead_time_cut_wks": round(float((move["current_lead_time_wks"] - move["nearshore_lead_time_wks"]).mean()), 1)
        if len(move) else 0.0,
    }


def write_csv(comp: pd.DataFrame, by_country: pd.DataFrame):
    """Saved next to the other data files so Power BI can load them."""
    comp.to_csv(config.DATA_DIR / "landed_cost_by_sku.csv", index=False)
    by_country.to_csv(config.DATA_DIR / "landed_cost_by_country.csv", index=False)
