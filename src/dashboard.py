"""
Step 8 - Interactive web dashboard (docs/index.html).

Writes ONE self-contained HTML file (no internet, no server needed) with all
results embedded as JSON. Put the repo on GitHub, turn on GitHub Pages for
the /docs folder, and anyone can open the dashboard from a link - no Python.

The disruption simulator runs in the browser using the same formulas as
src/scenario.py, so visitors can try their own what-if scenarios.
"""
import json

import numpy as np
import pandas as pd

import config
from inventory import SEGMENT_STRATEGY
from scenario import GROSS_MARKUP

DOCS_DIR = config.ROOT / "docs"
REPO_URL = "https://github.com/D2497/supply-chain-inventory-risk-analytics"


def _clean(obj):
    """Make numpy / pandas values JSON-safe."""
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_clean(v) for v in obj]
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating, float)):
        return None if np.isnan(obj) else round(float(obj), 3)
    if isinstance(obj, (pd.Timestamp,)):
        return obj.strftime("%Y-%m-%d")
    return obj


def build_payload(r: dict) -> dict:
    p = r["policy"].sort_values("annual_spend_cad", ascending=False)
    sc = r["scorecard"]
    k = r["kpis"]
    demand = r["demand"].sort_values("week_start")
    fc = r["forecasts"]

    skus = []
    for row in p.itertuples():
        hist = demand.loc[demand["sku"] == row.sku, "units"].tolist()
        future = fc.loc[fc["sku"] == row.sku, "forecast_units"].tolist()
        skus.append({
            "sku": row.sku, "name": row.product_name, "cat": row.category,
            "sup": row.supplier_id, "abc": row.abc_class, "xyz": row.xyz_class,
            "status": row.status, "onHand": row.on_hand_units, "onOrder": row.on_order_units,
            "pos": row.inventory_position, "ss": row.safety_stock, "rop": row.reorder_point,
            "eoq": row.eoq, "orderQty": row.suggested_order_qty, "d": row.forecast_weekly_avg,
            "lt": row.lead_time_mean_wks, "cost": row.unit_cost_cad, "spend": row.annual_spend_cad,
            "cover": row.weeks_of_cover, "method": row.best_method, "wape": row.holdout_wape,
            "hist": hist, "fc": future,
        })

    suppliers = [{
        "id": s.supplier_id, "name": s.supplier_name, "country": s.country,
        "score": s.risk_score, "tier": s.risk_tier, "otd": s.on_time_rate,
        "lt": s.lead_time_mean_wks, "ltcv": s.lead_time_cv, "defect": s.defect_rate,
        "share": s.share_of_spend, "driver": s.main_risk_driver, "action": s.recommended_action,
        "skus": int((p["supplier_id"] == s.supplier_id).sum()),
    } for s in sc.itertuples()]

    matrix = pd.crosstab(p["abc_class"], p["xyz_class"]).reindex(
        index=["A", "B", "C"], columns=["X", "Y", "Z"], fill_value=0)

    return _clean({
        "firstWeek": demand["week_start"].min(),
        "markup": GROSS_MARKUP,
        "serviceLevel": config.SERVICE_LEVEL,
        "kpi": {
            "spend": k["Annual spend (CAD)"], "skus": k["SKUs analysed"],
            "suppliers": k["Suppliers analysed"], "aCount": k["A-class SKUs"],
            "aShare": k["A-class share of spend"],
            "aSkuShare": k["A-class SKUs"] / k["SKUs analysed"],
            "wape": k["Avg forecast WAPE (best model)"], "wapeNaive": k["Avg forecast WAPE (naive)"],
            "eoqSave": k["Annual saving from EOQ (CAD)"], "eoqPct": k["EOQ saving vs current ordering cost"],
            "excess": k["Excess inventory value (CAD)"], "otd": k["Overall on-time delivery"],
            "orderNow": k["SKUs to order now (incl. critical)"], "critical": k["SKUs critical (below safety stock)"],
            "over": k["SKUs overstocked"], "highRisk": k["High-risk suppliers"],
        },
        "skus": skus,
        "suppliers": suppliers,
        "matrix": {a: {x: int(matrix.loc[a, x]) for x in "XYZ"} for a in "ABC"},
        "strategy": SEGMENT_STRATEGY,
        "repo": REPO_URL,
        "source": r["source"],
        "landedCountry": r["landed_by_country"].to_dict("records"),
        "landed": r["landed"][["sku", "product_name", "current_country", "nearshore_country",
                               "current_lead_time_wks", "nearshore_lead_time_wks", "annual_saving_cad",
                               "saving_pct", "inventory_released_cad", "recommendation"]].to_dict("records"),
        "landedSummary": r["landed_summary"],
    })


def write_dashboard(r: dict) -> None:
    DOCS_DIR.mkdir(exist_ok=True)
    template = (config.ROOT / "src" / "dashboard_template.html").read_text(encoding="utf-8")
    data = json.dumps(build_payload(r), separators=(",", ":"))
    html = template.replace("/*__DATA__*/null", data)
    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")
