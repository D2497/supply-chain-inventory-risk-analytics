"""
Run the full pipeline:

    python src/main.py

1. Generate data           (generate_data.py)
2. Load into SQLite + SQL  (database.py, sql/analysis_queries.sql)
3. Forecast demand         (forecasting.py)
4. Optimize inventory      (inventory.py)
5. Score supplier risk     (supplier_risk.py)
6. Disruption what-if      (scenario.py)
7. Charts + Excel report + key findings
8. Interactive web dashboard  (dashboard.py -> docs/index.html)
"""
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import build_excel
import charts
import config
import dashboard
import database
import forecasting
import generate_data
import load_kaggle_dataco
import inventory
import landed_cost
import scenario
import supplier_risk


def run_pipeline(write_outputs: bool = True, verbose: bool = True) -> dict:
    log = print if verbose else (lambda *a, **k: None)

    source = config.DATA_SOURCE
    if source == "auto":
        source = "kaggle" if config.KAGGLE_DATACO_FILE.exists() else "simulated"
    if source == "kaggle":
        log("1/7 Loading Kaggle DataCo data ...")
        source_info = load_kaggle_dataco.main()
    else:
        log("1/7 Generating simulated data ...")
        source_info = generate_data.main()

    log("2/7 Building SQLite database and running SQL ...")
    database.build_database()
    performance = database.run_query("supplier_delivery_performance")
    lead_times = database.run_query("supplier_lead_times")
    exposure = database.run_query("single_source_exposure")
    category_spend = database.run_query("annual_spend_by_category")
    late_trend = database.run_query("monthly_late_trend")

    products = pd.read_csv(config.DATA_DIR / "products.csv")
    demand = pd.read_csv(config.DATA_DIR / "weekly_demand.csv", parse_dates=["week_start"])
    stock = pd.read_csv(config.DATA_DIR / "inventory_snapshot.csv")

    log("3/7 Forecasting demand (5 methods per SKU) ...")
    accuracy, chosen, forecasts = forecasting.evaluate_and_forecast(demand)

    log("4/7 Calculating ABC-XYZ, safety stock, ROP, EOQ ...")
    lt_stats = inventory.supplier_lead_time_stats(lead_times)
    classified = inventory.abc_xyz(demand, products)
    policy = inventory.build_policy(classified, chosen, lt_stats, stock)

    log("5/7 Scoring supplier risk ...")
    scorecard = supplier_risk.score_suppliers(performance, lt_stats, exposure)

    log("6/7 Running disruption scenario on the highest-risk supplier ...")
    worst = scorecard.iloc[0]["supplier_id"]
    disruption = scenario.simulate_disruption(policy, worst, delay_weeks=4)
    disruption_text = scenario.summarize(disruption, worst, 4)

    log("   + Landed cost: offshore vs nearshore sourcing ...")
    landed = landed_cost.compare(policy, scorecard)
    landed_by_country = landed_cost.country_breakdown(scorecard)
    landed_cost.write_csv(landed, landed_by_country)

    results = dict(source=source_info, products=products, demand=demand, performance=performance,
                   category_spend=category_spend, late_trend=late_trend,
                   accuracy=accuracy, chosen=chosen, forecasts=forecasts,
                   classified=classified, policy=policy, scorecard=scorecard,
                   disruption=disruption, disruption_text=disruption_text,
                   landed=landed, landed_by_country=landed_by_country,
                   landed_summary=landed_cost.summary(landed))
    results["kpis"] = compute_kpis(results)

    if write_outputs:
        log("7/7 Writing charts, Excel report and key findings ...")
        make_charts(results)
        write_excel(results)
        write_findings(results)
        dashboard.write_dashboard(results)
        build_excel.build(results["source"]["note"])
        log("\nDone. Open: docs/index.html (dashboard), excel/Supply_Chain_Analytics.xlsx (live Excel model), reports/")
    return results


# ------------------------------------------------------------------
def compute_kpis(r: dict) -> dict:
    p, c, sc = r["policy"], r["chosen"], r["scorecard"]
    method_wins = c["best_method"].value_counts()
    total_spend = p["annual_spend_cad"].sum()
    current_cost = p["current_cycle_cost_cad"].sum()
    return {
        "SKUs analysed": len(p),
        "Suppliers analysed": len(sc),
        "Annual spend (CAD)": round(total_spend, 0),
        "A-class SKUs": int((p["abc_class"] == "A").sum()),
        "A-class share of spend": round(p.loc[p["abc_class"] == "A", "annual_spend_cad"].sum() / total_spend, 3),
        "Avg forecast WAPE (best model)": round(c["holdout_wape"].mean(), 3),
        "Avg forecast WAPE (naive)": round(c["naive_wape"].mean(), 3),
        "Most common winning model": method_wins.index[0],
        "SKUs to order now (incl. critical)": int(p["status"].isin(["Order now", "Critical - expedite"]).sum()),
        "SKUs critical (below safety stock)": int((p["status"] == "Critical - expedite").sum()),
        "SKUs overstocked": int((p["status"] == "Overstock").sum()),
        "Excess inventory value (CAD)": round(p["excess_value_cad"].sum(), 0),
        "Safety stock investment (CAD)": round(p["safety_stock_value_cad"].sum(), 0),
        "Annual saving from EOQ (CAD)": round(p["annual_saving_cad"].sum(), 0),
        "EOQ saving vs current ordering cost": round(p["annual_saving_cad"].sum() / current_cost, 3),
        "Overall on-time delivery": round(
            (r["performance"]["on_time_rate"] * r["performance"]["po_count"]).sum()
            / r["performance"]["po_count"].sum(), 3),
        "High-risk suppliers": int((sc["risk_tier"] == "High").sum()),
        "Highest-risk supplier": f"{sc.iloc[0]['supplier_name']} ({sc.iloc[0]['risk_score']:.0f}/100)",
        "Offshore SKUs reviewed for nearshoring": r["landed_summary"]["offshore_skus"],
        "SKUs cheaper to source nearshore": r["landed_summary"]["cheaper_nearshore"],
        "Nearshore candidates (incl. <3% premium)": r["landed_summary"]["nearshore_candidates"],
        "Annual saving from cheaper nearshore moves (CAD)": r["landed_summary"]["saving_if_cheaper_moved"],
        "Inventory released by nearshoring candidates (CAD)": r["landed_summary"]["inventory_released_if_moved"],
        "Avg lead-time cut for candidates (weeks)": r["landed_summary"]["avg_lead_time_cut_wks"],
    }


def make_charts(r: dict):
    charts.pareto(r["classified"])
    charts.abc_xyz_matrix(r["classified"])
    charts.inventory_status(r["policy"])
    charts.supplier_risk(r["scorecard"])
    charts.otd_vs_variability(r["scorecard"])
    # example: highest-spend PPE SKU (shows the flu-season pattern clearly)
    p = r["policy"].sort_values("annual_spend_cad", ascending=False)
    ppe = p[p["category"] == "PPE"]
    sku = (ppe if len(ppe) else p).iloc[0]["sku"]
    charts.forecast_example(r["demand"], r["forecasts"], r["chosen"], r["products"], sku)
    charts.landed_cost(r["landed_by_country"])


# ------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1D5AA6")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def _format_sheet(ws, money_cols=(), pct_cols=()):
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    headers = [c.value for c in ws[1]]
    for idx, h in enumerate(headers, start=1):
        letter = get_column_letter(idx)
        width = max(len(str(h)), *(len(str(c.value)) for c in ws[letter][1:200] if c.value is not None), 8)
        ws.column_dimensions[letter].width = min(width + 2, 60)
        fmt = "#,##0.00" if h in money_cols else "0.0%" if h in pct_cols else None
        if fmt:
            for c in ws[letter][1:]:
                c.number_format = fmt


def write_excel(r: dict):
    p = r["policy"]
    out = config.REPORT_DIR / "supply_chain_report.xlsx"
    kpi = pd.DataFrame(list(r["kpis"].items()), columns=["KPI", "Value"])

    policy_cols = ["sku", "product_name", "category", "supplier_id", "segment", "unit_cost_cad",
                   "forecast_weekly_avg", "lead_time_mean_wks", "safety_stock", "reorder_point",
                   "eoq", "on_hand_units", "on_order_units", "inventory_position",
                   "weeks_of_cover", "status", "suggested_order_qty", "annual_saving_cad", "strategy"]
    alerts = p[p["status"].isin(["Critical - expedite", "Order now"])].sort_values(
        ["status", "abc_class"])[["sku", "product_name", "supplier_id", "abc_class", "status",
                                  "inventory_position", "safety_stock", "reorder_point",
                                  "suggested_order_qty", "lead_time_mean_wks"]]
    abc_cols = ["sku", "product_name", "category", "annual_units", "annual_spend_cad",
                "cum_spend_pct", "abc_class", "demand_cv", "xyz_class", "segment"]
    sc_cols = ["supplier_id", "supplier_name", "country", "po_count", "total_spend_cad",
               "share_of_spend", "on_time_rate", "avg_days_late", "promised_lead_time_wks",
               "lead_time_mean_wks", "lead_time_cv", "fill_rate", "defect_rate",
               "country_risk_1to5", "risk_score", "risk_tier", "main_risk_driver",
               "recommended_action"]
    acc_pivot = r["accuracy"].pivot(index="sku", columns="method", values="wape").round(4).reset_index()
    acc_pivot = acc_pivot.merge(r["chosen"][["sku", "best_method"]], on="sku")
    disruption = r["disruption"].copy()
    disruption.insert(0, "scenario", r["disruption_text"])

    sheets = [
        ("Summary", kpi, {}, {}),
        ("Reorder Alerts", alerts, (), ()),
        ("Inventory Policy", p[policy_cols], ("unit_cost_cad", "annual_saving_cad"), ()),
        ("ABC-XYZ", p.sort_values("annual_spend_cad", ascending=False)[abc_cols],
         ("annual_spend_cad",), ("cum_spend_pct",)),
        ("Supplier Scorecard", r["scorecard"][sc_cols], ("total_spend_cad",),
         ("share_of_spend", "on_time_rate", "fill_rate", "defect_rate")),
        ("Forecast Accuracy", acc_pivot, (), [c for c in acc_pivot.columns if c not in ("sku", "best_method")]),
        ("12-Week Forecast", r["forecasts"], (), ()),
        ("Disruption Scenario", disruption, ("revenue_at_risk_cad", "extra_buffer_cost_cad"), ()),
        ("Spend by Category", r["category_spend"], ("spend_52w_cad",), ()),
        ("Landed Cost by SKU", r["landed"], ("current_landed_unit_cad", "nearshore_landed_unit_cad",
                                             "current_total_cad", "nearshore_total_cad", "annual_saving_cad",
                                             "inventory_released_cad"), ("saving_pct",)),
        ("Landed Cost by Country", r["landed_by_country"], (), ("on_time_rate",)),
    ]
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for name, df, money, pct in sheets:
            df.to_excel(xw, sheet_name=name, index=False)
            _format_sheet(xw.sheets[name], money, pct)


def write_findings(r: dict):
    k, sc, p = r["kpis"], r["scorecard"], r["policy"]
    top_alerts = p[p["status"] == "Critical - expedite"].sort_values("annual_spend_cad", ascending=False).head(5)
    lines = [
        "# Key findings (auto-generated by `python src/main.py`)",
        "",
        f"> **Data source:** {r['source']['note']}",
        "",
        "## Headline numbers",
        "",
        "| KPI | Value |",
        "|---|---|",
    ]
    for key, val in k.items():
        if isinstance(val, float) and val < 1 and ("share" in key or "WAPE" in key or "delivery" in key or "saving vs" in key):
            val = f"{val:.1%}"
        elif isinstance(val, float):
            val = f"{val:,.0f}"
        lines.append(f"| {key} | {val} |")
    lines += [
        "",
        "## Supplier risk ranking",
        "",
        "| Supplier | Country | Risk score | Tier | Main driver | Action |",
        "|---|---|---|---|---|---|",
    ]
    for s in sc.itertuples():
        lines.append(f"| {s.supplier_name} | {s.country} | {s.risk_score:.0f} | {s.risk_tier} | "
                     f"{s.main_risk_driver} | {s.recommended_action} |")
    lines += ["", "## Critical SKUs to expedite (highest spend first)", ""]
    if top_alerts.empty:
        lines.append("None.")
    for a in top_alerts.itertuples():
        lines.append(f"- **{a.sku} {a.product_name}** - position {a.inventory_position:.0f} units vs "
                     f"safety stock {a.safety_stock:.0f}; order {a.suggested_order_qty:.0f} units "
                     f"from {a.supplier_id}")
    lines += ["", "## Disruption scenario", "", r["disruption_text"], ""]
    ls, ld = r["landed_summary"], r["landed"]
    lines += ["## Logistics & landed cost (offshore vs nearshore)", "",
              f"Of {ls['offshore_skus']} SKUs bought offshore, {ls['cheaper_nearshore']} are cheaper to source nearshore "
              f"(CAD {ls['saving_if_cheaper_moved']:,.0f}/yr saving), and {ls['nearshore_candidates']} are worth moving "
              f"if a premium of up to 3% is accepted for resilience. Moving them releases CAD "
              f"{ls['inventory_released_if_moved']:,.0f} of safety-stock and in-transit inventory and cuts lead time "
              f"by {ls['avg_lead_time_cut_wks']} weeks on average.", "",
              "| SKU | Product | Now | Nearshore | Saving/yr | Inventory released | Recommendation |",
              "|---|---|---|---|---|---|---|"]
    for x in ld.head(10).itertuples():
        lines.append(f"| {x.sku} | {x.product_name} | {x.current_country} | {x.nearshore_country} | "
                     f"{x.annual_saving_cad:,.0f} | {x.inventory_released_cad:,.0f} | {x.recommendation} |")
    lines.append("")
    (config.REPORT_DIR / "key_findings.md").write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    run_pipeline()
