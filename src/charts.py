"""
Step 7 - Charts (saved as PNG in reports/charts).
Colors come from a colorblind-checked palette; status colors are only used
for status, and every colored mark also has a text label.
"""
import matplotlib

matplotlib.use("Agg")  # no screen needed
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import config

INK, INK_2, GRID, SURFACE = "#0b0b0b", "#52514e", "#e1e0d9", "#fcfcfb"
BLUE, ORANGE = "#2a78d6", "#eb6834"
STATUS = {"good": "#0ca30c", "warning": "#fab219", "serious": "#ec835a", "critical": "#d03b3b"}
SEQ_BLUES = ["#eef5fd", "#c6dcf6", "#8fbaec", "#4f92de", "#2a78d6", "#1d5aa6", "#123d73"]

plt.rcParams.update({
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
    "axes.edgecolor": GRID, "axes.labelcolor": INK_2, "xtick.color": INK_2, "ytick.color": INK_2,
    "text.color": INK, "font.size": 10, "axes.titlesize": 13, "axes.titleweight": "bold",
    "axes.titlelocation": "left", "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6,
})


def _save(fig, name):
    config.CHART_DIR.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(config.CHART_DIR / name, dpi=150)
    plt.close(fig)


def pareto(classified: pd.DataFrame):
    df = classified.sort_values("annual_spend_cad", ascending=False).reset_index(drop=True)
    x = (np.arange(1, len(df) + 1) / len(df)) * 100
    y = df["cum_spend_pct"] * 100
    n_a = (df["abc_class"] == "A").sum()

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(x, y, color=BLUE, linewidth=2)
    ax.fill_between(x, 0, y, color=BLUE, alpha=0.08)
    a_x = n_a / len(df) * 100
    ax.axvline(a_x, color=INK_2, linewidth=1, linestyle="--")
    ax.annotate(f"{n_a} A-class SKUs ({a_x:.0f}% of items)\n= {df.loc[n_a-1,'cum_spend_pct']*100:.0f}% of annual spend",
                xy=(a_x, df.loc[n_a - 1, "cum_spend_pct"] * 100), xytext=(a_x + 8, 55),
                arrowprops=dict(arrowstyle="-", color=INK_2), color=INK)
    ax.set(xlim=(0, 100), ylim=(0, 102), xlabel="% of SKUs (sorted by spend)",
           ylabel="Cumulative % of annual spend")
    ax.set_title("ABC analysis: a few SKUs drive most of the spend")
    _save(fig, "01_abc_pareto.png")


def abc_xyz_matrix(classified: pd.DataFrame):
    m = pd.crosstab(classified["abc_class"], classified["xyz_class"]).reindex(
        index=["A", "B", "C"], columns=["X", "Y", "Z"], fill_value=0)
    fig, ax = plt.subplots(figsize=(6, 4.6))
    ax.grid(False)
    vmax = max(m.values.max(), 1)
    for i, abc in enumerate(m.index):
        for j, xyz in enumerate(m.columns):
            v = m.loc[abc, xyz]
            step = int(round(v / vmax * (len(SEQ_BLUES) - 1)))
            ax.add_patch(plt.Rectangle((j, 2 - i), 0.96, 0.96, color=SEQ_BLUES[step]))
            ax.text(j + 0.48, 2 - i + 0.48, f"{abc}{xyz}\n{v} SKUs", ha="center", va="center",
                    color="#ffffff" if step >= 4 else INK, fontsize=11)
    ax.set(xlim=(0, 3), ylim=(0, 3), xticks=[0.48, 1.48, 2.48], yticks=[2.48, 1.48, 0.48])
    ax.set_xticklabels(["X (stable)", "Y (variable)", "Z (erratic)"])
    ax.set_yticklabels(["A (high $)", "B", "C (low $)"])
    for s in ax.spines.values():
        s.set_visible(False)
    ax.set_title("ABC-XYZ segmentation (SKU count)")
    _save(fig, "02_abc_xyz_matrix.png")


def inventory_status(policy: pd.DataFrame):
    order = ["Critical - expedite", "Order now", "Healthy", "Overstock"]
    colors = [STATUS["critical"], STATUS["serious"], STATUS["good"], STATUS["warning"]]
    counts = policy["status"].value_counts().reindex(order, fill_value=0)
    fig, ax = plt.subplots(figsize=(8, 3.6))
    ax.barh(order[::-1], counts.values[::-1], color=colors[::-1], height=0.6)
    ax.grid(axis="y", visible=False)
    for i, v in enumerate(counts.values[::-1]):
        ax.text(v + 0.3, i, f"{v}", va="center", color=INK)
    ax.set_xlabel("Number of SKUs")
    ax.set_title("Current inventory status vs reorder policy")
    _save(fig, "03_inventory_status.png")


def supplier_risk(scorecard: pd.DataFrame):
    df = scorecard.sort_values("risk_score")
    tier_color = {"Low": STATUS["good"], "Medium": STATUS["warning"], "High": STATUS["critical"]}
    fig, ax = plt.subplots(figsize=(9, 4.8))
    labels = df["supplier_name"] + " (" + df["country"] + ")"
    ax.barh(labels, df["risk_score"], color=df["risk_tier"].map(tier_color), height=0.6)
    ax.grid(axis="y", visible=False)
    for i, (score, tier) in enumerate(zip(df["risk_score"], df["risk_tier"])):
        ax.text(score + 1, i, f"{score:.0f}  {tier}", va="center", color=INK)
    ax.set(xlim=(0, 100), xlabel="Risk score (0 = low, 100 = high)")
    ax.set_title("Supplier risk score")
    _save(fig, "04_supplier_risk.png")


def otd_vs_variability(scorecard: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.scatter(scorecard["lead_time_cv"] * 100, scorecard["on_time_rate"] * 100,
               s=scorecard["share_of_spend"] * 1500 + 40, color=BLUE, alpha=0.75,
               edgecolor=SURFACE, linewidth=2)
    for r in scorecard.itertuples():
        ax.annotate(r.supplier_name.split()[0], (r.lead_time_cv * 100, r.on_time_rate * 100),
                    xytext=(12, 8), textcoords="offset points", color=INK_2)
    ax.set(xlabel="Lead-time variability (CV %)", ylabel="On-time delivery (%)")
    ax.set_title("Reliability: on-time delivery vs lead-time variability\n(bubble size = share of spend)")
    _save(fig, "05_otd_vs_variability.png")


def forecast_example(demand: pd.DataFrame, forecasts: pd.DataFrame, chosen: pd.DataFrame,
                     products: pd.DataFrame, sku: str):
    hist = demand[demand["sku"] == sku]
    fc = forecasts[forecasts["sku"] == sku]
    info = chosen.set_index("sku").loc[sku]
    name = products.set_index("sku").loc[sku, "product_name"]
    fig, ax = plt.subplots(figsize=(10, 4.6))
    ax.plot(hist["week_start"], hist["units"], color=BLUE, linewidth=2, label="Actual demand")
    ax.plot(pd.to_datetime(fc["week_start"]), fc["forecast_units"], color=ORANGE,
            linewidth=2, linestyle="--", label=f"Forecast ({info['best_method']})")
    ax.set(ylabel="Units per week")
    ax.legend(frameon=False, loc="lower left", bbox_to_anchor=(0, 1.0), ncol=2)
    ax.set_title(f"{sku} {name}: weekly demand and 12-week forecast "
                 f"(test WAPE {info['holdout_wape']:.0%})", pad=28)
    _save(fig, "06_forecast_example.png")


def landed_cost(by_country: pd.DataFrame):
    """Stacked bars: cost to land CAD 100 of product (China-level price) from each country."""
    df = by_country.sort_values("landed_total", ascending=False)
    parts = [("price", "Product price", BLUE), ("freight", "Freight & insurance", ORANGE),
             ("duty", "Import duty", "#1baf7a"), ("brokerage", "Brokerage & handling", "#eda100")]
    fig, ax = plt.subplots(figsize=(9, 5))
    left = np.zeros(len(df))
    labels = df["country"] + "  (" + df["avg_lead_time_wks"].round(1).astype(str) + " wks)"
    for col, name, color in parts:
        ax.barh(labels, df[col], left=left, color=color, height=0.62, label=name,
                edgecolor=SURFACE, linewidth=1.5)
        left += df[col].to_numpy()
    for i, total in enumerate(df["landed_total"]):
        ax.text(total + 1.5, i, f"{total:.0f}", va="center", color=INK)
    ax.grid(axis="y", visible=False)
    ax.set_xlim(0, df["landed_total"].max() * 1.12)
    ax.set_xlabel("CAD to land 100 CAD of product (China price level)")
    ax.legend(frameon=False, ncol=4, loc="lower left", bbox_to_anchor=(0, 1.0), fontsize=9)
    ax.set_title("Landed cost by sourcing country (lead time in brackets)", pad=30)
    _save(fig, "07_landed_cost_by_country.png")
