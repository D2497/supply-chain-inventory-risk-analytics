"""
Step 1 (real-data option) - Load the Kaggle "DataCo Smart Supply Chain" dataset.

Download (free Kaggle account needed):
  https://www.kaggle.com/datasets/shashwatwork/dataco-smart-supply-chain-for-big-data-analysis
Put DataCoSupplyChainDataset.csv in data/raw/ and run:  python src/main.py

What comes from Kaggle (REAL):
  * products, categories, departments and prices
  * ~180,000 order lines -> weekly demand per product (forecasting, ABC-XYZ,
    safety stock, reorder point, EOQ)

What is still simulated (and labelled as such):
  * supplier purchase orders and current stock levels. DataCo is a sales /
    outbound shipping dataset; it has no supplier or inbound PO data, and no
    public dataset I know of does. Each product is assigned to a supplier based
    on its department, and POs are simulated with src/generate_data.py.

Output: the same 5 CSV files the rest of the pipeline already uses.
"""
import numpy as np
import pandas as pd

import config
import generate_data as sim

# Kaggle column names (matched case-insensitively, first match wins)
COLS = {
    "order_date": ["order date (DateOrders)", "order_date"],
    "product_id": ["Product Card Id", "Order Item Cardprod Id"],
    "product_name": ["Product Name"],
    "category": ["Category Name"],
    "department": ["Department Name"],
    "price": ["Order Item Product Price", "Product Price"],
    "profit_ratio": ["Order Item Profit Ratio"],
    "qty": ["Order Item Quantity"],
    "status": ["Order Status"],
}

# Which suppliers plausibly make each DataCo department's products
DEPARTMENT_SUPPLIERS = {
    "Apparel": ["S07", "S06"],
    "Footwear": ["S07", "S08"],
    "Fan Shop": ["S08", "S02", "S04"],
    "Golf": ["S08", "S02"],
    "Outdoors": ["S08", "S04", "S05"],
    "Fitness": ["S08", "S04"],
    "Technology": ["S08", "S05"],
    "Health and Beauty": ["S01", "S02"],
    "Book Shop": ["S01", "S02"],
    "Discs Shop": ["S01", "S02"],
    "Pet Shop": ["S01", "S04"],
}

# Generic supplier names that fit a sporting-goods retailer (same behaviour
# parameters as the simulated suppliers in generate_data.py)
SUPPLIER_NAMES = {
    "S01": "Maple Leaf Manufacturing",
    "S02": "Great Lakes Sporting Goods",
    "S03": "Rhein Precision GmbH",
    "S04": "Monterrey Industrial SA",
    "S05": "Penang Manufacturing Sdn Bhd",
    "S06": "Chennai Textile Exports",
    "S07": "Saigon Apparel & Footwear",
    "S08": "Shenzhen Consumer Goods Co",
}

SOURCE_INFO = {
    "mode": "kaggle",
    "company": "DataCo Global",
    "company_desc": "a global retailer of sporting goods, apparel and electronics",
    "note": ("Product, price and demand data are real: {n:,} order lines from the Kaggle "
             "\"DataCo Smart Supply Chain\" dataset ({start} to {end}, prices converted USD to CAD). "
             "Supplier purchase orders and stock levels are simulated, because no public dataset "
             "includes inbound supplier data."),
}


def _find(df: pd.DataFrame, key: str, required: bool = True):
    lower = {c.lower().strip(): c for c in df.columns}
    for name in COLS[key]:
        if name.lower() in lower:
            return lower[name.lower()]
    if required:
        raise KeyError(f"Column for '{key}' not found. Expected one of {COLS[key]}")
    return None


def read_dataco(path) -> pd.DataFrame:
    """The Kaggle file is Latin-1 encoded; fall back gracefully."""
    for enc in ("utf-8", "latin-1"):
        try:
            return pd.read_csv(path, encoding=enc, low_memory=False)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Could not read {path}")


def build_weekly(raw: pd.DataFrame):
    c = {k: _find(raw, k, required=k not in ("department", "profit_ratio", "status")) for k in COLS}
    df = raw.copy()
    if c["status"]:  # cancelled / fraud orders are not real demand
        df = df[~df[c["status"]].astype(str).str.upper().isin(["CANCELED", "CANCELLED", "SUSPECTED_FRAUD"])]
    df["date"] = pd.to_datetime(df[c["order_date"]], errors="coerce")
    df = df.dropna(subset=["date"])
    df["week_start"] = (df["date"] - pd.to_timedelta(df["date"].dt.weekday, unit="D")).dt.normalize()
    df["sku"] = "DC-" + df[c["product_id"]].astype(str)
    df["qty"] = pd.to_numeric(df[c["qty"]], errors="coerce").fillna(0)
    df["price"] = pd.to_numeric(df[c["price"]], errors="coerce")
    df["profit_ratio"] = pd.to_numeric(df[c["profit_ratio"]], errors="coerce") if c["profit_ratio"] else 0.25
    df["department"] = df[c["department"]].astype(str) if c["department"] else "Other"
    df["category"] = df[c["category"]].astype(str)
    df["product_name"] = df[c["product_name"]].astype(str).str.strip()
    return df


def choose_window(df: pd.DataFrame) -> pd.DatetimeIndex:
    """
    The DataCo file thins out at the very end (partial final weeks and a change
    in product range), so use the last stretch of 'normal' weeks.
    """
    totals = df.groupby("week_start")["qty"].sum()
    totals = totals.reindex(pd.date_range(totals.index.min(), totals.index.max(), freq="W-MON"), fill_value=0)
    normal = totals >= 0.5 * totals.median()
    last_normal = normal[normal].index.max()
    weeks = totals.loc[:last_normal].index[-config.N_WEEKS:]
    if len(weeks) < config.KAGGLE_MIN_WEEKS:
        raise ValueError(f"Only {len(weeks)} usable weeks of history; need {config.KAGGLE_MIN_WEEKS}.")
    return weeks


def main():
    path = config.KAGGLE_DATACO_FILE
    raw = read_dataco(path)
    df = build_weekly(raw)
    weeks = choose_window(df)
    df = df[df["week_start"].isin(weeks)]

    # weekly units per product, missing weeks = 0
    grid = (df.groupby(["sku", "week_start"])["qty"].sum().unstack(fill_value=0)
              .reindex(columns=weeks, fill_value=0))
    active_share = (grid > 0).mean(axis=1)

    # product master: most common name/category/department, median price & margin
    info = df.groupby("sku").agg(product_name=("product_name", lambda s: s.mode().iat[0]),
                                 category=("category", lambda s: s.mode().iat[0]),
                                 department=("department", lambda s: s.mode().iat[0]),
                                 price_usd=("price", "median"),
                                 profit_ratio=("profit_ratio", "median"))
    info["cost_share"] = (1 - info["profit_ratio"].fillna(0.25)).clip(0.30, 0.95)
    info["unit_cost_cad"] = (info["price_usd"] * info["cost_share"] * config.USD_TO_CAD).round(2)
    info["annual_spend"] = grid.iloc[:, -52:].sum(axis=1) * info["unit_cost_cad"]

    # keep products that sell in at least 80% of weeks (forecastable), top N by spend
    keep = info[active_share.reindex(info.index) >= 0.8].nlargest(config.KAGGLE_MAX_SKUS, "annual_spend")
    if len(keep) < 5:
        raise ValueError("Too few products with steady weekly sales in the chosen window.")

    rng = sim.rng
    products = pd.DataFrame({
        "sku": keep.index,
        "product_name": keep["product_name"].values,
        "category": keep["category"].values,
        "unit_cost_cad": keep["unit_cost_cad"].values,
        "supplier_id": [str(rng.choice(DEPARTMENT_SUPPLIERS.get(d, list(SUPPLIER_NAMES))))
                        for d in keep["department"]],
    })
    demand = (grid.loc[products["sku"]].stack().rename("units").reset_index()
                  .rename(columns={"level_1": "week_start"}))
    demand.columns = ["sku", "week_start", "units"]
    demand["units"] = demand["units"].astype(int)
    demand = demand[["week_start", "sku", "units"]].sort_values(["sku", "week_start"])

    suppliers = sim.build_suppliers()
    suppliers["supplier_name"] = suppliers["supplier_id"].map(SUPPLIER_NAMES)
    pos = sim.build_purchase_orders(products, suppliers, demand)
    stock = sim.build_inventory_snapshot(products, suppliers, demand)

    config.DATA_DIR.mkdir(exist_ok=True)
    public = lambda d: d[[c for c in d.columns if not c.startswith("sim_")]]
    public(suppliers).to_csv(config.DATA_DIR / "suppliers.csv", index=False)
    products.to_csv(config.DATA_DIR / "products.csv", index=False)
    demand.to_csv(config.DATA_DIR / "weekly_demand.csv", index=False)
    pos.to_csv(config.DATA_DIR / "purchase_orders.csv", index=False)
    stock.to_csv(config.DATA_DIR / "inventory_snapshot.csv", index=False)
    print(f"Loaded Kaggle DataCo: {len(raw):,} order lines -> {len(products)} SKUs x "
          f"{len(weeks)} weeks ({weeks[0].date()} to {weeks[-1].date()}).")
    info = dict(SOURCE_INFO)
    info["note"] = info["note"].format(n=len(raw), start=weeks[0].strftime("%b %Y"),
                                       end=weeks[-1].strftime("%b %Y"))
    return info


if __name__ == "__main__":
    main()
