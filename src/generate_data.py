"""
Step 1 - Generate a realistic (simulated) dataset.

Company: "NorthBridge Medical Supply" - a FICTIONAL Ontario-based distributor
of medical consumables that buys from 8 global suppliers and sells to
clinics and long-term-care homes.

Real company data is confidential, so this script builds a simulated dataset
with realistic patterns:
  * a few SKUs drive most of the spend (Pareto / 80-20 rule)
  * PPE demand spikes in flu season (Nov-Feb)
  * far-away suppliers have longer and less reliable lead times
  * each supplier has its own on-time and defect behaviour

Output: 5 CSV files in /data
"""
import numpy as np
import pandas as pd

import config

rng = np.random.default_rng(config.RANDOM_SEED)

# --------------------------------------------------------------------------
# Suppliers (fictional names)
# columns: id, name, country, promised lead time (weeks), real mean lead time,
#          lead time std, on-time probability, defect rate
# --------------------------------------------------------------------------
SUPPLIERS = [
    ("S01", "Maple Leaf Medical Mfg",     "Canada",   2, 2.1, 0.3, 0.96, 0.004),
    ("S02", "Great Lakes Health Products","USA",      3, 3.2, 0.5, 0.93, 0.006),
    ("S03", "Rhein MedTech GmbH",         "Germany",  6, 6.3, 0.8, 0.91, 0.003),
    ("S04", "Monterrey Surgical SA",      "Mexico",   4, 4.6, 1.0, 0.85, 0.012),
    ("S05", "Penang Glove Industries",    "Malaysia", 8, 9.1, 1.8, 0.78, 0.015),
    ("S06", "Chennai Healthcare Exports", "India",    9, 10.2, 2.2, 0.72, 0.021),
    ("S07", "Saigon Textile Medical",     "Vietnam",  8, 9.0, 2.0, 0.76, 0.018),
    ("S08", "Shenzhen Diagnostics Co",    "China",   10, 11.8, 2.9, 0.68, 0.025),
]

# --------------------------------------------------------------------------
# Products: category -> (items, likely suppliers, price range, seasonal amplitude)
# --------------------------------------------------------------------------
CATALOG = {
    "PPE": (
        ["Nitrile Exam Gloves M", "Nitrile Exam Gloves L", "Vinyl Gloves S",
         "Surgical Mask Level 2", "N95 Respirator", "Isolation Gown",
         "Face Shield", "Shoe Covers", "Bouffant Cap", "Latex Gloves M"],
        ["S05", "S07", "S08", "S02"], (4, 45), 0.45),
    "Wound Care": (
        ["Gauze Pad 4x4", "Foam Dressing 10cm", "Adhesive Bandage Strip",
         "Hydrocolloid Dressing", "Elastic Bandage 3in", "Medical Tape 1in",
         "Alginate Dressing", "Transparent Film Dressing", "Skin Closure Strips"],
        ["S01", "S02", "S06"], (6, 60), 0.05),
    "Diagnostics": (
        ["Rapid Antigen Test Kit", "Glucose Test Strips", "Urine Test Strips",
         "Digital Thermometer", "Pulse Oximeter", "BP Cuff Adult",
         "Lancets 28G", "Specimen Collection Cup"],
        ["S08", "S02", "S03"], (8, 120), 0.30),
    "Surgical": (
        ["Suture Kit 3-0", "Scalpel Blade #10", "Sterile Drape",
         "Surgical Sponge", "Skin Stapler", "Hemostat Forceps",
         "Suction Catheter", "Syringe 10ml", "IV Start Kit"],
        ["S03", "S04", "S06"], (5, 150), 0.05),
    "Patient Care": (
        ["Incontinence Brief L", "Underpad 30x36", "Bed Bath Wipes",
         "Denture Cup", "Emesis Bag", "Patient ID Band", "Mouth Swab",
         "Hand Sanitizer 500ml"],
        ["S01", "S04", "S07"], (3, 40), 0.15),
    "Sterilization": (
        ["Sterilization Pouch", "Indicator Tape", "Biological Indicator",
         "Enzymatic Cleaner 4L", "Instrument Wrap"],
        ["S01", "S03"], (10, 90), 0.0),
}


SOURCE_INFO = {
    "mode": "simulated",
    "company": "NorthBridge Medical Supply",
    "company_desc": "a fictional Ontario-based distributor of medical supplies",
    "note": ("NorthBridge is a fictional company. The data is simulated because real supply chain "
             "data is confidential. The analysis is built to run the same way on a real ERP export."),
}


def build_suppliers() -> pd.DataFrame:
    cols = ["supplier_id", "supplier_name", "country", "promised_lead_time_wks",
            "sim_lt_mean", "sim_lt_std", "sim_on_time_p", "sim_defect_rate"]
    return pd.DataFrame(SUPPLIERS, columns=cols)


def build_products() -> pd.DataFrame:
    rows, n = [], 1
    for category, (items, suppliers, (lo, hi), season_amp) in CATALOG.items():
        for item in items:
            rows.append({
                "sku": f"NB-{n:04d}",
                "product_name": item,
                "category": category,
                "unit_cost_cad": round(float(rng.uniform(lo, hi)), 2),
                "supplier_id": str(rng.choice(suppliers)),
                # hidden simulation parameters (removed before saving)
                "sim_base": float(rng.lognormal(mean=3.4, sigma=1.3)),   # Pareto-like demand
                "sim_cv": float(rng.uniform(0.08, 0.45)),                # noise level
                "sim_season": season_amp * float(rng.uniform(0.6, 1.2)),
                "sim_trend": float(rng.normal(0.0015, 0.002)),           # weekly growth
            })
            n += 1
    return pd.DataFrame(rows)


def build_demand(products: pd.DataFrame) -> pd.DataFrame:
    weeks = pd.date_range(config.START_DATE, periods=config.N_WEEKS, freq="W-MON")
    week_of_year = weeks.isocalendar().week.to_numpy().astype(float)
    # cosine peaks in early January (flu season), lowest in July
    season_curve = np.cos(2 * np.pi * (week_of_year - 2) / 52)

    frames = []
    for p in products.itertuples():
        t = np.arange(len(weeks))
        mean = p.sim_base * (1 + p.sim_trend) ** t * (1 + p.sim_season * season_curve)
        noise = rng.normal(0, p.sim_cv, len(weeks))
        units = np.maximum(0, np.round(mean * (1 + noise))).astype(int)
        frames.append(pd.DataFrame({"week_start": weeks, "sku": p.sku, "units": units}))
    return pd.concat(frames, ignore_index=True)


def build_purchase_orders(products, suppliers, demand) -> pd.DataFrame:
    """Simulate roughly monthly replenishment orders for every SKU."""
    sup = suppliers.set_index("supplier_id")
    weekly_avg = demand.groupby("sku")["units"].mean()
    weeks = sorted(demand["week_start"].unique())

    rows, po_n = [], 1
    for p in products.itertuples():
        s = sup.loc[p.supplier_id]
        for wk in weeks[::4]:  # an order every ~4 weeks
            order_date = pd.Timestamp(wk) + pd.Timedelta(days=int(rng.integers(0, 5)))
            qty = int(max(10, round(weekly_avg[p.sku] * 4 * rng.uniform(0.85, 1.2))))
            promised = order_date + pd.Timedelta(weeks=int(s.promised_lead_time_wks))

            if rng.random() < s.sim_on_time_p:
                # on time: arrives within the promised window (0-3 days early)
                received = promised - pd.Timedelta(days=int(rng.integers(0, 4)))
            else:
                # late: delay grows with how unreliable the route is
                delay_days = int(abs(rng.normal(s.sim_lt_std * 7, s.sim_lt_std * 3))) + 1
                received = promised + pd.Timedelta(days=delay_days)

            short = rng.random() < 0.05                     # 5% of POs ship short
            qty_received = int(qty * rng.uniform(0.85, 0.98)) if short else qty
            defects = int(rng.binomial(qty_received, s.sim_defect_rate))

            rows.append({
                "po_id": f"PO-{po_n:05d}",
                "supplier_id": p.supplier_id,
                "sku": p.sku,
                "order_date": order_date.date(),
                "promised_date": promised.date(),
                "received_date": received.date(),
                "qty_ordered": qty,
                "qty_received": qty_received,
                "qty_defective": defects,
                "po_value_cad": round(qty * p.unit_cost_cad, 2),
            })
            po_n += 1
    return pd.DataFrame(rows)


def build_inventory_snapshot(products, suppliers, demand) -> pd.DataFrame:
    """Current stock position. Most SKUs are stocked sensibly for their
    supplier's lead time, but some are deliberately low or too high so the
    reorder-point logic has something to flag."""
    weekly_avg = demand.groupby("sku")["units"].mean()
    lead_time = suppliers.set_index("supplier_id")["promised_lead_time_wks"]
    rows = []
    for p in products.itertuples():
        # cover expressed as a multiple of the supplier's lead time
        multiple = float(rng.choice([0.4, 1.0, 1.6, 2.0, 2.4, 3.5],
                                    p=[0.08, 0.14, 0.26, 0.26, 0.16, 0.10]))
        weeks_of_cover = multiple * lead_time[p.supplier_id]
        on_hand = int(round(weekly_avg[p.sku] * weeks_of_cover))
        on_order = int(round(weekly_avg[p.sku] * 4)) if rng.random() < 0.35 else 0
        rows.append({"sku": p.sku, "on_hand_units": on_hand, "on_order_units": on_order})
    return pd.DataFrame(rows)


def main():
    config.DATA_DIR.mkdir(exist_ok=True)
    suppliers = build_suppliers()
    products = build_products()
    demand = build_demand(products)
    pos = build_purchase_orders(products, suppliers, demand)
    inventory = build_inventory_snapshot(products, suppliers, demand)

    # drop the hidden "sim_" simulation columns before saving - the analysis must
    # discover supplier behaviour from the data, just like in real life
    public = lambda df: df[[c for c in df.columns if not c.startswith("sim_")]]
    public(suppliers).to_csv(config.DATA_DIR / "suppliers.csv", index=False)
    public(products).to_csv(config.DATA_DIR / "products.csv", index=False)
    demand.to_csv(config.DATA_DIR / "weekly_demand.csv", index=False)
    pos.to_csv(config.DATA_DIR / "purchase_orders.csv", index=False)
    inventory.to_csv(config.DATA_DIR / "inventory_snapshot.csv", index=False)

    print(f"Generated {len(suppliers)} suppliers, {len(products)} SKUs, "
          f"{len(demand):,} demand rows, {len(pos):,} purchase orders.")
    return SOURCE_INFO


if __name__ == "__main__":
    main()
