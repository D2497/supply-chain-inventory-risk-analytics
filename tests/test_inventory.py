"""
Unit tests for the core supply chain formulas.
Run with:  pytest
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import forecasting  # noqa: E402
import inventory    # noqa: E402


def test_z_score_95_percent():
    assert inventory.z_score(0.95) == pytest.approx(1.645, abs=0.001)


def test_eoq_textbook_example():
    # D = 1,000 units/yr, S = $10/order, H = $0.50/unit/yr -> EOQ = 200
    assert inventory.eoq(1000, 10, 0.5) == pytest.approx(200)


def test_eoq_handles_zero_demand():
    assert inventory.eoq(0, 10, 0.5) == 0


def test_safety_stock_demand_variability_only():
    # no lead-time variability -> SS = z * sigma_d * sqrt(LT)
    ss = inventory.safety_stock(z=1.645, demand_mean=100, demand_std=20, lt_mean=4, lt_std=0)
    assert ss == pytest.approx(1.645 * 20 * 2)


def test_safety_stock_grows_with_lead_time_variability():
    low = inventory.safety_stock(1.645, 100, 20, 4, 0.5)
    high = inventory.safety_stock(1.645, 100, 20, 4, 2.0)
    assert high > low


def test_reorder_point():
    assert inventory.reorder_point(demand_mean=50, lt_mean=3, ss=40) == 190


def test_eoq_is_cheapest_order_quantity():
    d, s, h = 5000, 85, 2.2
    q_star = inventory.eoq(d, s, h)
    best = inventory.annual_cycle_cost(d, q_star, s, h)
    for q in (q_star * 0.5, q_star * 0.8, q_star * 1.25, q_star * 2):
        assert inventory.annual_cycle_cost(d, q, s, h) > best


def test_wape():
    actual = np.array([10, 20, 30])
    forecast = np.array([12, 18, 30])
    assert forecasting.wape(actual, forecast) == pytest.approx(4 / 60)


def test_abc_classes_follow_spend():
    weeks = pd.date_range("2025-01-06", periods=52, freq="W-MON")
    demand = pd.concat([
        pd.DataFrame({"week_start": weeks, "sku": sku, "units": units})
        for sku, units in [("BIG", 1000), ("MID", 100), ("SMALL", 10)]
    ])
    products = pd.DataFrame({"sku": ["BIG", "MID", "SMALL"], "unit_cost_cad": [10, 10, 10]})
    result = inventory.abc_xyz(demand, products).set_index("sku")
    assert result.loc["BIG", "abc_class"] == "A"
    assert result.loc["SMALL", "abc_class"] == "C"
    assert (result["xyz_class"] == "X").all()  # constant demand = stable


# ---------------- landed cost ----------------
import landed_cost  # noqa: E402


def test_landed_unit_cost_includes_freight_duty_brokerage():
    opt = landed_cost._option_cost(d=0, sigma_d=0, base_price=100, country="China", lt=10, lt_std=2, z=1.645)
    assert opt["unit"] == pytest.approx(100 * 1.00 * (1 + 0.10 + 0.065 + 0.015))


def test_longer_lead_time_means_more_inventory_cost():
    short = landed_cost._option_cost(50, 10, 100, "Mexico", lt=4, lt_std=1, z=1.645)
    long_ = landed_cost._option_cost(50, 10, 100, "Mexico", lt=11, lt_std=3, z=1.645)
    assert long_["inventory_value"] > short["inventory_value"]
    assert long_["total"] > short["total"]


def test_nearshore_recommendation_rules():
    policy = pd.DataFrame({"sku": ["X1"], "product_name": ["Test"], "abc_class": ["A"], "supplier_id": ["S08"],
                           "unit_cost_cad": [10.0], "forecast_weekly_avg": [100.0], "forecast_rmse_units": [20.0]})
    sc = pd.DataFrame({"supplier_id": ["S08", "S04"], "supplier_name": ["CN", "MX"], "country": ["China", "Mexico"],
                       "lead_time_mean_wks": [11.0, 4.0], "lead_time_std_wks": [3.0, 1.0]})
    out = landed_cost.compare(policy, sc)
    assert len(out) == 1 and out.loc[0, "nearshore_country"] == "Mexico"
    assert out.loc[0, "inventory_released_cad"] > 0
    assert out.loc[0, "recommendation"] in ("Nearshore: lower total cost",
                                            "Nearshore candidate: small premium for resilience", "Keep offshore")
