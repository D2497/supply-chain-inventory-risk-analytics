"""
Central settings for the project.

Every business assumption lives here, so you can change one number
(for example the service level) and re-run the whole analysis.
"""
from pathlib import Path

# ---------- Folders ----------
ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
REPORT_DIR = ROOT / "reports"
CHART_DIR = REPORT_DIR / "charts"
DB_PATH = DATA_DIR / "supply_chain.db"

# ---------- Simulation ----------
RANDOM_SEED = 42          # same seed = same data every run (reproducible)
N_WEEKS = 104             # two years of weekly demand history
START_DATE = "2024-01-01"

# ---------- Inventory policy ----------
SERVICE_LEVEL = 0.95      # target: 95% of order cycles with no stockout
ORDER_COST = 85.0         # CAD per purchase order (admin, freight booking, receiving)
HOLDING_COST_RATE = 0.22  # 22% of unit cost per year (capital, storage, insurance, obsolescence)
WEEKS_PER_YEAR = 52

# ---------- ABC / XYZ thresholds ----------
ABC_A_CUTOFF = 0.80       # SKUs making up the first 80% of spend = A
ABC_B_CUTOFF = 0.95       # next 15% = B, last 5% = C
XYZ_X_MAX_CV = 0.25       # coefficient of variation <= 0.25 = X (stable)
XYZ_Y_MAX_CV = 0.50       # <= 0.50 = Y (variable), above = Z (erratic)

# ---------- Forecasting ----------
FORECAST_HOLDOUT_WEEKS = 12   # last 12 weeks are hidden and used to test each model
FORECAST_HORIZON_WEEKS = 12   # how far ahead we forecast

# ---------- Supplier risk weights (must add up to 1.0) ----------
RISK_WEIGHTS = {
    "late_delivery": 0.30,
    "lead_time_variability": 0.20,
    "quality": 0.20,
    "country_risk": 0.15,
    "spend_dependency": 0.15,
}

# Simple 1-5 country/route risk score (distance, port congestion, tariffs,
# geopolitical exposure). These are illustrative assumptions, not official ratings.
COUNTRY_RISK = {
    "Canada": 1,
    "USA": 2,
    "Germany": 2,
    "Mexico": 3,
    "Malaysia": 3,
    "India": 4,
    "Vietnam": 4,
    "China": 5,
}

# ---------- Data source ----------
# "auto"      -> use the Kaggle DataCo file if it is in data/raw/, otherwise simulate
# "simulated" -> always use src/generate_data.py
# "kaggle"    -> always use the Kaggle DataCo file (error if missing)
DATA_SOURCE = "auto"
KAGGLE_DATACO_FILE = DATA_DIR / "raw" / "DataCoSupplyChainDataset.csv"
USD_TO_CAD = 1.35         # DataCo prices are in USD; fixed conversion rate (assumption)
KAGGLE_MAX_SKUS = 60      # keep the top products by spend that sell almost every week
KAGGLE_MIN_WEEKS = 80     # forecasting needs at least ~80 weeks of history

# ---------- Logistics & landed cost (illustrative assumptions) ----------
# Per country: unit price index (China = 1.00 for the same item), freight +
# insurance, import duty and customs brokerage/handling, each as a share of the
# unit price. Duty rates are ILLUSTRATIVE: real rates depend on the HS code and
# trade agreement (CUSMA for USA/Mexico, CETA for Germany, CPTPP for Vietnam and
# Malaysia). Check the Canada Customs Tariff before using real numbers.
LANDED_COST = {
    #            price  freight  duty   brokerage  sourcing region
    "China":    (1.00, 0.100, 0.065, 0.015, "Offshore"),
    "India":    (1.00, 0.110, 0.065, 0.015, "Offshore"),
    "Vietnam":  (1.02, 0.110, 0.000, 0.015, "Offshore"),
    "Malaysia": (0.97, 0.100, 0.000, 0.015, "Offshore"),
    "Germany":  (1.20, 0.070, 0.000, 0.010, "Offshore"),
    "Mexico":   (1.18, 0.045, 0.000, 0.010, "Nearshore"),
    "USA":      (1.38, 0.030, 0.000, 0.010, "Nearshore"),
    "Canada":   (1.45, 0.015, 0.000, 0.000, "Domestic"),
}
NEARSHORE_PREMIUM_OK = 0.03   # accept up to 3% higher landed cost for a much shorter supply chain
