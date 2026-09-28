"""
Step 3 - Demand forecasting with model selection per SKU.

Instead of using one method for every product, each SKU gets a small
"competition" between five forecasting methods:
  * the winner is chosen on a validation window (weeks -24 to -13)
  * its accuracy is then reported on a separate test window (last 12 weeks)
This avoids the common mistake of choosing and grading on the same data.

Methods
  1. Naive            - next week = last week
  2. Moving average   - average of the last 4 weeks
  3. Exponential smoothing (SES) - weighted average, alpha tuned per SKU
  4. Seasonal naive   - same week last year
  5. Seasonal exponential smoothing - removes the yearly pattern
     (seasonal index), smooths the rest, then puts the pattern back

Accuracy metric: WAPE (Weighted Absolute Percentage Error)
  WAPE = sum(|actual - forecast|) / sum(actual)
WAPE is used instead of MAPE because MAPE breaks when a week has 0 sales.
"""
import numpy as np
import pandas as pd

import config

SEASON = 52


def wape(actual: np.ndarray, forecast: np.ndarray) -> float:
    total = np.sum(np.abs(actual))
    return float(np.sum(np.abs(actual - forecast)) / total) if total else np.nan


def ses_level(history: np.ndarray, alpha: float) -> float:
    """Return the final smoothed level after running SES over the history."""
    level = history[0]
    for x in history[1:]:
        level = alpha * x + (1 - alpha) * level
    return level


def best_alpha(train: np.ndarray) -> float:
    """Pick alpha that minimises one-step-ahead error on the training data."""
    best, best_err = 0.3, np.inf
    for alpha in np.arange(0.05, 0.96, 0.05):
        level, err = train[0], 0.0
        for x in train[1:]:
            err += abs(x - level)
            level = alpha * x + (1 - alpha) * level
        if err < best_err:
            best, best_err = alpha, err
    return round(float(best), 2)


def seasonal_indices(history: np.ndarray) -> np.ndarray:
    """
    Multiplicative seasonal index for each of the 52 weeks of the year.
    index > 1 means that week is usually busier than average (e.g. flu season).
    """
    s = pd.Series(history)
    trend = s.rolling(13, center=True, min_periods=7).mean()      # remove noise/trend
    ratio = (s / trend).replace([np.inf, -np.inf], np.nan)
    by_week = ratio.groupby(np.arange(len(s)) % SEASON).mean()     # average same week across years
    by_week = by_week.reindex(range(SEASON)).fillna(1.0)
    # smooth over neighbouring weeks (circular) so one odd week does not dominate
    padded = pd.concat([by_week.iloc[-2:], by_week, by_week.iloc[:2]])
    smooth = padded.rolling(5, center=True).mean().iloc[2:-2].to_numpy()
    return smooth / smooth.mean()


def forecast_methods(history: np.ndarray, horizon: int) -> dict[str, np.ndarray]:
    """Produce a `horizon`-week forecast from each method using only `history`."""
    alpha = best_alpha(history)
    out = {
        "Naive": np.repeat(history[-1], horizon),
        "Moving Average (4w)": np.repeat(history[-4:].mean(), horizon),
        "Exp. Smoothing": np.repeat(ses_level(history, alpha), horizon),
    }
    if len(history) >= SEASON:
        # same weeks one year earlier; wrap if horizon > 52
        idx = [len(history) - SEASON + (h % SEASON) for h in range(horizon)]
        out["Seasonal Naive"] = history[idx].astype(float)

        si = seasonal_indices(history)
        n = len(history)
        deseason = history / si[np.arange(n) % SEASON]
        level = ses_level(deseason, best_alpha(deseason))
        out["Seasonal Exp. Smoothing"] = level * si[(n + np.arange(horizon)) % SEASON]
    return out


def evaluate_and_forecast(demand: pd.DataFrame):
    """
    Returns
      accuracy  : one row per SKU x method with WAPE on the holdout
      chosen    : one row per SKU with the winning method + forecast stats
      forecasts : 12-week future forecast per SKU (long format)
    """
    hold = config.FORECAST_HOLDOUT_WEEKS
    horizon = config.FORECAST_HORIZON_WEEKS
    last_week = demand["week_start"].max()
    future_weeks = pd.date_range(last_week + pd.Timedelta(weeks=1), periods=horizon, freq="W-MON")

    acc_rows, chosen_rows, fc_rows = [], [], []
    for sku, g in demand.sort_values("week_start").groupby("sku"):
        y = g["units"].to_numpy(dtype=float)

        # 1) SELECT: train on data up to 24 weeks ago, score on the next 12
        #    weeks (validation). The winner is picked here.
        val_train, val = y[:-2 * hold], y[-2 * hold:-hold]
        val_scores = {m: wape(val, fc) for m, fc in forecast_methods(val_train, hold).items()}
        winner = min(val_scores, key=val_scores.get)

        # 2) TEST: train on everything except the last 12 weeks and score on
        #    them. The winner never saw these weeks during selection, so the
        #    reported accuracy is honest (no "picking the best on the test set").
        train, test = y[:-hold], y[-hold:]
        test_fc = forecast_methods(train, hold)
        test_scores = {m: wape(test, fc) for m, fc in test_fc.items()}
        for m, s in test_scores.items():
            acc_rows.append({"sku": sku, "method": m, "wape": s})
        # forecast error std-dev -> feeds the safety stock formula
        rmse = float(np.sqrt(np.mean((test - test_fc[winner]) ** 2)))

        # 3) FORECAST: refit the winner on ALL history and look 12 weeks ahead
        future = forecast_methods(y, horizon)[winner]
        fc_rows += [{"sku": sku, "week_start": w.date(), "forecast_units": round(float(v), 1)}
                    for w, v in zip(future_weeks, future)]

        chosen_rows.append({
            "sku": sku,
            "best_method": winner,
            "holdout_wape": round(test_scores[winner], 4),
            "naive_wape": round(test_scores["Naive"], 4),
            "forecast_rmse_units": round(rmse, 2),
            "forecast_weekly_avg": round(float(np.mean(future)), 2),
        })

    accuracy = pd.DataFrame(acc_rows)
    chosen = pd.DataFrame(chosen_rows)
    forecasts = pd.DataFrame(fc_rows)
    return accuracy, chosen, forecasts
