"""Unit tests for 03_metrics.py: rate of change, ecological lags, and no-analogue attribution."""
import unittest
import importlib
import numpy as np
import pandas as pd
from scipy import stats
import config as C
import lib

metrics_mod = importlib.import_module("03_metrics")
rate_of_change = metrics_mod.rate_of_change
_best_lag_pair = metrics_mod._best_lag_pair
no_analogue = metrics_mod.no_analogue
lag_summary = metrics_mod.lag_summary


class TestRateOfChange(unittest.TestCase):
    """Test Hellinger vegetation turnover and climate warming rate differencing."""

    def test_adjacent_bins_rate_calculation(self):
        # Two consecutive bins for site 1 (age 1000 and 500, dt = 500)
        p1 = [0.5, 0.5] + [0.0] * (len(lib.COLS) - 2)
        p2 = [0.2, 0.8] + [0.0] * (len(lib.COLS) - 2)
        b = pd.DataFrame([
            {"site_id": 1, "bin_age": 1000.0, **dict(zip(lib.COLS, p1))},
            {"site_id": 1, "bin_age": 500.0, **dict(zip(lib.COLS, p2))},
        ])
        clim_reg = pd.DataFrame({
            "bin_age": [1000.0, 500.0],
            "temp_c": [5.0, 7.0],       # 2°C warming over 500 yr -> 0.4°C / century
            "precip_mm_day": [2.0, 2.5]
        })
        roc, crate = rate_of_change(b, clim_reg, bw=500)
        self.assertEqual(len(roc), 1)
        self.assertEqual(roc.iloc[0]["bin_age"], 750.0)

        # Expected Hellinger distance
        exp_h = lib.hellinger(p1, p2)
        self.assertAlmostEqual(roc.iloc[0]["hellinger"], exp_h, places=6)
        self.assertAlmostEqual(roc.iloc[0]["rate_per_century"], exp_h / 500.0 * 100.0, places=6)

        # Climate rate (shift(-1) is younger bin 500 BP (7°C), - older bin 1000 BP (5°C) -> +2°C / 500yr * 100 = 0.4)
        self.assertEqual(len(crate), 1)
        self.assertAlmostEqual(crate.iloc[0]["climate_rate_c_per_century"], 0.4, places=6)


class TestLagEstimation(unittest.TestCase):
    """Test cross-correlation lag recovery, stationary differencing, and Ebisuzaki testing."""

    def test_planted_lag_recovered_under_diff_mode(self):
        rng = np.random.default_rng(42)
        n = 30
        t = np.linspace(0, 4 * np.pi, n)
        clim = np.sin(t) + rng.normal(0, 0.05, n)

        # Planted lag of k = 2 bins
        planted_k = 2
        veg = np.zeros(n)
        veg[planted_k:] = clim[:-planted_k] + rng.normal(0, 0.05, n - planted_k)

        res = _best_lag_pair(veg, clim, lag_mode="diff", min_bins_lag=10, max_lag_bins=4, ebisuzaki_n=100, rng=rng)
        self.assertIsNotNone(res)
        k_est, r_est, n_pairs, p_eb = res
        self.assertEqual(k_est, planted_k)
        self.assertGreater(r_est, 0.8)
        self.assertLess(p_eb, 0.05)

    def test_negative_lag_correlation_detected_for_cold_retreaters(self):
        rng = np.random.default_rng(84)
        n = 30
        t = np.linspace(0, 4 * np.pi, n)
        clim = np.sin(t)
        # Veg negatively tracks climate with k = 1 bin lag
        veg = np.zeros(n)
        veg[1:] = -clim[:-1]

        res = _best_lag_pair(veg, clim, lag_mode="diff", min_bins_lag=10, max_lag_bins=3, ebisuzaki_n=100, rng=rng)
        self.assertIsNotNone(res)
        k_est, r_est, _, _ = res
        self.assertEqual(k_est, 1)
        self.assertLess(r_est, -0.9)


class TestNoAnalogueSCD(unittest.TestCase):
    """Test Squared Chord Distance novelty detection, thresholding, and additive decomposition."""

    def test_additive_taxonomic_driver_sum_equals_scd(self):
        # 3 fossil samples and 4 modern samples
        rng = np.random.default_rng(101)
        F_props = rng.dirichlet(np.ones(len(lib.COLS)), size=3)
        M_props = rng.dirichlet(np.ones(len(lib.COLS)), size=4)

        b = pd.DataFrame(F_props, columns=lib.COLS)
        b["site_id"] = [1, 2, 3]
        b["bin_age"] = [14000.0, 10000.0, 500.0]

        modern = pd.DataFrame(M_props, columns=lib.COLS)

        out, thr, drv = no_analogue(b, modern)
        self.assertEqual(len(out), 3)
        self.assertGreater(thr, 0.0)

        # Verify that for every row, nearest modern distance matches scd_to_modern
        dmat = lib.scd_matrix(F_props, M_props)
        min_scd = dmat.min(axis=1)
        np.testing.assert_allclose(out["scd_to_modern"].values, min_scd, atol=1e-12)

        # Check binary classification
        expected_no_analogue = min_scd > thr
        np.testing.assert_array_equal(out["no_analogue"].values, expected_no_analogue)


class TestLagSummaryKruskal(unittest.TestCase):
    """Test aggregation of genus lags, median confidence bounds, and Kruskal-Wallis test."""

    def test_kruskal_wallis_detects_inter_genus_clock_differences(self):
        # Genera with distinctly different planted lag distributions
        rows = []
        # Fast pioneer (Larix: lags clustered at 0-1)
        for i in range(15):
            rows.append({"genus": "Larix", "lag_years": 500 if i % 2 == 0 else 1000, "r": 0.5, "sig_05": True})
        # Hardwood (Quercus: lags clustered at 3)
        for i in range(15):
            rows.append({"genus": "Quercus", "lag_years": 1500, "r": 0.6, "sig_05": True})
        # Climax (Fagus: lags clustered at 5)
        for i in range(15):
            rows.append({"genus": "Fagus", "lag_years": 2500, "r": 0.7, "sig_05": True})

        df_lags = pd.DataFrame(rows)
        summary, tests = lag_summary(df_lags)
        self.assertIn("kruskal_H", tests)
        self.assertIn("kruskal_p", tests)
        self.assertLess(tests["kruskal_p"], 0.001)
        self.assertGreater(tests["kruskal_H"], 20.0)


if __name__ == "__main__":
    unittest.main()
