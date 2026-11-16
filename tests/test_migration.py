"""Unit tests for 04_migration_rates.py: 1D northward, 2D planar arrival, and velocity deficits."""
import unittest
import importlib
import numpy as np
import pandas as pd
from scipy import stats
import config as C

migration_mod = importlib.import_module("04_migration_rates")
arrival_age = migration_mod.arrival_age
fit_migration_1d = migration_mod.fit_migration_1d
fit_migration_2d = migration_mod.fit_migration_2d


class TestArrivalAge(unittest.TestCase):
    """Test empirical arrival age detection and initial presence censoring."""

    def test_two_consecutive_bins_trigger_arrival(self):
        # 14ka to 2ka descending
        ages = np.arange(14000, 1000, -500)
        props = np.zeros(len(ages))
        # Becomes >= 0.05 at 10,000 BP (index 8 and 9)
        idx_10k = np.where(ages == 10000)[0][0]
        props[idx_10k] = 0.06
        props[idx_10k + 1] = 0.08

        s = pd.DataFrame({"bin_age": ages, "Quercus": props})
        arr = arrival_age(s, "Quercus")
        self.assertEqual(arr, 10000.0)

    def test_single_isolated_blip_does_not_trigger_arrival(self):
        ages = np.arange(14000, 1000, -500)
        props = np.zeros(len(ages))
        idx_10k = np.where(ages == 10000)[0][0]
        props[idx_10k] = 0.07  # only 1 bin >= 0.05, next bin is 0.01
        props[idx_10k + 1] = 0.01

        s = pd.DataFrame({"bin_age": ages, "Quercus": props})
        arr = arrival_age(s, "Quercus")
        self.assertTrue(np.isnan(arr))

    def test_already_present_at_start_is_censored(self):
        ages = np.arange(14000, 1000, -500)
        props = np.zeros(len(ages))
        props[0] = 0.10  # present at 14,000 BP (earliest bin)
        props[1] = 0.12

        s = pd.DataFrame({"bin_age": ages, "Picea": props})
        arr = arrival_age(s, "Picea")
        # Must be censored (NaN) because tree was already established before window
        self.assertTrue(np.isnan(arr))


class TestMigration1D(unittest.TestCase):
    """Test 1D latitudinal migration regression and bootstrap confidence intervals."""

    def test_synthetic_northward_wave_recovers_speed(self):
        rng = np.random.default_rng(42)
        # 20 sites spanning latitude 36 to 46 (dist_km = 0 to ~1112 km)
        lats = np.linspace(36.0, 46.0, 20)
        dist_km = (lats - 36.0) * 111.2
        # True speed: 200 m/yr = 0.2 km/yr -> dt/dx = -1000 / 200 = -5.0 yr/km
        # Arrival age = 14000 - 5.0 * dist_km
        true_speed = 200.0
        slope = -1000.0 / true_speed
        arrival_ages = 14000.0 + slope * dist_km + rng.normal(0, 100, 20)

        df = pd.DataFrame({"lat": lats, "arrival_age": arrival_ages})
        res = fit_migration_1d(df, lat0=36.0, rng=rng)

        self.assertLess(res["p"], 0.001)
        self.assertAlmostEqual(res["rate_m_per_yr"], true_speed, delta=30.0)
        self.assertLess(res["ci_low"], res["rate_m_per_yr"])
        self.assertGreater(res["ci_high"], res["rate_m_per_yr"])

    def test_positive_slope_rejected_as_unphysical(self):
        rng = np.random.default_rng(99)
        lats = np.linspace(36.0, 46.0, 15)
        # Arriving later in south, earlier in north (positive slope)
        arrival_ages = 8000.0 + 5.0 * (lats - 36.0) * 111.2
        df = pd.DataFrame({"lat": lats, "arrival_age": arrival_ages})
        res = fit_migration_1d(df, lat0=36.0, rng=rng)
        self.assertTrue(np.isnan(res["rate_m_per_yr"]))


class TestMigration2D(unittest.TestCase):
    """Test 2D planar arrival surface, planar F-test gating, and azimuth vector calculations."""

    def test_directional_2d_wave_recovers_planar_f_test_and_speed(self):
        # Create a grid of 16 points with arrival progressing North-East
        # x (East) and y (North) in km
        xs = np.linspace(-300, 300, 5)
        ys = np.linspace(0, 600, 5)
        grid_x, grid_y = np.meshgrid(xs, ys)
        gx, gy = grid_x.ravel(), grid_y.ravel()

        # Speed = 250 m/yr along heading 45° (NE)
        # dx_km/dt = 0.25 * sin(45°), dy_km/dt = 0.25 * cos(45°)
        # grad_mag = 1000 / 250 = 4.0 yr/km
        bx = -4.0 * np.sin(np.radians(45))
        by = -4.0 * np.cos(np.radians(45))
        ages = 12000.0 + bx * gx + by * gy

        # Convert local km back to approx lat/lon
        lat0, lon0 = 40.0, -80.0
        lats = lat0 + gy / 111.32
        lons = lon0 + gx / (111.32 * np.cos(np.radians(lat0)))

        df = pd.DataFrame({"lat": lats, "lon": lons, "arrival_age": ages})
        rec, grid_df = fit_migration_2d(df, "Carya", grid_n=15)

        self.assertLess(rec["p_2d"], 0.001)
        self.assertGreater(rec["r2_2d"], 0.95)
        self.assertAlmostEqual(rec["rate_2d_median_m_per_yr"], 250.0, delta=15.0)
        # Azimuth should be approx 45°
        self.assertAlmostEqual(rec["mean_azimuth_deg"], 45.0, delta=5.0)
        self.assertIsNotNone(grid_df)

    def test_random_spatial_noise_fails_planar_f_test(self):
        rng = np.random.default_rng(123)
        lats = rng.uniform(36.0, 48.0, 20)
        lons = rng.uniform(-90.0, -70.0, 20)
        # Pure uncorrelated random arrival ages
        ages = rng.uniform(4000.0, 14000.0, 20)

        df = pd.DataFrame({"lat": lats, "lon": lons, "arrival_age": ages})
        rec, grid_df = fit_migration_2d(df, "Acer", grid_n=10)
        # Planar F-test should fail (p > 0.05) and rate should be NaN
        self.assertGreater(rec["p_2d"], 0.05)
        self.assertTrue(np.isnan(rec["rate_2d_median_m_per_yr"]))
        self.assertIsNone(grid_df)


class TestVelocityDeficit(unittest.TestCase):
    """Test climate velocity derivation and modern deficit ratio math."""

    def test_velocity_formula_consistency(self):
        # 1.9 °C/century / (0.01207 °C/km) = 157.41 km/century = 1,574.1 m/yr
        warming_rate = 1.9
        spatial_grad = 0.0120729
        v_clim = (warming_rate / 100.0) / spatial_grad * 1000.0
        self.assertAlmostEqual(v_clim, 1573.8, delta=1.0)

        # Deficit ratio against 169.55 m/yr tree median
        v_tree = 169.55
        ratio = v_clim / v_tree
        self.assertAlmostEqual(ratio, 9.28, delta=0.05)


if __name__ == "__main__":
    unittest.main()
