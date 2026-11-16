"""Unit tests for lib.py: core mathematical and statistical primitives."""
import unittest
import numpy as np
from scipy import stats
import config as C
import lib


class TestTaxonMapping(unittest.TestCase):
    """Test taxonomic parsing and standardization to C.GENERA."""

    def test_known_genera_extracted(self):
        self.assertEqual(lib.to_genus("Picea glauca"), "Picea")
        self.assertEqual(lib.to_genus("Quercus rubra"), "Quercus")
        self.assertEqual(lib.to_genus("Pinus strobus"), "Pinus")
        self.assertEqual(lib.to_genus("Betula papyrifera"), "Betula")
        self.assertEqual(lib.to_genus("Fagus grandifolia"), "Fagus")

    def test_delimiters_slash_and_hyphen(self):
        self.assertEqual(lib.to_genus("Quercus/Castanea"), "Quercus")
        self.assertEqual(lib.to_genus("Betula-type"), "Betula")
        self.assertEqual(lib.to_genus("Ostrya/Carpinus"), "Other")

    def test_non_focal_and_weeds_map_to_other(self):
        self.assertEqual(lib.to_genus("Ambrosia artemisiifolia"), "Other")
        self.assertEqual(lib.to_genus("Poaceae"), "Other")
        self.assertEqual(lib.to_genus("Asteraceae"), "Other")
        self.assertEqual(lib.to_genus("Cyperaceae"), "Other")

    def test_edge_cases_empty_and_invalid(self):
        self.assertEqual(lib.to_genus(""), "Other")
        self.assertEqual(lib.to_genus("   "), "Other")
        self.assertEqual(lib.to_genus(None), "Other")
        self.assertEqual(lib.to_genus(123), "Other")


class TestDistanceMetrics(unittest.TestCase):
    """Test Hellinger distance and Squared Chord Distance matrix implementations."""

    def test_hellinger_identical_vectors_is_zero(self):
        p = np.array([0.2, 0.3, 0.5])
        self.assertAlmostEqual(lib.hellinger(p, p), 0.0, places=7)

    def test_hellinger_disjoint_vectors_is_one(self):
        p = np.array([1.0, 0.0, 0.0])
        q = np.array([0.0, 1.0, 0.0])
        # sum((sqrt(p) - sqrt(q))^2) = 1 + 1 = 2; sqrt(0.5 * 2) = 1.0
        self.assertAlmostEqual(lib.hellinger(p, q), 1.0, places=7)

    def test_hellinger_symmetry(self):
        p = np.array([0.1, 0.4, 0.5])
        q = np.array([0.3, 0.2, 0.5])
        self.assertAlmostEqual(lib.hellinger(p, q), lib.hellinger(q, p), places=7)

    def test_hellinger_bounded_zero_to_one(self):
        rng = np.random.default_rng(42)
        for _ in range(20):
            p = rng.dirichlet(np.ones(5))
            q = rng.dirichlet(np.ones(5))
            h = lib.hellinger(p, q)
            self.assertGreaterEqual(h, 0.0)
            self.assertLessEqual(h, 1.0)

    def test_scd_matrix_identical_diagonal_zero(self):
        A = np.array([
            [0.1, 0.2, 0.7],
            [0.5, 0.3, 0.2],
            [0.0, 0.0, 1.0]
        ])
        dmat = lib.scd_matrix(A, A)
        self.assertEqual(dmat.shape, (3, 3))
        np.testing.assert_allclose(np.diag(dmat), 0.0, atol=1e-12)

    def test_scd_matrix_symmetry(self):
        A = np.array([[0.2, 0.8], [0.5, 0.5]])
        B = np.array([[0.1, 0.9], [0.4, 0.6], [0.7, 0.3]])
        dmat1 = lib.scd_matrix(A, B)
        dmat2 = lib.scd_matrix(B, A)
        self.assertEqual(dmat1.shape, (2, 3))
        self.assertEqual(dmat2.shape, (3, 2))
        np.testing.assert_allclose(dmat1, dmat2.T, atol=1e-12)

    def test_scd_matrix_disjoint_vectors_max_two(self):
        A = np.array([[1.0, 0.0]])
        B = np.array([[0.0, 1.0]])
        dmat = lib.scd_matrix(A, B)
        # (sqrt(1)-0)^2 + (0-sqrt(1))^2 = 1 + 1 = 2.0
        self.assertAlmostEqual(float(dmat[0, 0]), 2.0, places=7)


class TestSpatialProjection(unittest.TestCase):
    """Test local equal-area tangent plane projection in km."""

    def test_origin_maps_to_zero(self):
        lat0, lon0 = 40.0, -80.0
        x, y, _, _ = lib.latlon_to_xy_km(lat0, lon0, lat0=lat0, lon0=lon0)
        self.assertAlmostEqual(float(x), 0.0, places=6)
        self.assertAlmostEqual(float(y), 0.0, places=6)

    def test_one_degree_latitude_is_111_km(self):
        lat0, lon0 = 40.0, -80.0
        x, y, _, _ = lib.latlon_to_xy_km(lat0 + 1.0, lon0, lat0=lat0, lon0=lon0)
        self.assertAlmostEqual(float(x), 0.0, places=6)
        self.assertAlmostEqual(float(y), 111.32, places=2)

    def test_longitude_scales_with_cos_lat(self):
        lat0, lon0 = 60.0, -80.0
        x, y, _, _ = lib.latlon_to_xy_km(lat0, lon0 + 1.0, lat0=lat0, lon0=lon0)
        # cos(60 deg) = 0.5 -> 111.32 * 0.5 = 55.66 km
        self.assertAlmostEqual(float(x), 55.66, places=2)
        self.assertAlmostEqual(float(y), 0.0, places=6)


class TestEbisuzakiSurrogates(unittest.TestCase):
    """Test FFT phase-randomized surrogate generation and p-value calculation."""

    def test_surrogates_shape_and_preservation_of_mean(self):
        rng = np.random.default_rng(123)
        x = np.sin(np.linspace(0, 10, 30)) + rng.normal(0, 0.2, 30)
        surr = lib.ebisuzaki_surrogates(x, n_sim=50, rng=rng)
        self.assertEqual(surr.shape, (50, 30))
        # Mean of surrogates must equal mean of original series
        np.testing.assert_allclose(surr.mean(axis=1), x.mean(), atol=1e-10)

    def test_surrogates_variance_preservation(self):
        rng = np.random.default_rng(456)
        x = np.linspace(-2, 2, 28) + rng.normal(0, 0.5, 28)
        surr = lib.ebisuzaki_surrogates(x, n_sim=100, rng=rng)
        orig_var = np.var(x, ddof=0)
        surr_vars = np.var(surr, axis=1, ddof=0)
        np.testing.assert_allclose(surr_vars, orig_var, rtol=1e-6)

    def test_ebisuzaki_p_identical_series_is_near_zero(self):
        rng = np.random.default_rng(789)
        x = np.linspace(0, 5, 20) + rng.normal(0, 0.1, 20)
        p_val = lib.ebisuzaki_p(x, x, n_sim=200, rng=rng)
        self.assertLess(p_val, 0.05)

    def test_ebisuzaki_p_uncorrelated_series_not_significant(self):
        rng = np.random.default_rng(999)
        x = rng.normal(0, 1, 40)
        y = rng.normal(0, 1, 40)
        p_val = lib.ebisuzaki_p(x, y, n_sim=500, rng=rng)
        self.assertGreater(p_val, 0.05)

    def test_ebisuzaki_p_edge_cases(self):
        # Sample too short (<4)
        p_short = lib.ebisuzaki_p([1.0, 2.0, 3.0], [1.0, 2.0, 3.0])
        self.assertTrue(np.isnan(p_short))

        # Zero variance series
        p_const = lib.ebisuzaki_p([5.0] * 10, [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0])
        self.assertTrue(np.isnan(p_const))


if __name__ == "__main__":
    unittest.main()
