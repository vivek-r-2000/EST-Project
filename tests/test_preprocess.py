"""Unit tests for 02_preprocess.py: filtering, binning, and RPP standardization."""
import unittest
import numpy as np
import pandas as pd
import config as C
import importlib

# In Python, module name starting with a number must be imported via importlib
preprocess_mod = importlib.import_module("02_preprocess")
counts_table = preprocess_mod.counts_table
bin_fossil = preprocess_mod.bin_fossil


class TestCountsTable(unittest.TestCase):
    """Test raw grain aggregation, TRSH-only filtering, RPP weights, and QC filtering."""

    def setUp(self):
        # Create a small synthetic fossil DataFrame
        self.raw_df = pd.DataFrame({
            "dataset_id": [1, 1, 1, 1, 2, 2],
            "lat": [42.0, 42.0, 42.0, 42.0, 20.0, 43.0],  # site 2 row 1 is outside BBOX lat 34-52
            "lon": [-75.0, -75.0, -75.0, -75.0, -75.0, -75.0],
            "depth": [10.0, 10.0, 10.0, 10.0, 5.0, 5.0],
            "age_bp": [1000.0, 1000.0, 1000.0, 1000.0, 500.0, 500.0],
            "taxon": ["Picea glauca", "Quercus alba", "Ambrosia artemisiifolia", "Pinus strobus", "Pinus", "Picea"],
            "group": ["TRSH", "TRSH", "UPHE", "TRSH", "TRSH", "TRSH"],
            "count": [60.0, 40.0, 30.0, 50.0, 100.0, 20.0]
        })

    def test_trsh_only_removes_uphe_weeds(self):
        kept_trsh, audit = counts_table(self.raw_df, trsh_only=True)
        # Ambrosia (UPHE) should have been filtered out before table assembly
        self.assertIn("after_trsh_datasets", audit)
        self.assertIn("after_trsh_rows", audit)
        self.assertEqual(audit["after_trsh_rows"], 5)  # 6 - 1 UPHE

    def test_bbox_filtering_drops_outside_sites(self):
        kept, audit = counts_table(self.raw_df, trsh_only=False)
        # site 2 row 1 has lat=20 (outside BBOX lat 34-52)
        self.assertGreaterEqual(audit["dropped_outside_bbox"], 0)

    def test_min_grains_quality_filter(self):
        # Dataset 1 has 60+40+30+50 = 180 grains (>= 100 MIN_GRAINS -> kept)
        # Dataset 2 inside BBOX has only 20 grains (< 100 MIN_GRAINS -> dropped)
        kept, audit = counts_table(self.raw_df, trsh_only=False)
        self.assertEqual(len(kept), 1)
        self.assertEqual(int(kept.iloc[0]["dataset_id"]), 1)
        self.assertEqual(int(kept.iloc[0]["raw_total"]), 180)

    def test_rpp_division_scales_proportions_without_dropping_sample(self):
        # With RPP, Pinus count (50) is divided by 14.1, Quercus (40) divided by 2.08, Picea (60) divided by 2.80
        kept_raw, _ = counts_table(self.raw_df, use_rpp=False)
        kept_rpp, _ = counts_table(self.raw_df, use_rpp=True)

        row_raw = kept_raw.iloc[0]
        row_rpp = kept_rpp.iloc[0]

        # Raw Pinus is 50.0
        self.assertEqual(row_raw["Pinus"], 50.0)
        # RPP Pinus is 50.0 / 14.10 ~ 3.546
        self.assertAlmostEqual(row_rpp["Pinus"], 50.0 / 14.10, places=3)
        # Raw total remains 180 so sample is not falsely rejected
        self.assertEqual(row_rpp["raw_total"], 180.0)


class TestBinFossil(unittest.TestCase):
    """Test 500-yr binning, proportion re-normalization, and consecutive bin filtering."""

    def test_binning_formula_and_normalization(self):
        # Create a site with 15 continuous bins (meeting MIN_BINS = 12)
        ages = np.arange(500, 8000, 500)  # 15 bins
        rows = []
        for a in ages:
            row = {"dataset_id": 101, "lat": 44.0, "lon": -76.0, "depth": a / 100.0, "age_bp": float(a) + 10.0}
            for g in C.GENERA:
                row[g] = 10.0
            row["Other"] = 10.0
            rows.append(row)
        df_g = pd.DataFrame(rows)

        binned, dropped = bin_fossil(df_g, bin_width=500)
        self.assertEqual(len(dropped), 0)
        self.assertGreaterEqual(len(binned), 12)

        # Proportions across all taxa must sum to 1.0 for every bin
        prop_sums = binned[C.GENERA + ["Other"]].sum(axis=1)
        np.testing.assert_allclose(prop_sums.values, 1.0, atol=1e-12)

    def test_insufficient_bins_dropped(self):
        # Site with only 3 bins (< MIN_BINS 12)
        ages = [500.0, 1000.0, 1500.0]
        rows = []
        for a in ages:
            row = {"dataset_id": 999, "lat": 44.0, "lon": -76.0, "depth": 1.0, "age_bp": a}
            for g in C.GENERA + ["Other"]:
                row[g] = 10.0
            rows.append(row)
        df_g = pd.DataFrame(rows)

        binned, dropped = bin_fossil(df_g, bin_width=500)
        self.assertEqual(len(binned), 0)
        self.assertEqual(len(dropped), 1)
        self.assertIn("min 12", dropped[0])


if __name__ == "__main__":
    unittest.main()
