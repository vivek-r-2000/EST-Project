"""Integration tests for output datasets and cross-artifact analytical consistency."""
import unittest
import json
from pathlib import Path
import pandas as pd
import config as C


class TestGroundTruthOutputs(unittest.TestCase):
    """Verify ground truth output artifacts match validated analytical statistics."""

    def test_tests_final_2d_statistics(self):
        p = C.OUT / "tests_final_2d.json"
        self.assertTrue(p.exists(), "tests_final_2d.json missing")
        d = json.loads(p.read_text(encoding="utf-8"))

        # Kruskal-Wallis inter-genus lag test
        self.assertAlmostEqual(d["kruskal_H"], 30.54, places=1)
        self.assertAlmostEqual(d["kruskal_p"], 0.0023, places=3)

        # No-analogue SCD cutoff (5th percentile modern TRSH)
        self.assertAlmostEqual(d["no_analogue_threshold_scd"], 0.1060, places=3)

        # Spearman correlations
        rho_clim, p_clim = d["spearman_noanalogue_vs_abs_climate_rate"]
        self.assertAlmostEqual(rho_clim, 0.607, places=2)
        self.assertLess(p_clim, 0.001)

        rho_veg, p_veg = d["spearman_noanalogue_vs_veg_rate"]
        self.assertAlmostEqual(rho_veg, 0.814, places=2)
        self.assertLess(p_veg, 1e-6)

        # Ebisuzaki surrogate p-value for vegetation turnover
        self.assertAlmostEqual(d["ebisuzaki_p_veg"], 0.033, places=2)
        self.assertEqual(d["n_sites"], 69)

    def test_velocity_comparison_values(self):
        p = C.OUT / "velocity_comparison_final_2d.json"
        self.assertTrue(p.exists(), "velocity_comparison_final_2d.json missing")
        d = json.loads(p.read_text(encoding="utf-8"))

        # Spatial gradient ~0.01207 °C/km (1.21 °C / 100 km)
        self.assertAlmostEqual(d["spatial_gradient_c_per_km"], 0.01207, places=4)

        # Past 95th percentile climate velocity ~332.1 m/yr
        self.assertAlmostEqual(d["past_p95_climate_velocity_m_per_yr"], 332.1, delta=2.0)

        # Modern climate velocity ~1,573.8 m/yr
        self.assertAlmostEqual(d["modern_climate_velocity_m_per_yr"], 1573.8, delta=2.0)

        # 2D median tree migration speed ~169.6 m/yr
        self.assertAlmostEqual(d["median_genus_migration_2d_m_per_yr"], 169.6, delta=2.0)

        # Ratio ~9.28x
        ratio = d["modern_climate_velocity_m_per_yr"] / d["median_genus_migration_2d_m_per_yr"]
        self.assertAlmostEqual(ratio, 9.28, delta=0.1)

    def test_migration_final_2d_table(self):
        p = C.OUT / "migration_final_2d.csv"
        self.assertTrue(p.exists(), "migration_final_2d.csv missing")
        df = pd.read_csv(p)

        # 5 validated genera with p_2d <= 0.05
        sig_2d = df[df["p_2d"] <= 0.05].set_index("genus")
        self.assertEqual(len(sig_2d), 5)
        self.assertSetEqual(set(sig_2d.index), {"Carya", "Alnus", "Fagus", "Larix", "Tsuga"})

        # Speeds and azimuths
        self.assertAlmostEqual(sig_2d.loc["Carya", "rate_2d_median_m_per_yr"], 277.0, delta=1.0)
        self.assertAlmostEqual(sig_2d.loc["Carya", "mean_azimuth_deg"], 65.0, delta=1.0)

        self.assertAlmostEqual(sig_2d.loc["Alnus", "rate_2d_median_m_per_yr"], 264.6, delta=1.0)
        self.assertAlmostEqual(sig_2d.loc["Alnus", "mean_azimuth_deg"], 308.2, delta=1.0)

        self.assertAlmostEqual(sig_2d.loc["Fagus", "rate_2d_median_m_per_yr"], 169.6, delta=1.0)
        self.assertAlmostEqual(sig_2d.loc["Fagus", "mean_azimuth_deg"], 344.4, delta=1.0)

        self.assertAlmostEqual(sig_2d.loc["Larix", "rate_2d_median_m_per_yr"], 124.0, delta=1.0)
        self.assertAlmostEqual(sig_2d.loc["Larix", "mean_azimuth_deg"], 1.2, delta=1.0)

        self.assertAlmostEqual(sig_2d.loc["Tsuga", "rate_2d_median_m_per_yr"], 116.9, delta=1.0)
        self.assertAlmostEqual(sig_2d.loc["Tsuga", "mean_azimuth_deg"], 344.3, delta=1.0)

    def test_no_analogue_trajectory_values(self):
        p = C.OUT / "noanalogue_vs_climate_final_2d.csv"
        self.assertTrue(p.exists(), "noanalogue_vs_climate_final_2d.csv missing")
        df = pd.read_csv(p).sort_values("bin_age")

        # Peak at 14.5 ka BP is ~79.0%
        pk = df.loc[df.frac_no_analogue.idxmax()]
        self.assertEqual(pk["bin_age"], 14500.0)
        self.assertAlmostEqual(pk["frac_no_analogue"] * 100.0, 79.0, delta=0.5)

        # Baseline at 500 BP is ~8.0%
        r500 = df.loc[df.bin_age == 500.0].iloc[0]
        self.assertAlmostEqual(r500["frac_no_analogue"] * 100.0, 8.0, delta=0.5)

    def test_climate_regional_temperature_and_precip_bounds(self):
        p = C.PROC / "climate_regional.csv"
        self.assertTrue(p.exists(), "climate_regional.csv missing")
        df = pd.read_csv(p).sort_values("bin_age")

        # Temperature bounds (-10.27 °C at 14.75 ka to +5.26 °C at 750 BP)
        self.assertAlmostEqual(df.temp_c.min(), -10.27, delta=0.1)
        self.assertAlmostEqual(df.temp_c.max(), 5.26, delta=0.1)
        total_warming = df.temp_c.max() - df.temp_c.min()
        self.assertAlmostEqual(total_warming, 15.53, delta=0.1)

        # Precipitation bounds (2.22 to 2.49 mm/day)
        self.assertAlmostEqual(df.precip_mm_day.min(), 2.22, delta=0.05)
        self.assertAlmostEqual(df.precip_mm_day.max(), 2.49, delta=0.05)


class TestPresentationArtifactIntegrity(unittest.TestCase):
    """Verify that presentation code and PPTX output reflect exact verified figures."""

    def test_presentation_no_deprecated_figures(self):
        gen_script = Path("generate_presentation.py").read_text(encoding="utf-8")
        # Check that none of the outdated figures or threshold numbers appear
        deprecated_tokens = ["0.138", "38.2%", "0.00121", "p = 0.0003", "25°–50°N", "95°–65°W", "9.6°C", "-1.1°C", "+8.5°C"]
        for token in deprecated_tokens:
            self.assertNotIn(token, gen_script, f"Deprecated token '{token}' found in generate_presentation.py")

        # Verify correct study bounding box and verified temperature appear
        self.assertIn("34°–52°N, 100°–60°W", gen_script)
        self.assertIn("-10.3°C", gen_script)
        self.assertIn("+5.3°C", gen_script)
        self.assertIn("+15.5°C", gen_script)

    def test_presentation_file_dimensions_and_slide_count(self):
        import pptx
        pptx_path = Path("TBD2_Final_Project_Presentation_unslopped_new.pptx")
        if not pptx_path.exists():
            pptx_path = Path("TBD2_Final_Project_Presentation_unslopped.pptx")
        self.assertTrue(pptx_path.exists(), "Presentation PPTX missing")
        prs = pptx.Presentation(str(pptx_path))
        self.assertEqual(len(prs.slides), 22, "Must contain exactly 22 slides")
        self.assertAlmostEqual(prs.slide_width.inches, 13.333, places=2)
        self.assertAlmostEqual(prs.slide_height.inches, 7.500, places=2)

        for i, s in enumerate(prs.slides):
            self.assertTrue(s.has_notes_slide, f"Slide {i+1} missing notes slide")
            notes = s.notes_slide.notes_text_frame.text
            self.assertGreater(len(notes.strip()), 50, f"Slide {i+1} notes too short")


if __name__ == "__main__":
    unittest.main()
