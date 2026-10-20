"""Invariant tests: python3 -m unittest test_price_feasibility (or pytest)."""
import unittest

import pandas as pd

import price_feasibility as pf

PAR = pd.Series({"e0_mt": 100.0, "e_avg": 0.7, "e_coal": 0.9, "e_hi": 0.9, "p_res": 120.0,
                 "phi": 1.0, "s_max": 0.4, "tau_switch": 60.0, "b40": 0.04,
                 "hi_gen_share": 0.7})


class TestMetrics(unittest.TestCase):
    def test_feebate_has_no_visible_price_or_household_cost(self):
        m = pf.evaluate(PAR, 0, 80)
        self.assertEqual(float(m["F1"]), 0.0)
        self.assertEqual(float(m["F2"]), 0.0)
        self.assertEqual(float(m["A_dem"]), 0.0)
        self.assertGreater(float(m["F3"]), 0.0)

    def test_same_marginal_price_same_substitution(self):
        tax, fb = pf.evaluate(PAR, 50, 0), pf.evaluate(PAR, 0, 50)
        self.assertAlmostEqual(float(tax["A_sub"]), float(fb["A_sub"]))
        self.assertGreater(float(tax["E"]), float(fb["E"]))

    def test_tariff_rise_formula(self):
        # 0.7 t/MWh * $50 / $120 = 29.2%
        self.assertAlmostEqual(float(pf.evaluate(PAR, 50, 0)["F1"]), 100 * 0.7 * 50 / 120)

    def test_zero_pass_through_moves_burden_to_producers(self):
        p = PAR.copy(); p["phi"] = 0.0
        m = pf.evaluate(p, 50, 0)
        self.assertEqual(float(m["F1"]), 0.0)
        self.assertGreater(float(m["F3"]), 0.0)

    def test_rebate_needed_to_protect_bottom_40(self):
        need = PAR.phi * pf.PRIORS["res_share"] * pf.PRIORS["q40"]
        self.assertAlmostEqual(float(pf.evaluate(PAR, 50, 0, rebate=need)["F2"]), 0.0)
        self.assertGreater(float(pf.evaluate(PAR, 50, 0, rebate=need / 2)["F2"]), 0.0)

    def test_hybrid_dominates_and_relaxing_helps(self):
        th = {"f1": 10.0, "f2": 1.0, "f3": 30.0}
        tax, fb, hy = (pf.best(PAR, th, 0.0, m) for m in ("tax", "feebate", "hybrid"))
        self.assertGreaterEqual(hy["E"], max(tax["E"], fb["E"]) - 1e-9)
        self.assertLessEqual(tax["F1"], 10.0)
        self.assertEqual(tax["binding"], "F1")
        loose = pf.best(PAR, dict(th, f1=11.0), 0.0, "tax")
        self.assertGreater(loose["E"], tax["E"])

    def test_real_data_runs(self):
        cal = pf.build_calibration()
        m = pf.country_metrics(cal.loc[["ZAF", "IND", "DEU"]])
        self.assertEqual(len(m), 3)
        self.assertTrue((m.hybrid_E >= m[["tax_E", "feebate_E"]].max(axis=1) - 1e-9).all())


if __name__ == "__main__":
    unittest.main()
