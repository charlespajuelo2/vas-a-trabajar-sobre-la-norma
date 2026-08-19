from __future__ import annotations

import unittest

from app import ExcelSource, SOURCE_PATH, calculate_bend, calculate_development_and_laps, calculate_hook


class E060ValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = ExcelSource(SOURCE_PATH)
        cls.tolerance = cls.source.sources["validation_tolerance"]["value"]

    def assert_close_to_validation_case(self, result, case):
        self.assertEqual(result["status"], "DERIVADO")
        self.assertAlmostEqual(result["value"], case["expected"], delta=self.tolerance)

    def test_t1_hook_180(self):
        case = self.source.cases[0]
        result = calculate_hook(self.source, 0, case["db"], "longitudinal", case["range"])
        self.assert_close_to_validation_case(result, case)

    def test_t2_hook_90_longitudinal(self):
        case = self.source.cases[1]
        result = calculate_hook(self.source, 1, case["db"], "longitudinal", case["range"])
        self.assert_close_to_validation_case(result, case)

    def test_t3_hook_135_stirrup(self):
        case = self.source.cases[2]
        result = calculate_hook(self.source, 4, case["db"], "estribo-grapa", case["range"])
        self.assert_close_to_validation_case(result, case)

    def test_t4_minimum_bend(self):
        case = self.source.cases[3]
        result = calculate_bend(self.source, case["db"], "longitudinal", case["range"])
        self.assert_close_to_validation_case(result, case)

    def test_t5_compression_development(self):
        case = self.source.cases[4]
        values = {
            "fc": case["fc"],
            "fy": case["fy"],
            "db": case["db"],
            "cover": case["cover"],
            "spacing": case["spacing"],
            "psi_t": case["psi_t"],
            "psi_e": case["psi_e"],
            "psi_s": case["psi_s"],
            "lambda": case["lambda"],
            "ktr": case["ktr"],
            "as_req": case["as_req"],
            "as_prov": case["as_prov"],
            "pct_spliced": case["pct_spliced"],
            "confinement": self.source.sources["one"]["value"],
        }
        result = calculate_development_and_laps(self.source, values)["ld_compression"]
        self.assert_close_to_validation_case(result, case)

    def test_t6_tension_lap(self):
        case = self.source.cases[5]
        values = {
            "fc": case["fc"],
            "fy": case["fy"],
            "db": case["db"],
            "cover": case["cover"],
            "spacing": case["spacing"],
            "psi_t": case["psi_t"],
            "psi_e": case["psi_e"],
            "psi_s": case["psi_s"],
            "lambda": case["lambda"],
            "ktr": case["ktr"],
            "as_req": case["as_req"],
            "as_prov": case["as_prov"],
            "pct_spliced": case["pct_spliced"],
            "confinement": self.source.sources["one"]["value"],
            "ld_base": case["ld_base"],
        }
        result = calculate_development_and_laps(self.source, values)["lap_tension"]
        self.assert_close_to_validation_case(result, case)

    def test_calculation_formulas_reference_inputs(self):
        formulas = [row["formula"] for row in self.source.calc_rows if row["formula"]]
        self.assertTrue(formulas)
        self.assertTrue(all("'ENTRADAS'!" in formula for formula in formulas))


if __name__ == "__main__":
    unittest.main(verbosity=2)
