import unittest
from pathlib import Path

from equest_extractor import (
    convert_value,
    extract_beps_report,
    extract_ls_a_peak_loads,
    extract_lv_b_spaces,
    extract_lv_d_report,
    extract_lv_i_constructions,
    extract_lv_m_conversions,
    extract_es_d_energy_cost_summary,
    extract_ps_h_details,
)


class TestBepsExtractor(unittest.TestCase):
    def test_extracts_all_beps_columns_and_totals_by_fuel(self):
        sim_text = """
        REPORT- BEPS Building Energy Performance
        COMM ELECTRICITY
            MBTU          0.0      0.0     51.6      0.0      0.0      0.0     43.3      0.0      0.0      0.0      0.0     10.9     105.8
        FM1  NATURAL-GAS
            MBTU          0.0      0.0      0.0   2702.0      0.0      0.0      0.0      0.0      0.0      0.0   2389.0      0.0    5091.0
        """
        result = extract_beps_report(sim_text)
        self.assertEqual(result["report"], "BEPS")

    def test_lv_b_space_count_matches_sample_sim(self):
        sim_text = Path("St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM").read_text(errors="ignore")
        result = extract_lv_b_spaces(sim_text)
        self.assertEqual(result["space_count"], 172)

    def test_lv_d_extracts_expected_summary_rows_from_sample_sim(self):
        sim_text = Path("St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM").read_text(errors="ignore")
        result = extract_lv_d_report(sim_text)
        self.assertEqual(result["missing_orientations"], [])

    def test_lv_i_extracts_from_sample_sim(self):
        sim_text = Path("St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM").read_text(errors="ignore")
        result = extract_lv_i_constructions(sim_text)
        self.assertIn("0EWall Construction", result["constructions"])

    def test_ls_a_extracts_peak_loads_and_associates_with_lv_b_spaces(self):
        sim_text = """
        REPORT- LV-B Summary of Spaces
        Spaces on floor: Cellar
        010-Bike Storage                     1.0   INT   89.4    0.80    1.0    0.50   AIR-CHANGE  0.10      1038.1      12457.7
        015-corridor                         1.0   INT    0.0    0.83    0.6    0.20   AIR-CHANGE  0.10       634.8       7617.4
        CONDITIONED FLOOR AREA          =     107479.2  SQFT

        REPORT- LS-A Space Peak Loads Summary
        SPACE NAME                      SPACE  FLOOR      (KBTU/HR)                           (KBTU/HR)
        010-Bike Storage                   1.     1.          0.000                   0.F   0.F         -15.645   DEC 27  5 AM   11.F   9.F
        015-corridor                       1.     1.          2.184   JUN 27  9 PM   85.F  71.F           0.000                   0.F   0.F

        REPORT- LV-B
        """

        lv_b_result = extract_lv_b_spaces(sim_text)
        result = extract_ls_a_peak_loads(sim_text, lv_b_result=lv_b_result)

        self.assertEqual(result["load_units"], "KBTU/HR")
        self.assertEqual(result["space_peak_loads"]["010-Bike Storage"]["cooling_load"], 0.0)
        self.assertEqual(result["space_peak_loads"]["010-Bike Storage"]["heating_load"], -15.645)
        self.assertEqual(
            result["spaces_with_peak_loads"]["015-corridor"]["peak_loads"]["cooling_load"],
            2.184,
        )

    def test_ls_a_parses_integer_load_values(self):
        sim_text = """
        REPORT- LV-B Summary of Spaces
        Spaces on floor: Level 1
        100-Lobby                           1.0   INT    0.0    0.83    0.6    0.20   AIR-CHANGE  0.10       634.8       7617.4

        REPORT- LS-A Space Peak Loads Summary
        SPACE NAME                      SPACE  FLOOR      (KBTU/HR)                           (KBTU/HR)
        100-Lobby                          1.     1.          2       JUN 27  9 PM   85.F  71.F           -3      DEC 27  5 AM   11.F   9.F

        REPORT- LV-B
        """
        lv_b_result = extract_lv_b_spaces(sim_text)
        result = extract_ls_a_peak_loads(sim_text, lv_b_result=lv_b_result)
        self.assertEqual(result["space_peak_loads"]["100-Lobby"]["cooling_load"], 2.0)
        self.assertEqual(result["space_peak_loads"]["100-Lobby"]["heating_load"], -3.0)

    def test_lv_m_conversions_and_convert_value(self):
        sim_text = """
        REPORT- LV-M DOE-2.2 Units Conversion Table
          3        BTU                      0.293000   WH                       3.412969   BTU
          4        BTU/HR                   0.293000   WATT                     3.412969   BTU/HR
        REPORT- LV-B
        """
        result = extract_lv_m_conversions(sim_text)
        self.assertEqual(result["conversions"]["BTU"]["WH"], 0.293)
        self.assertEqual(result["conversions"]["WH"]["BTU"], 3.412969)
        converted = convert_value(10.0, "BTU", "WH", result["conversions"])
        self.assertAlmostEqual(converted, 2.93)

    def test_es_d_extracts_virtual_rate_unit_and_total_charge(self):
        sim_text = """
        REPORT- ES-D Energy Cost Summary
        UTILITY-RATE                       RESOURCE           METERS              UNITS/YR               ($)     ($/UNIT)   ALL YEAR?
        Elec                               ELECTRICITY        EM1   COMM       636613. KWH           108224.       0.1700      YES
        Gas                                NATURAL-GAS        FM1               50910. THERM          59056.       1.1600      YES
        REPORT- ES-E
        """

        result = extract_es_d_energy_cost_summary(sim_text)
        self.assertEqual(result["utility_rates"]["Elec"]["unit"], "KWH")
        self.assertEqual(result["utility_rates"]["Elec"]["total_charge"], 108224.0)
        self.assertEqual(result["utility_rates"]["Elec"]["virtual_rate"], 0.17)
        self.assertEqual(result["utility_rates"]["Gas"]["unit"], "THERM")

    def test_es_d_extracts_from_sample_sim(self):
        sim_text = Path("St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM").read_text(errors="ignore")
        result = extract_es_d_energy_cost_summary(sim_text)
        self.assertEqual(result["utility_rates"]["Elec"]["virtual_rate"], 0.17)
        self.assertEqual(result["utility_rates"]["Gas"]["total_charge"], 59056.0)

    def test_ps_h_extracts_loops_pumps_and_equipment(self):
        sim_text = Path("St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM").read_text(errors="ignore")
        result = extract_ps_h_details(sim_text)

        self.assertEqual(result["loops"]["HW Loop"]["heating_capacity"], -1.84)
        self.assertEqual(result["loops"]["HW Loop"]["units"]["heating_capacity"], "MBTU/HR")
        self.assertEqual(result["pumps"]["HW Loop Pump"]["flow"], 94.9)
        self.assertEqual(result["pumps"]["HW Loop Pump"]["capacity_control"], "ONE-SPEED")
        self.assertEqual(result["equipment"]["HW Boiler 1"]["capacity"], -0.92)
        self.assertEqual(result["equipment"]["HW Boiler 1"]["heat_hir"], 1.3333)


if __name__ == "__main__":
    unittest.main()
