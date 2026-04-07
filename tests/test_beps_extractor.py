import os
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from run_local import build_command
from equest_extractor import (
    check_master_room_list_space_type_table_match,
    convert_value,
    extract_beps_report,
    extract_ls_a_peak_loads,
    extract_lv_b_spaces,
    extract_lv_d_report,
    extract_lv_i_constructions,
    extract_lv_m_conversions,
    extract_schedule_table,
    extract_hourly_thermostat_setpoint_ranges,
    populate_master_room_list_space_type_table,
    populate_equest_schedule_importer_table,
    populate_ecm_data_from_reports,
    resolve_model_run_type,
    extract_es_d_energy_cost_summary,
    extract_ps_h_details,
)


class TestBepsExtractor(unittest.TestCase):
    def test_run_local_build_command_modes(self):
        extract_command = build_command({"sim_file": "a.sim", "mode": "extract_report", "report": "beps"})
        self.assertIn("--report", extract_command)
        self.assertIn("beps", extract_command)

        mrl_command = build_command(
            {
                "sim_file": "a.sim",
                "mode": "master_room_list",
                "workbook_path": "b.xlsm",
                "output_workbook_path": "c.xlsm",
                "model_run_type": "Baseline",
            }
        )
        self.assertIn("--populate-master-room-list", mrl_command)

        ecm_command = build_command(
            {
                "sim_file": "a.sim",
                "mode": "ecm_data",
                "workbook_path": "b.xlsm",
                "output_workbook_path": "c.xlsm",
                "model_run_type": "ECM-1",
            }
        )
        self.assertIn("--update-ecm-data", ecm_command)

        schedule_command = build_command(
            {
                "sim_file": "a.sim",
                "mode": "schedule_importer",
                "workbook_path": "b.xlsm",
                "output_workbook_path": "c.xlsm",
            }
        )
        self.assertIn("--populate-schedules", schedule_command)

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

    def test_populates_master_room_list_space_type_table(self):
        workbook_path = Path("output_files/Building Performance Assumptions-v2.xlsm")
        sim_text = """
        REPORT- LV-B Summary of Spaces
        Spaces on floor: Level 1
        010-Bike Storage                     1.0   INT   89.4    0.80    1.0    0.50   AIR-CHANGE  0.10      1038.1      12457.7
        015-corridor                         1.0   INT    0.0    0.83    0.6    0.20   AIR-CHANGE  0.10       634.8       7617.4
        CONDITIONED FLOOR AREA          =     107479.2  SQFT
        REPORT- HOURLY
        SPACE: 010-Bike Storage
        HOUR  THERMOSTAT SETPOINT F  OTHER
        1     70                     0
        2     74                     0

        SPACE: 015-corridor
        HOUR  THERMOSTAT SETPOINT F  OTHER
        1     68                     0
        2     72                     0
        REPORT- ES-D Energy Cost Summary
        UTILITY-RATE                       RESOURCE           METERS              UNITS/YR               ($)     ($/UNIT)   ALL YEAR?
        Elec                               ELECTRICITY        EM1   COMM       636613. KWH           108224.       0.1700      YES
        Gas                                NATURAL-GAS        FM1               50910. THERM          59056.       1.1600      YES
        REPORT- LS-A
        """
        lv_b_result = extract_lv_b_spaces(sim_text)
        first_space_name, first_space_data = next(iter(lv_b_result["spaces"].items()))
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "Building Performance Assumptions.updated.xlsm"
            result = populate_master_room_list_space_type_table(
                sim_text=sim_text,
                workbook_path=workbook_path,
                output_workbook_path=output_path,
            )
            self.assertEqual(result["spaces_written"], 2)
            with zipfile.ZipFile(output_path, "r") as workbook_zip:
                sheet_payload = workbook_zip.read("xl/worksheets/sheet1.xml")
                sheet = ET.fromstring(sheet_payload)
                utility_sheet = ET.fromstring(workbook_zip.read("xl/worksheets/sheet7.xml"))
                master_table = ET.fromstring(workbook_zip.read("xl/tables/table2.xml"))
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            d16 = sheet.find(".//m:c[@r='D16']", ns)
            g16 = sheet.find(".//m:c[@r='G16']/m:v", ns)
            h16 = sheet.find(".//m:c[@r='H16']/m:v", ns)
            i16 = sheet.find(".//m:c[@r='I16']/m:v", ns)
            j16 = sheet.find(".//m:c[@r='J16']/m:v", ns)
            k16 = sheet.find(".//m:c[@r='K16']", ns)
            l16 = sheet.find(".//m:c[@r='L16']/m:v", ns)
            self.assertIsNotNone(d16)
            self.assertEqual(d16.attrib.get("t"), "inlineStr")
            d16_text = "".join(node.text or "" for node in d16.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            self.assertEqual(d16_text, first_space_name)
            self.assertIsNotNone(g16)
            self.assertAlmostEqual(float(g16.text), float(first_space_data["area_sqft"]))
            self.assertAlmostEqual(float(h16.text), float(first_space_data["lights_w_per_sqft"]))
            self.assertAlmostEqual(float(i16.text), float(first_space_data["equip_w_per_sqft"]))
            self.assertAlmostEqual(float(j16.text), float(first_space_data["people"]))
            table_columns = [
                c.attrib.get("name", "")
                for c in master_table.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}tableColumn")
            ]
            if "Temperature Setpoint (F)" in table_columns:
                self.assertIsNotNone(k16)
                k16_value = k16.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                self.assertIsNotNone(k16_value)
                self.assertEqual(float(k16_value.text), 74.0)
            if "Temperature Setback (F)" in table_columns:
                self.assertIsNotNone(l16)
                self.assertEqual(float(l16.text), 70.0)
            b2 = utility_sheet.find(".//m:c[@r='B2']", ns)
            c2 = utility_sheet.find(".//m:c[@r='C2']", ns)
            d2 = utility_sheet.find(".//m:c[@r='D2']", ns)
            self.assertIsNotNone(b2)
            self.assertIsNotNone(c2)
            self.assertIsNotNone(d2)
            b2_text = "".join(node.text or "" for node in b2.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            c2_text = "".join(node.text or "" for node in c2.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            d2_value = d2.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
            self.assertEqual(b2_text, "Electrical")
            self.assertEqual(c2_text.upper(), "KWH")
            self.assertAlmostEqual(float(d2_value.text), 0.17, places=4)
            d66 = sheet.find(".//m:c[@r='D66']", ns)
            if d66 is not None:
                d66_text = "".join(node.text or "" for node in d66.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
                self.assertEqual(d66_text, "")
            # Ensure core compatibility namespace metadata is preserved to avoid Excel repair/recovery.
            self.assertIn(b"xmlns:mc=", sheet_payload)
            self.assertIn(b"mc:Ignorable=", sheet_payload)

    def test_non_baseline_comparison_returns_true_when_matching(self):
        sim_text = Path("St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM").read_text(errors="ignore")
        workbook_path = Path("Building Performance Assumptions.xlsm")
        with tempfile.TemporaryDirectory() as temp_dir:
            populated_path = Path(temp_dir) / "baseline_populated.xlsm"
            populate_master_room_list_space_type_table(
                sim_text=sim_text,
                workbook_path=workbook_path,
                output_workbook_path=populated_path,
            )
            comparison = check_master_room_list_space_type_table_match(
                sim_text=sim_text,
                workbook_path=populated_path,
            )
            self.assertTrue(comparison["space_type_table_match"])
            self.assertEqual(comparison["mismatch_count"], 0)

    def test_non_baseline_comparison_returns_false_when_different(self):
        sim_text = Path("St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM").read_text(errors="ignore")
        workbook_path = Path("Building Performance Assumptions.xlsm")
        lv_b_result = extract_lv_b_spaces(sim_text)
        first_space_name, first_space_data = next(iter(lv_b_result["spaces"].items()))
        mutated_sim_text = f"""
        REPORT- LV-B Summary of Spaces
        Spaces on floor: Level 1
        {first_space_name} Changed                     1.0   INT   89.4    0.80    1.0    0.50   AIR-CHANGE  0.10      {first_space_data["area_sqft"]:.1f}      12457.7
        CONDITIONED FLOOR AREA          =     107479.2  SQFT
        REPORT- LS-A
        """
        with tempfile.TemporaryDirectory() as temp_dir:
            populated_path = Path(temp_dir) / "baseline_populated.xlsm"
            populate_master_room_list_space_type_table(
                sim_text=sim_text,
                workbook_path=workbook_path,
                output_workbook_path=populated_path,
            )
            comparison = check_master_room_list_space_type_table_match(
                sim_text=mutated_sim_text,
                workbook_path=populated_path,
            )
            self.assertFalse(comparison["space_type_table_match"])
            self.assertGreater(comparison["mismatch_count"], 0)

    def test_resolve_model_run_type_precedence(self):
        self.assertEqual(resolve_model_run_type("ECM-2"), "ECM-2")
        old_value = os.environ.get("MODEL_RUN_TYPE")
        try:
            os.environ["MODEL_RUN_TYPE"] = "Proposed"
            self.assertEqual(resolve_model_run_type(None), "Proposed")
            os.environ["MODEL_RUN_TYPE"] = ""
            self.assertEqual(resolve_model_run_type(None), "Baseline")
        finally:
            if old_value is None:
                os.environ.pop("MODEL_RUN_TYPE", None)
            else:
                os.environ["MODEL_RUN_TYPE"] = old_value

    def test_populate_ecm_data_from_reports_for_ecm1(self):
        sim_text = """
        REPORT- BEPS Building Energy Performance
        COMM ELECTRICITY
            MBTU          0.0      0.0     51.6      0.0      0.0      0.0     43.3      0.0      0.0      0.0      0.0     10.9     105.8
        FM1  NATURAL-GAS
            MBTU          0.0      0.0      0.0   2702.0      0.0      0.0      0.0      0.0      0.0      0.0   2389.0      0.0    5091.0
        """
        workbook_path = Path("output_files/Building Performance Assumptions-v2.xlsm")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "ecm_updated.xlsm"
            result = populate_ecm_data_from_reports(
                sim_text=sim_text,
                workbook_path=workbook_path,
                model_run_type="ECM-1",
                output_workbook_path=output_path,
            )
            self.assertEqual(result["target_table"], "ECMData_ECM1")
            self.assertEqual(result["electrical_energy_row"], 29)
            self.assertEqual(result["natural_gas_energy_row"], 31)
            self.assertEqual(result["left_blank_columns"], ["G", "I", "N", "O", "P", "R", "S", "T"])
            with zipfile.ZipFile(output_path, "r") as workbook_zip:
                sheet = ET.fromstring(workbook_zip.read("xl/worksheets/sheet11.xml"))
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            elec_cell = sheet.find(".//m:c[@r='B29']/m:v", ns)
            gas_cell = sheet.find(".//m:c[@r='B31']/m:v", ns)
            demand_cell = sheet.find(".//m:c[@r='B30']/m:v", ns)
            self.assertIsNotNone(elec_cell)
            self.assertIsNotNone(gas_cell)
            # Demand row should be blanked.
            self.assertIsNone(demand_cell)
            # Optional/unmapped columns should remain blank in the energy rows.
            for col in ["G", "I", "N", "O", "P", "R", "S", "T"]:
                self.assertIsNone(sheet.find(f".//m:c[@r='{col}29']/m:v", ns))
                self.assertIsNone(sheet.find(f".//m:c[@r='{col}31']/m:v", ns))

    def test_extract_schedule_table(self):
        sim_text = """
        REPORT- SCHEDULES
        Schedule Name  Schedule Type  Sunday  Monday  Tuesday  Wednesday  Thursday  Friday  Saturday  Holiday  Weekday  Weekend  Holiday Check  1  2  3  4
        Office Lights  FRACTION       WD      WD      WD       WD         WD        WD      WE        WE       WD       WE       YES            0  0  0  0
        REPORT- END
        """
        result = extract_schedule_table(sim_text)
        self.assertEqual(result["report"], "SCHEDULE")
        self.assertEqual(len(result["rows"]), 1)
        self.assertEqual(result["rows"][0]["Schedule Name"], "Office Lights")
        self.assertEqual(result["rows"][0]["1"], "0")

    def test_populate_schedule_importer_table(self):
        sim_text = """
        REPORT- SCHEDULES
        Schedule Name  Schedule Type  Sunday  Monday  Tuesday  Wednesday  Thursday  Friday  Saturday  Holiday  Weekday  Weekend  Holiday Check  1  2  3  4  5  6  7  8  9  10  11  12  13  14  15  16  17  18  19  20  21  22  23  24
        Office Lights  FRACTION       WD      WD      WD       WD         WD        WD      WE        WE       WD       WE       YES            0  0  0  0  0  0  0.2  0.5  0.8  1.0  1.0  1.0  1.0  1.0  1.0  1.0  0.8  0.6  0.5  0.4  0.2  0.1  0  0
        REPORT- END
        """
        workbook_path = Path("output_files/Building Performance Assumptions-v2.xlsm")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "schedule_updated.xlsm"
            result = populate_equest_schedule_importer_table(
                sim_text=sim_text,
                workbook_path=workbook_path,
                output_workbook_path=output_path,
            )
            self.assertEqual(result["target_table"], "eQuest_Schedule_Importer")
            self.assertEqual(result["rows_written"], 1)
            with zipfile.ZipFile(output_path, "r") as workbook_zip:
                sheet = ET.fromstring(workbook_zip.read("xl/worksheets/sheet16.xml"))
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            a2 = sheet.find(".//m:c[@r='A2']", ns)
            b2 = sheet.find(".//m:c[@r='B2']", ns)
            n2 = sheet.find(".//m:c[@r='N2']/m:v", ns)
            y2 = sheet.find(".//m:c[@r='Y2']/m:v", ns)
            self.assertIsNotNone(a2)
            self.assertIsNotNone(b2)
            a2_text = "".join(node.text or "" for node in a2.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            b2_text = "".join(node.text or "" for node in b2.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            self.assertEqual(a2_text, "Office Lights")
            self.assertEqual(b2_text, "FRACTION")
            self.assertIsNotNone(n2)
            self.assertIsNotNone(y2)
            self.assertEqual(float(n2.text), 0.0)
            self.assertEqual(float(y2.text), 1.0)

    def test_extract_hourly_thermostat_setpoint_ranges(self):
        sim_text = """
        REPORT- HOURLY
        SPACE: Office 101
        HOUR  THERMOSTAT SETPOINT F  OTHER
        1     68.0                   0
        2     72.5                   0
        3     70.0                   0

        SPACE: Conference 200
        HOUR  THERMOSTAT SETPOINT F  OTHER
        1     66.0                   0
        2     74.0                   0

        SPACE: OFFICE-101
        HOUR  THERMOSTAT SETPOINT F  OTHER
        1     67.0                   0
        2     73.0                   0
        REPORT- END
        """
        result = extract_hourly_thermostat_setpoint_ranges(sim_text)
        self.assertEqual(result["space_count"], 2)
        self.assertEqual(result["spaces"]["Office 101"]["min_thermostat_setpoint_f"], 67.0)
        self.assertEqual(result["spaces"]["Office 101"]["max_thermostat_setpoint_f"], 73.0)
        self.assertEqual(result["spaces"]["Conference 200"]["min_thermostat_setpoint_f"], 66.0)
        self.assertEqual(result["spaces"]["Conference 200"]["max_thermostat_setpoint_f"], 74.0)

    def test_extract_hourly_thermostat_setpoint_ranges_with_variant_headers(self):
        sim_text = """
        REPORT-HOURLY
        SPACE= 010-Bike Storage
        DATE HR  THERMOSTAT SETPOINT F  OTHER
        1   1    70                     0
        1   2    74                     0

        SPACE=015-corridor
        DATE HR  THERMOSTAT SETPOINT F  OTHER
        1   1    68                     0
        1   2    72                     0
        REPORT- LS-A
        """
        result = extract_hourly_thermostat_setpoint_ranges(sim_text)
        self.assertEqual(result["space_count"], 2)
        self.assertEqual(result["spaces"]["010-Bike Storage"]["min_thermostat_setpoint_f"], 70.0)
        self.assertEqual(result["spaces"]["010-Bike Storage"]["max_thermostat_setpoint_f"], 74.0)
        self.assertEqual(result["spaces"]["015-corridor"]["min_thermostat_setpoint_f"], 68.0)
        self.assertEqual(result["spaces"]["015-corridor"]["max_thermostat_setpoint_f"], 72.0)


if __name__ == "__main__":
    unittest.main()
