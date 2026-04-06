#!/usr/bin/env python3
"""Extract BEPS, LV-B, LV-D, LV-I, LS-A, LV-M, ES-D, and PS-H report data from an eQuest .SIM file."""
from __future__ import annotations
import argparse
import io
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple
try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover - optional dependency for safer workbook writes
    load_workbook = None
END_USE_COLUMNS = [
    "LIGHTS",
    "TASK LIGHTS",
    "MISC EQUIP",
    "SPACE HEATING",
    "SPACE COOLING",
    "HEAT REJECT",
    "PUMPS & AUX",
    "VENT FANS",
    "REFRIG DISPLAY",
    "HT PUMP SUPPLEM",
    "DOMEST HOT WTR",
    "EXT USAGE",
    "TOTAL",
]
LV_D_COLUMNS = [
    "orientation",
    "avg_u_value_windows",
    "avg_u_value_walls",
    "avg_u_value_walls_plus_windows",
    "window_area",
    "wall_area",
    "window_plus_wall_area",
]
LV_D_UNITS = {
    "avg_u_value_windows": "BTU/HR-SQFT-F",
    "avg_u_value_walls": "BTU/HR-SQFT-F",
    "avg_u_value_walls_plus_windows": "BTU/HR-SQFT-F",
    "window_area": "SQFT",
    "wall_area": "SQFT",
    "window_plus_wall_area": "SQFT",
}
LV_D_TARGET_ORIENTATIONS = {
    "NORTH",
    "NORTH-EAST",
    "EAST",
    "SOUTH-EAST",
    "SOUTH",
    "SOUTH-WEST",
    "WEST",
    "NORTH-WEST",
    "FLOOR",
    "ROOF",
    "ALL WALLS",
    "WALLS+ROOFS",
    "UNDERGRND",
    "BUILDING",
}
NUMBER_PATTERN = re.compile(r"-?[\d,]+(?:\.\d*)?")
CONDITIONED_FLOOR_AREA_PATTERN = re.compile(r"CONDITIONED FLOOR AREA\s*=\s*([\d,]+(?:\.\d+)?)\s+SQFT", re.IGNORECASE)
LV_D_SUMMARY_ROW_PATTERN = re.compile(
    r"^\s*([A-Z][A-Z\-\+\s]+?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+"
    r"(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s*$"
)
LV_I_ROW_PATTERN = re.compile(
    r"^\s*(.+?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(\d+)\s+(DELAYED|QUICK)\s+(\d+)\s*$"
)
LV_I_UVALUE_UNIT_PATTERN = re.compile(r"U-VALUE\s*\(([^\)]+)\)", re.IGNORECASE)
LS_A_LOAD_UNIT_PATTERN = re.compile(r"COOLING LOAD\s*\(([^\)]+)\)", re.IGNORECASE)
REPORT_HEADER_PATTERN = re.compile(r"REPORT-\s*([A-Z0-9\-]+)", re.IGNORECASE)
SCHEDULE_HOURLY_COLUMN_PATTERN = re.compile(r"^(\d{1,2})(?:\s*(AM|PM))?$", re.IGNORECASE)
THERMOSTAT_HEADER_PATTERN = re.compile(r"THERMOSTAT\s+SETPOINT\s+F", re.IGNORECASE)
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"m": MAIN_NS}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_MAP_NS = {"m": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS}
MASTER_ROOM_LIST_SHEET_XML_PATH = "xl/worksheets/sheet1.xml"
MASTER_ROOM_LIST_SPACE_START_ROW = 16
MASTER_ROOM_LIST_SPACE_MAX_ROWS = 50
ECM_DATA_SHEET_XML_PATH = "xl/worksheets/sheet10.xml"
ECM_DATA_MODEL_START_ROWS = {
    "BASELINE": 4,
    "PROPOSED": 34,
    "ECM-1": 44,
    "ECM-2": 54,
    "ECM-3": 64,
    "ECM-4": 74,
    "ECM-5": 84,
    "ECM-6": 94,
    "ECM-7": 104,
}
BEPS_TO_ECM_END_USE_COLUMNS = {
    "LIGHTS": "B",  # Internal Lighting
    "EXT USAGE": "C",  # External Lighting
    "SPACE HEATING": "D",
    "SPACE COOLING": "E",
    "PUMPS & AUX": "F",
    "VENT FANS": "H",  # Fans Interior
    "DOMEST HOT WTR": "J",
    "MISC EQUIP": "K",  # Receptacle Equipment
    "TASK LIGHTS": "L",  # Interior Lighting Process
    "REFRIG DISPLAY": "M",
    "HEAT REJECT": "Q",
}
ECM_OPTIONAL_BLANK_COLUMNS = ["G", "I", "N", "O", "P", "R", "S", "T"]
KBTU_PER_UNIT = {
    "KWH": 3.412,
    "THERM": 100.0,
    "KBTU": 1.0,
    "MBTU": 1000.0,
    "MMBTU": 1000.0,
    "BTU": 0.001,
}

MODEL_RUN_TYPE_TO_ECM_TABLE = {
    "BASELINE": "ECMData_Baseline",
    "PROPOSED": "ECMData_Proposed",
    "ECM-1": "ECMData_ECM1",
    "ECM-2": "ECMData_ECM2",
    "ECM-3": "ECMData_ECM3",
    "ECM-4": "ECMData_ECM4",
    "ECM-5": "ECMData_ECM5",
    "ECM-6": "ECMData_ECM6",
    "ECM-7": "ECMData_ECM7",
}
def _clean_number(value: str) -> float:
    return float(value.replace(",", ""))
def _parse_values_line(line: str) -> tuple[str, List[float]]:
    stripped = line.strip()
    if not stripped:
        raise ValueError("Expected values line but got blank line.")
    unit = stripped.split()[0]
    numbers = [_clean_number(token) for token in NUMBER_PATTERN.findall(line)]
    if len(numbers) != len(END_USE_COLUMNS):
        raise ValueError(
            f"Expected {len(END_USE_COLUMNS)} numeric BEPS columns but found {len(numbers)} in line: {line!r}"
        )
    return unit, numbers
            "usage": "Use convert_value(value, from_unit, to_unit, conversions) for future transformations.",
    }
def convert_value(value: float, from_unit: str, to_unit: str, conversions: Dict[str, Dict[str, float]]) -> float:
    """Convert a value between units using LV-M conversion factors (supports chained conversions)."""
    if from_unit == to_unit:
        return value
    visited = set()
    queue: List[tuple[str, float]] = [(from_unit, value)]
    while queue:
        current_unit, current_value = queue.pop(0)
        if current_unit in visited:
            continue
        visited.add(current_unit)
        for next_unit, factor in conversions.get(current_unit, {}).items():
            next_value = current_value * factor
            if next_unit == to_unit:
                return next_value
            if next_unit not in visited:
                queue.append((next_unit, next_value))
    raise ValueError(f"No conversion path found from '{from_unit}' to '{to_unit}'.")


def _ensure_row(sheet_data: ET.Element, row_number: int) -> ET.Element:
    row = sheet_data.find(f"m:row[@r='{row_number}']", NS)
    if row is not None:
        return row
    row = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_number)})
    inserted = False
    for existing in sheet_data.findall("m:row", NS):
        if int(existing.attrib["r"]) > row_number:
            sheet_data.insert(list(sheet_data).index(existing), row)
            inserted = True
            break
    if not inserted:
        sheet_data.append(row)
    return row


def _set_inline_string_cell(row: ET.Element, cell_ref: str, value: str, style: str | None = None) -> None:
    cell = row.find(f"m:c[@r='{cell_ref}']", NS)
    if cell is None:
        attrs = {"r": cell_ref}
        if style is not None:
            attrs["s"] = style
        cell = ET.SubElement(row, f"{{{MAIN_NS}}}c", attrs)
    elif style is not None and "s" not in cell.attrib:
        cell.attrib["s"] = style
    for child in list(cell):
        cell.remove(child)
    cell.attrib["t"] = "inlineStr"
    is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
    text_node = ET.SubElement(is_node, f"{{{MAIN_NS}}}t")
    text_node.text = value


def _set_numeric_cell(row: ET.Element, cell_ref: str, value: float | None, style: str | None = None) -> None:
    cell = row.find(f"m:c[@r='{cell_ref}']", NS)
    if cell is None:
        attrs = {"r": cell_ref}
        if style is not None:
            attrs["s"] = style
        cell = ET.SubElement(row, f"{{{MAIN_NS}}}c", attrs)
    elif style is not None and "s" not in cell.attrib:
        cell.attrib["s"] = style
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if value is not None:
        v_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
        v_node.text = f"{value:.6f}".rstrip("0").rstrip(".")


def _load_zip_file_map(workbook_path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(workbook_path, "r") as workbook_zip:
        return {name: workbook_zip.read(name) for name in workbook_zip.namelist()}


def _save_zip_file_map(file_map: Dict[str, bytes], output_workbook_path: Path) -> None:
    with zipfile.ZipFile(output_workbook_path, "w", compression=zipfile.ZIP_DEFLATED) as dst_zip:
        for name, payload in file_map.items():
            dst_zip.writestr(name, payload)


def _parse_xml_with_registered_namespaces(xml_payload: bytes) -> ET.Element:
    for _, ns in ET.iterparse(io.BytesIO(xml_payload), events=("start-ns",)):
        prefix, uri = ns
        ET.register_namespace(prefix or "", uri)
    return ET.fromstring(xml_payload)


def _to_kbtu(value: float, from_unit: str) -> float:
    normalized_unit = from_unit.upper().strip()
    if normalized_unit not in KBTU_PER_UNIT:
        raise ValueError(f"Unsupported unit conversion to kBtu from '{from_unit}'.")
    return value * KBTU_PER_UNIT[normalized_unit]


def _column_to_index(column: str) -> int:
    index = 0
    for char in column:
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index


def _index_to_column(index: int) -> str:
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _split_cell_ref(cell_ref: str) -> Tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref.upper())
    if not match:
        raise ValueError(f"Unsupported cell reference format: {cell_ref}")
    return match.group(1), int(match.group(2))


def _parse_range_ref(range_ref: str) -> Tuple[str, int, str, int]:
    start, end = range_ref.split(":")
    start_col, start_row = _split_cell_ref(start)
    end_col, end_row = _split_cell_ref(end)
    return start_col, start_row, end_col, end_row


def _build_table_index(file_map: Dict[str, bytes]) -> Dict[str, Dict[str, object]]:
    workbook_root = _parse_xml_with_registered_namespaces(file_map["xl/workbook.xml"])
    workbook_rels_root = _parse_xml_with_registered_namespaces(file_map["xl/_rels/workbook.xml.rels"])
    workbook_rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in workbook_rels_root.findall("pr:Relationship", REL_MAP_NS)
    }
    table_index: Dict[str, Dict[str, object]] = {}
    for sheet_node in workbook_root.findall("m:sheets/m:sheet", REL_MAP_NS):
        sheet_name = sheet_node.attrib.get("name", "")
        rel_id = sheet_node.attrib.get(f"{{{REL_NS}}}id")
        if not rel_id or rel_id not in workbook_rel_targets:
            continue
        sheet_target = workbook_rel_targets[rel_id]
        sheet_xml_path = f"xl/{sheet_target.lstrip('/')}"
        if sheet_xml_path not in file_map:
            continue
        sheet_rels_path = sheet_xml_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if sheet_rels_path not in file_map:
            continue
        sheet_root = _parse_xml_with_registered_namespaces(file_map[sheet_xml_path])
        sheet_rels_root = _parse_xml_with_registered_namespaces(file_map[sheet_rels_path])
        sheet_rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in sheet_rels_root.findall("pr:Relationship", REL_MAP_NS)
        }
        for table_part in sheet_root.findall("m:tableParts/m:tablePart", REL_MAP_NS):
            table_rel_id = table_part.attrib.get(f"{{{REL_NS}}}id")
            if not table_rel_id or table_rel_id not in sheet_rel_targets:
                continue
            table_target = sheet_rel_targets[table_rel_id]
            table_xml_path = table_target.replace("../", "xl/")
            if table_xml_path not in file_map:
                continue
            table_root = _parse_xml_with_registered_namespaces(file_map[table_xml_path])
            table_name = table_root.attrib.get("name")
            table_ref = table_root.attrib.get("ref")
            if not table_name or not table_ref:
                continue
            table_columns = [
                column.attrib.get("name", "")
                for column in table_root.findall("m:tableColumns/m:tableColumn", REL_MAP_NS)
            ]
            table_index[table_name] = {
                "sheet_name": sheet_name,
                "sheet_xml_path": sheet_xml_path,
                "table_xml_path": table_xml_path,
                "ref": table_ref,
                "columns": table_columns,
            }
    return table_index


def _normalize_model_run_type(model_run_type: str) -> str:
    normalized = model_run_type.strip().upper().replace(" ", "")
    if normalized.startswith("ECM") and "-" not in normalized and len(normalized) == 4:
        normalized = f"ECM-{normalized[-1]}"
    return normalized


def _table_column_letter(table_ref: str, table_columns: List[str], target_column_name: str) -> str:
    if target_column_name not in table_columns:
        raise ValueError(f"Column '{target_column_name}' not found in target table.")
    start_col, _, _, _ = _parse_range_ref(table_ref)
    start_col_index = _column_to_index(start_col)
    target_index = table_columns.index(target_column_name)
    return _index_to_column(start_col_index + target_index)


def _load_master_room_list_sheet(workbook_path: Path) -> ET.Element:
    with zipfile.ZipFile(workbook_path, "r") as workbook_zip:
        try:
            sheet_xml = workbook_zip.read(MASTER_ROOM_LIST_SHEET_XML_PATH)
        except KeyError as exc:
            raise ValueError("Could not find Master Room List worksheet XML in the workbook.") from exc
    return _parse_xml_with_registered_namespaces(sheet_xml)


def _read_cell_text(row: ET.Element, cell_ref: str) -> str:
    cell = row.find(f"m:c[@r='{cell_ref}']", NS)
    if cell is None:
        return ""
    if cell.attrib.get("t") == "inlineStr":
        text_parts = [node.text or "" for node in cell.findall(".//m:t", NS)]
        return "".join(text_parts)
    value_node = cell.find("m:v", NS)
    return value_node.text.strip() if value_node is not None and value_node.text else ""


def _read_cell_float(row: ET.Element, cell_ref: str) -> float | None:
    cell = row.find(f"m:c[@r='{cell_ref}']", NS)
    if cell is None:
        return None
    value_node = cell.find("m:v", NS)
    if value_node is None or value_node.text is None or not value_node.text.strip():
        return None
    try:
        return float(value_node.text.strip())
    except ValueError:
        return None


def populate_master_room_list_space_type_table(
    sim_text: str,
    workbook_path: Path,
    output_workbook_path: Path,
) -> Dict[str, object]:
    """Populate Master Room List and Utilities table data using LV-B and ES-D content."""
    lv_b_result = extract_lv_b_spaces(sim_text)
    es_d_result = extract_es_d_energy_cost_summary(sim_text)
    try:
        thermostat_result = extract_hourly_thermostat_setpoint_ranges(sim_text)
    except ValueError:
        thermostat_result = {"spaces": {}}
    spaces = list(lv_b_result["spaces"].items())
    if not spaces:
        raise ValueError("No LV-B spaces found to populate the Master Room List.")
    ET.register_namespace("", MAIN_NS)
    file_map = _load_zip_file_map(workbook_path)
    table_index = _build_table_index(file_map)
    master_table = table_index.get("MasterRoomList_SpaceTypeTable")
    utility_table = table_index.get("Utility_Rates")
    if master_table is None:
        raise ValueError("Could not locate MasterRoomList_SpaceTypeTable table in workbook.")
    if utility_table is None:
        raise ValueError("Could not locate Utility_Rates table in workbook.")

    master_sheet_path = str(master_table["sheet_xml_path"])
    sheet_root = _parse_xml_with_registered_namespaces(file_map[master_sheet_path])
    sheet_data = sheet_root.find("m:sheetData", NS)
    if sheet_data is None:
        raise ValueError("Workbook sheet is missing sheetData.")

    master_table_ref = str(master_table["ref"])
    master_table_columns = list(master_table["columns"])
    start_col, start_row, _, end_row = _parse_range_ref(master_table_ref)
    max_rows = end_row - start_row

    column_name_ref = _table_column_letter(master_table_ref, master_table_columns, "Space Name")
    column_area_ref = _table_column_letter(master_table_ref, master_table_columns, "Area (sf)")
    column_lighting_ref = _table_column_letter(master_table_ref, master_table_columns, "Lighting Power Density (W/sqft)")
    column_equip_ref = _table_column_letter(master_table_ref, master_table_columns, "Equipment Power Density (W/sqft)")
    column_people_ref = _table_column_letter(master_table_ref, master_table_columns, "Occupants (# of People)")
    setpoint_column_name = "Temperature Setpoint (F)"
    setback_column_name = "Temperature Setback (F)"
    column_setpoint_ref = (
        _table_column_letter(master_table_ref, master_table_columns, setpoint_column_name)
        if setpoint_column_name in master_table_columns
        else None
    )
    column_setback_ref = (
        _table_column_letter(master_table_ref, master_table_columns, setback_column_name)
        if setback_column_name in master_table_columns
        else None
    )
    thermostat_by_space_key = {
        _canonical_space_key(space_name): setpoint_data
        for space_name, setpoint_data in thermostat_result.get("spaces", {}).items()
    }

    for index in range(max_rows):
        row_number = (start_row + 1) + index
        row = _ensure_row(sheet_data, row_number)
        if index < len(spaces):
            space_name, space_data = spaces[index]
            _set_inline_string_cell(row, f"{column_name_ref}{row_number}", space_name)
            _set_numeric_cell(row, f"{column_area_ref}{row_number}", float(space_data["area_sqft"]))
            _set_numeric_cell(row, f"{column_lighting_ref}{row_number}", float(space_data["lights_w_per_sqft"]))
            _set_numeric_cell(row, f"{column_equip_ref}{row_number}", float(space_data["equip_w_per_sqft"]))
            _set_numeric_cell(row, f"{column_people_ref}{row_number}", float(space_data["people"]))
            if column_setpoint_ref is not None or column_setback_ref is not None:
                setpoint_data = thermostat_by_space_key.get(_canonical_space_key(space_name))
                if setpoint_data is not None:
                    min_temp = float(setpoint_data["min_thermostat_setpoint_f"])
                    max_temp = float(setpoint_data["max_thermostat_setpoint_f"])
                else:
                    min_temp = None
                    max_temp = None
                if column_setpoint_ref is not None:
                    _set_numeric_cell(row, f"{column_setpoint_ref}{row_number}", max_temp)
                if column_setback_ref is not None:
                    _set_numeric_cell(row, f"{column_setback_ref}{row_number}", min_temp)
        else:
            _set_inline_string_cell(row, f"{column_name_ref}{row_number}", "")
            _set_numeric_cell(row, f"{column_area_ref}{row_number}", None)
            _set_numeric_cell(row, f"{column_lighting_ref}{row_number}", None)
            _set_numeric_cell(row, f"{column_equip_ref}{row_number}", None)
            _set_numeric_cell(row, f"{column_people_ref}{row_number}", None)
            if column_setpoint_ref is not None:
                _set_numeric_cell(row, f"{column_setpoint_ref}{row_number}", None)
            if column_setback_ref is not None:
                _set_numeric_cell(row, f"{column_setback_ref}{row_number}", None)
    file_map[master_sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    utility_sheet_path = str(utility_table["sheet_xml_path"])
    utility_sheet_root = _parse_xml_with_registered_namespaces(file_map[utility_sheet_path])
    utility_sheet_data = utility_sheet_root.find("m:sheetData", NS)
    if utility_sheet_data is None:
        raise ValueError("Utilities sheet is missing sheetData.")
    utility_ref = str(utility_table["ref"])
    utility_columns = list(utility_table["columns"])
    _, utility_start_row, _, utility_end_row = _parse_range_ref(utility_ref)
    utility_rows = utility_end_row - utility_start_row
    utility_mapping = {
        "Elec": "Electrical",
        "Gas": "Natural Gas",
    }
    provider_col = _table_column_letter(utility_ref, utility_columns, "Utility Provider")
    power_type_col = _table_column_letter(utility_ref, utility_columns, "Power Type")
    unit_col = _table_column_letter(utility_ref, utility_columns, "Unit of Measurement")
    cost_col = _table_column_letter(utility_ref, utility_columns, "Cost per Measurement")
    available_rates = es_d_result["utility_rates"]
    utility_rows_written = 0
    for index in range(utility_rows):
        row_number = (utility_start_row + 1) + index
        row = _ensure_row(utility_sheet_data, row_number)
        if index < len(utility_mapping):
            provider = list(utility_mapping.keys())[index]
            power_type = utility_mapping[provider]
            rate_data = available_rates.get(provider, {})
            _set_inline_string_cell(row, f"{provider_col}{row_number}", provider)
            _set_inline_string_cell(row, f"{power_type_col}{row_number}", power_type)
            _set_inline_string_cell(row, f"{unit_col}{row_number}", str(rate_data.get("unit", "")))
            _set_numeric_cell(row, f"{cost_col}{row_number}", float(rate_data["virtual_rate"]) if "virtual_rate" in rate_data else None)
            if rate_data:
                utility_rows_written += 1
        else:
            _set_inline_string_cell(row, f"{provider_col}{row_number}", "")
            _set_inline_string_cell(row, f"{power_type_col}{row_number}", "")
            _set_inline_string_cell(row, f"{unit_col}{row_number}", "")
            _set_numeric_cell(row, f"{cost_col}{row_number}", None)
    file_map[utility_sheet_path] = ET.tostring(utility_sheet_root, encoding="utf-8", xml_declaration=True)
    _save_zip_file_map(file_map, output_workbook_path)
    return {
        "target_sheet": "Master Room List",
        "target_table": "Space Type Table",
        "utility_sheet": "Utilities",
        "utility_table": "Utility_Rates",
        "rows_available": max_rows,
        "spaces_found": len(spaces),
        "spaces_written": min(len(spaces), max_rows),
        "spaces_truncated": max(len(spaces) - max_rows, 0),
        "utility_rows_written": utility_rows_written,
        "output_workbook": str(output_workbook_path),
    }


def populate_ecm_data_from_reports(
    sim_text: str,
    workbook_path: Path,
    model_run_type: str,
    output_workbook_path: Path,
) -> Dict[str, object]:
    """Populate ECM Data table rows for electrical and natural gas end uses."""
    normalized_model_run_type = _normalize_model_run_type(model_run_type)
    if normalized_model_run_type not in MODEL_RUN_TYPE_TO_ECM_TABLE:
        raise ValueError(
            "Unsupported model run type for ECM Data. Supported: Baseline, Proposed, ECM-1..ECM-7 "
            "(Baseline-2 and Baseline-3 are intentionally ignored)."
        )
    beps_result = extract_beps_report(sim_text)
    elec_totals = beps_result["totals_by_fuel"]["electricity"]
    gas_totals = beps_result["totals_by_fuel"]["natural_gas"]
    elec_unit = elec_totals["unit"]
    gas_unit = gas_totals["unit"]
    elec_end_use_values_kbtu: Dict[str, float] = {}
    gas_end_use_values_kbtu: Dict[str, float] = {}
    for beps_column, target_column in BEPS_TO_ECM_END_USE_COLUMNS.items():
        elec_value = float(elec_totals["by_end_use"][beps_column])
        gas_value = float(gas_totals["by_end_use"][beps_column])

        elec_end_use_values_kbtu[target_column] = _to_kbtu(elec_value, elec_unit) if abs(elec_value) > 1e-9 else 0.0
        gas_end_use_values_kbtu[target_column] = _to_kbtu(gas_value, gas_unit) if abs(gas_value) > 1e-9 else 0.0

    target_table_name = MODEL_RUN_TYPE_TO_ECM_TABLE[normalized_model_run_type]
    ET.register_namespace("", MAIN_NS)
    file_map = _load_zip_file_map(workbook_path)
    table_index = _build_table_index(file_map)
    target_table = table_index.get(target_table_name)
    if target_table is None:
        raise ValueError(f"Could not locate ECM Data table '{target_table_name}' in workbook.")
    sheet_path = str(target_table["sheet_xml_path"])
    sheet_root = _parse_xml_with_registered_namespaces(file_map[sheet_path])
    sheet_data = sheet_root.find("m:sheetData", NS)
    if sheet_data is None:
        raise ValueError("ECM Data sheet is missing sheetData.")

    table_ref = str(target_table["ref"])
    _, start_row, end_col, end_row = _parse_range_ref(table_ref)
    if end_row - start_row < 4:
        raise ValueError(f"ECM table '{target_table_name}' does not have enough rows for expected mapping.")

    energy_row_number = start_row + 1
    demand_row_number = start_row + 2
    gas_energy_row_number = start_row + 3
    gas_demand_row_number = start_row + 4
    energy_row = _ensure_row(sheet_data, energy_row_number)
    demand_row = _ensure_row(sheet_data, demand_row_number)
    gas_energy_row = _ensure_row(sheet_data, gas_energy_row_number)
    gas_demand_row = _ensure_row(sheet_data, gas_demand_row_number)

    start_energy_col_idx = _column_to_index("B")
    end_energy_col_idx = _column_to_index(end_col)
    all_energy_columns = [_index_to_column(idx) for idx in range(start_energy_col_idx, end_energy_col_idx + 1)]
    for col in "BCDEFGHIJKLMNOPQRST":
        if col not in all_energy_columns:
            continue
        _set_numeric_cell(energy_row, f"{col}{energy_row_number}", None)
        _set_numeric_cell(gas_energy_row, f"{col}{gas_energy_row_number}", None)
        _set_numeric_cell(demand_row, f"{col}{demand_row_number}", None)
        _set_numeric_cell(gas_demand_row, f"{col}{gas_demand_row_number}", None)
    for col, value in elec_end_use_values_kbtu.items():
        _set_numeric_cell(energy_row, f"{col}{energy_row_number}", value)
    for col, value in gas_end_use_values_kbtu.items():
        _set_numeric_cell(gas_energy_row, f"{col}{gas_energy_row_number}", value)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
    _save_zip_file_map(file_map, output_workbook_path)
    return {
        "sheet": "ECM Data",
        "model_run_type": model_run_type,
        "target_table": target_table_name,
        "table_start_row": start_row,
        "electrical_energy_row": energy_row_number,
        "natural_gas_energy_row": gas_energy_row_number,
        "left_blank_columns": ECM_OPTIONAL_BLANK_COLUMNS,
        "end_use_columns_written": sorted(elec_end_use_values_kbtu.keys()),
        "output_workbook": str(output_workbook_path),
    }


def check_master_room_list_space_type_table_match(sim_text: str, workbook_path: Path) -> Dict[str, object]:
    """Compare LV-B spaces against existing Master Room List Space Type Table values."""
    lv_b_result = extract_lv_b_spaces(sim_text)
    expected_spaces = list(lv_b_result["spaces"].items())[:MASTER_ROOM_LIST_SPACE_MAX_ROWS]
    if load_workbook is not None:
        workbook = load_workbook(workbook_path, keep_vba=True, data_only=True)
        sheet = workbook["Master Room List"]
        mismatches: List[Dict[str, object]] = []
        compared_rows = 0
        for index in range(MASTER_ROOM_LIST_SPACE_MAX_ROWS):
            row_number = MASTER_ROOM_LIST_SPACE_START_ROW + index
            existing_name = sheet[f"D{row_number}"].value
            existing_area = sheet[f"G{row_number}"].value
            existing_name = str(existing_name).strip() if existing_name is not None else ""
            existing_area = float(existing_area) if existing_area is not None else None
            if index < len(expected_spaces):
                expected_name, expected_space_data = expected_spaces[index]
                expected_area = float(expected_space_data["area_sqft"])
                compared_rows += 1
                name_matches = existing_name == expected_name
                area_matches = existing_area is not None and abs(existing_area - expected_area) < 1e-6
                if not (name_matches and area_matches):
                    mismatches.append(
                        {
                            "row": row_number,
                            "expected_space_name": expected_name,
                            "existing_space_name": existing_name,
                            "expected_area_sqft": expected_area,
                            "existing_area_sqft": existing_area,
                        }
                    )
        return {
            "target_sheet": "Master Room List",
            "target_table": "Space Type Table",
            "writer": "openpyxl",
            "rows_checked": compared_rows,
            "spaces_compared": len(expected_spaces),
            "space_type_table_match": len(mismatches) == 0,
            "mismatch_count": len(mismatches),
            "mismatches": mismatches,
        }
    sheet_root = _load_master_room_list_sheet(workbook_path)
    sheet_data = sheet_root.find("m:sheetData", NS)
    if sheet_data is None:
        raise ValueError("Workbook sheet is missing sheetData.")
    mismatches: List[Dict[str, object]] = []
    compared_rows = 0
    for index in range(MASTER_ROOM_LIST_SPACE_MAX_ROWS):
        row_number = MASTER_ROOM_LIST_SPACE_START_ROW + index
        row = sheet_data.find(f"m:row[@r='{row_number}']", NS)
        existing_name = ""
        existing_area = None
        if row is not None:
            existing_name = _read_cell_text(row, f"D{row_number}").strip()
            existing_area = _read_cell_float(row, f"G{row_number}")
        if index < len(expected_spaces):
                        expected_name, expected_space_data = expected_spaces[index]
            expected_area = float(expected_space_data["area_sqft"])
            compared_rows += 1
            name_matches = existing_name == expected_name
            area_matches = existing_area is not None and abs(existing_area - expected_area) < 1e-6
            if not (name_matches and area_matches):
                mismatches.append(
                    {
                        "row": row_number,
                        "expected_space_name": expected_name,
                        "existing_space_name": existing_name,
                        "expected_area_sqft": expected_area,
                        "existing_area_sqft": existing_area,
                    }
                )
    return {
        "target_sheet": "Master Room List",
        "target_table": "Space Type Table",
        "rows_checked": compared_rows,
        "spaces_compared": len(expected_spaces),
        "space_type_table_match": len(mismatches) == 0,
        "mismatch_count": len(mismatches),
        "mismatches": mismatches,
    }


def resolve_model_run_type(cli_model_run_type: str | None) -> str:
    """Resolve model run type from CLI first, then environment, defaulting to Baseline."""
    if cli_model_run_type and cli_model_run_type.strip():
        return cli_model_run_type.strip()
    env_value = os.getenv("MODEL_RUN_TYPE", "").strip()
    if env_value:
        return env_value
    return "Baseline"
def extract_es_d_energy_cost_summary(sim_text: str) -> Dict[str, object]:
    """Extract utility-rate virtual rate, metered unit, and total charge from ES-D."""
    lines = sim_text.splitlines()
    in_esd = False
    utility_rates: Dict[str, Dict[str, object]] = {}
    for raw_line in lines:
        stripped = raw_line.strip()
        upper = stripped.upper()
        if "REPORT- ES-D" in upper:
            in_esd = True
            continue
        if in_esd and upper.startswith("REPORT-") and "REPORT- ES-D" not in upper:
            in_esd = False
        if not in_esd:
            continue
        if (
            not stripped
            or upper.startswith("UTILITY-RATE")
            or upper.startswith("METERED")
            or upper.startswith("ENERGY COST/")
            or set(stripped) <= {"-", "=", "+"}
        ):
            continue
        parts = [part.strip() for part in re.split(r"\s{2,}", stripped) if part.strip()]
        if len(parts) < 6:
            continue
        utility_rate = parts[0]
        metered_energy = parts[-4]
        total_charge_str = parts[-3]
        virtual_rate_str = parts[-2]
        metered_tokens = metered_energy.split()
        if len(metered_tokens) < 2:
            continue
        unit = metered_tokens[-1]
        try:
            total_charge = float(total_charge_str.replace(",", ""))
            virtual_rate = float(virtual_rate_str.replace(",", ""))
        except ValueError:
            continue
        utility_rates[utility_rate] = {
            "unit": unit,
            "total_charge": total_charge,
            "total_charge_unit": "$",
            "virtual_rate": virtual_rate,
            "virtual_rate_unit": "$/Unit",
        }
    if not utility_rates:
        raise ValueError("Could not parse ES-D utility-rate rows from the SIM file.")
    return {
        "report": "ES-D",
        "utility_rates": utility_rates,
    }
def extract_ps_h_details(sim_text: str) -> Dict[str, object]:
    """Extract PS-H loop, pump, and equipment sizing details."""
    lines = sim_text.splitlines()
    loops: Dict[str, Dict[str, object]] = {}
    pumps: Dict[str, Dict[str, object]] = {}
    equipment: Dict[str, Dict[str, object]] = {}
    report_indices = [i for i, line in enumerate(lines) if "REPORT- PS-H" in line.upper()]
    for start in report_indices:
        line = lines[start]
        name_match = re.search(r"REPORT-\s*PS-H\s+Loads and Energy Usage for\s+(.+?)\s+WEATHER FILE", line, re.IGNORECASE)
        if not name_match:
            continue
        report_name = " ".join(name_match.group(1).split())
        end = len(lines)
        for j in range(start + 1, len(lines)):
            if lines[j].strip().upper().startswith("REPORT-"):
                end = j
                break
        block = lines[start:end]
        block_text = "\n".join(block)
        if "DETAILED SIZING INFORMATION" in block_text:
            units_line_idx = next((i for i, b in enumerate(block) if "(MBTU/HR)" in b and "(HOURS)" in b and "(BTU/BTU)" in b), None)
            if units_line_idx is not None:
                units = re.findall(r"\(([^\)]+)\)", block[units_line_idx])
                row = None
                for k in range(units_line_idx + 1, min(units_line_idx + 20, len(block))):
                    if "----" in block[k]:
                        continue
                    parts = [part.strip() for part in re.split(r"\s{2,}", block[k].strip()) if part.strip()]
                    if len(parts) >= 8 and re.fullmatch(r"-?\d+(?:\.\d+)?", parts[1]):
                        row = parts
                        break
                if row:
                    equipment[report_name] = {
                        "capacity": float(row[1]),
                        "start_up": float(row[2]),
                        "electric": float(row[3]),
                        "heat_eir": float(row[4]),
                        "aux_elec": float(row[5]),
                        "fuel": float(row[6]),
                        "heat_hir": float(row[7]),
                        "units": {
                            "capacity": units[0] if len(units) > 0 else "MBTU/HR",
                            "start_up": units[1] if len(units) > 1 else "HOURS",
                            "electric": units[2] if len(units) > 2 else "KW",
                            "heat_eir": units[3] if len(units) > 3 else "BTU/BTU",
                            "aux_elec": units[4] if len(units) > 4 else "KW",
                            "fuel": units[5] if len(units) > 5 else "MBTU/HR",
                            "heat_hir": units[6] if len(units) > 6 else "BTU/BTU",
                        },
                    }
        elif "HEATING     COOLING      LOOP" in block_text:
            # first PS-H loop instance table at top
            units_line_idx = next((i for i, b in enumerate(block) if "(MBTU/HR)" in b and "(GPM" in b and "(FT)" in b), None)
            if units_line_idx is None:
                continue
            units = re.findall(r"\(([^\)]+)\)", block[units_line_idx])
            value_parts = None
            for k in range(units_line_idx + 1, min(units_line_idx + 12, len(block))):
                candidate = block[k].strip()
                if not candidate or '----' in candidate:
                    continue
                parts = candidate.split()
                if len(parts) >= 10 and all(re.fullmatch(r"-?\d+(?:\.\d+)?", p) for p in parts[:10]):
                    value_parts = parts[:10]
                    break
            if value_parts:
                loops[report_name] = {
                    "heating_capacity": float(value_parts[0]),
                    "cooling_capacity": float(value_parts[1]),
                    "loop_flow": float(value_parts[2]),
                    "total_head": float(value_parts[3]),
                    "loop_volume": float(value_parts[8]),
                    "units": {
                        "heating_capacity": units[0] if len(units) > 0 else "MBTU/HR",
                        "cooling_capacity": units[1] if len(units) > 1 else "MBTU/HR",
                        "loop_flow": units[2] if len(units) > 2 else "GPM",
                        "total_head": units[3] if len(units) > 3 else "FT",
                        "loop_volume": units[8] if len(units) > 8 else "GAL",
                    },
                }
        elif "CAPACITY               MECHANICAL" in block_text and "ATTACHED TO" in block_text:
            header_idx = next((i for i, b in enumerate(block) if "ATTACHED TO" in b and "(GPM" in b and "(KW)" in b), None)
            if header_idx is None:
                continue
            units_line = block[header_idx]
            units = re.findall(r"\(([^\)]+)\)", units_line)
            row = None
            for k in range(header_idx + 1, min(header_idx + 12, len(block))):
                candidate = block[k].strip()
                if not candidate or '----' in candidate:
                    continue
                parts = [part.strip() for part in re.split(r"\s{2,}", candidate) if part.strip()]
                if len(parts) >= 8 and re.fullmatch(r"-?\d+(?:\.\d+)?", parts[1]):
                    row = parts
                    break
            if row:
                pumps[report_name] = {
                    "attached_to": row[0],
                    "flow": float(row[1]),
                    "head": float(row[2]),
                    "capacity_control": row[4],
                    "power": float(row[5]),
                    "mechanical_efficiency": float(row[6]),
                    "motor_efficiency": float(row[7]),
                    "units": {
                        "flow": units[0] if len(units) > 0 else "GPM",
                        "head": units[1] if len(units) > 1 else "FT",
                        "capacity_control": "unitless",
                        "power": units[3] if len(units) > 3 else "KW",
                        "mechanical_efficiency": units[4] if len(units) > 4 else "FRAC",
                        "motor_efficiency": units[5] if len(units) > 5 else "FRAC",
                    },
                }
    if not (loops or pumps or equipment):
        raise ValueError("Could not parse PS-H loop/pump/equipment details from the SIM file.")
    return {
        "report": "PS-H",
        "loops": loops,
        "pumps": pumps,
        "equipment": equipment,
    }


def extract_schedule_table(sim_text: str) -> Dict[str, object]:
    """Extract schedule importer-style tabular rows from SIM text."""
    lines = sim_text.splitlines()
    header_idx = None
    headers: List[str] = []
    for idx, raw_line in enumerate(lines):
        stripped = raw_line.strip()
        upper = stripped.upper()
        if "SCHEDULE NAME" in upper and "SCHEDULE TYPE" in upper:
            header_idx = idx
            headers = [part.strip() for part in re.split(r"\s{2,}", stripped) if part.strip()]
            break
    if header_idx is None or not headers:
        raise ValueError("Could not find a schedule table header containing 'Schedule Name' and 'Schedule Type'.")
    rows: List[Dict[str, str]] = []
    for raw_line in lines[header_idx + 1 :]:
        stripped = raw_line.strip()
        upper = stripped.upper()
        if not stripped:
            continue
        if upper.startswith("REPORT-"):
            break
        if set(stripped) <= {"-", "=", "+"}:
            continue
        parts = [part.strip() for part in re.split(r"\s{2,}", stripped) if part.strip()]
        if len(parts) < len(headers):
            continue
        row_values = parts[: len(headers)]
        rows.append(dict(zip(headers, row_values)))
    if not rows:
        raise ValueError("Could not parse any schedule rows from the SIM schedule table.")
    return {
        "report": "SCHEDULE",
        "headers": headers,
        "rows": rows,
    }


def _normalize_column_name(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", value.upper())


def _get_schedule_row_value(schedule_row: Dict[str, str], table_column_name: str) -> str | None:
    normalized_column = _normalize_column_name(table_column_name)
    direct_lookup = {_normalize_column_name(key): value for key, value in schedule_row.items()}
    if normalized_column in direct_lookup:
        return direct_lookup[normalized_column]
    hourly_match = SCHEDULE_HOURLY_COLUMN_PATTERN.match(table_column_name.strip())
    if hourly_match:
        hour = str(int(hourly_match.group(1)))
        for candidate in (hour, f"{hour}AM", f"{hour}PM"):
            normalized_candidate = _normalize_column_name(candidate)
            if normalized_candidate in direct_lookup:
                return direct_lookup[normalized_candidate]
    return None


def populate_equest_schedule_importer_table(
    sim_text: str,
    workbook_path: Path,
    output_workbook_path: Path,
) -> Dict[str, object]:
    """Populate eQuest_Schedule_Importer table from extracted schedule rows."""
    schedule_result = extract_schedule_table(sim_text)
    schedule_rows = schedule_result["rows"]
    ET.register_namespace("", MAIN_NS)
    file_map = _load_zip_file_map(workbook_path)
    table_index = _build_table_index(file_map)
    schedule_table = table_index.get("eQuest_Schedule_Importer")
    if schedule_table is None:
        raise ValueError("Could not locate eQuest_Schedule_Importer table in workbook.")
    schedule_sheet_path = str(schedule_table["sheet_xml_path"])
    sheet_root = _parse_xml_with_registered_namespaces(file_map[schedule_sheet_path])
    sheet_data = sheet_root.find("m:sheetData", NS)
    if sheet_data is None:
        raise ValueError("Schedule sheet is missing sheetData.")
    table_ref = str(schedule_table["ref"])
    table_columns = list(schedule_table["columns"])
    start_col, start_row, _, end_row = _parse_range_ref(table_ref)
    start_col_index = _column_to_index(start_col)
    max_rows = end_row - start_row
    for row_index in range(max_rows):
        row_number = (start_row + 1) + row_index
        target_row = _ensure_row(sheet_data, row_number)
        current_source = schedule_rows[row_index] if row_index < len(schedule_rows) else {}
        for col_offset, table_column_name in enumerate(table_columns):
            column_letter = _index_to_column(start_col_index + col_offset)
            cell_ref = f"{column_letter}{row_number}"
            raw_value = _get_schedule_row_value(current_source, table_column_name) if current_source else None
            if raw_value is None or raw_value == "":
                _set_inline_string_cell(target_row, cell_ref, "")
                continue
            try:
                numeric_value = float(str(raw_value).replace(",", ""))
                _set_numeric_cell(target_row, cell_ref, numeric_value)
            except ValueError:
                _set_inline_string_cell(target_row, cell_ref, str(raw_value))
    file_map[schedule_sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
    _save_zip_file_map(file_map, output_workbook_path)
    return {
        "sheet": "eQuest Schedule Importer",
        "target_table": "eQuest_Schedule_Importer",
        "rows_available": max_rows,
        "rows_found": len(schedule_rows),
        "rows_written": min(len(schedule_rows), max_rows),
        "rows_truncated": max(len(schedule_rows) - max_rows, 0),
        "output_workbook": str(output_workbook_path),
    }


def _infer_space_name(lines: List[str], header_index: int) -> str:
    for idx in range(header_index - 1, max(header_index - 10, -1), -1):
        candidate = lines[idx].strip()
        upper = candidate.upper()
        if not candidate or set(candidate) <= {"-", "=", "+"}:
            continue
        if "REPORT-" in upper or "THERMOSTAT SETPOINT" in upper:
            continue
        explicit_match = re.search(r"\bSPACE\b\s*[:=\-]\s*(.+)$", candidate, re.IGNORECASE)
        if explicit_match:
            return " ".join(explicit_match.group(1).split())
        if upper.startswith("SPACE ") and len(candidate.split()) > 1:
            return " ".join(candidate.split()[1:])
        if re.search(r"[A-Z]", candidate, re.IGNORECASE) and not re.fullmatch(r"[\d\.\-\s:APM]+", candidate, re.IGNORECASE):
            return " ".join(candidate.split())
    return f"Space_{header_index}"


def _canonical_space_key(space_name: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", space_name.upper())


def extract_hourly_thermostat_setpoint_ranges(sim_text: str) -> Dict[str, object]:
    """Extract min/max thermostat setpoint temperatures by space from hourly sections."""
    lines = sim_text.splitlines()
    spaces_by_key: Dict[str, Dict[str, float | str]] = {}
    for idx, raw_line in enumerate(lines):
        if not THERMOSTAT_HEADER_PATTERN.search(raw_line):
            continue
        headers = [part.strip() for part in re.split(r"\s{2,}", raw_line.strip()) if part.strip()]
        thermostat_col_index = next(
            (i for i, header in enumerate(headers) if THERMOSTAT_HEADER_PATTERN.search(header)),
            None,
        )
        if thermostat_col_index is None:
            continue
        space_name = _infer_space_name(lines, idx)
        values: List[float] = []
        for data_line in lines[idx + 1 :]:
            stripped = data_line.strip()
            upper = stripped.upper()
            if not stripped:
                if values:
                    break
                continue
            if upper.startswith("REPORT-") or THERMOSTAT_HEADER_PATTERN.search(data_line):
                break
            if set(stripped) <= {"-", "=", "+"}:
                continue
            parts = [part.strip() for part in re.split(r"\s{2,}", stripped) if part.strip()]
            if len(parts) <= thermostat_col_index:
                continue
            token = parts[thermostat_col_index]
            number_match = re.search(r"-?\d+(?:\.\d+)?", token)
            if number_match:
                values.append(float(number_match.group(0)))
        if not values:
            continue
        space_key = _canonical_space_key(space_name)
        existing = spaces_by_key.get(space_key)
        min_value = min(values)
        max_value = max(values)
        if existing is None:
            spaces_by_key[space_key] = {
                "space_name": space_name,
                "min_thermostat_setpoint_f": min_value,
                "max_thermostat_setpoint_f": max_value,
            }
        else:
            existing["min_thermostat_setpoint_f"] = min(float(existing["min_thermostat_setpoint_f"]), min_value)
            existing["max_thermostat_setpoint_f"] = max(float(existing["max_thermostat_setpoint_f"]), max_value)
    if not spaces_by_key:
        raise ValueError("Could not find hourly thermostat setpoint data in the SIM text.")
    spaces = {
        str(payload["space_name"]): {
            "min_thermostat_setpoint_f": float(payload["min_thermostat_setpoint_f"]),
            "max_thermostat_setpoint_f": float(payload["max_thermostat_setpoint_f"]),
        }
        for payload in spaces_by_key.values()
    }
    return {
        "report": "HOURLY-THERMOSTAT-SETPOINT",
        "spaces": spaces,
        "space_count": len(spaces),
    }
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract BEPS, LV-B, LV-D, LV-I, LS-A, LV-M, ES-D, and PS-H report data from an eQuest SIM file."
    )
    parser.add_argument("sim_file", type=Path, help="Path to the eQuest .SIM file")
    parser.add_argument(
        "--report",
        choices=["beps", "lv-b", "lv-d", "lv-i", "ls-a", "lv-m", "es-d", "ps-h", "hourly-thermostat", "all"],
        default="all",
        help="Which report(s) to extract (default: all)",
    )
    parser.add_argument(
        "--indent",
        type=int,
        default=2,
        help="JSON indentation level for output (default: 2)",
    )
    parser.add_argument(
        "--populate-master-room-list",
        type=Path,
        help="Path to Building Performance Assumptions .xlsm used for Space Type Table actions.",
    )
    parser.add_argument(
        "--output-workbook",
        type=Path,
        help="Output path for updated workbook (required for Baseline writes).",
    )
    parser.add_argument(
        "--model-run-type",
        type=str,
        help="Model run type from automation input (e.g., Baseline, Proposed, ECM-1).",
    )
    parser.add_argument(
        "--update-ecm-data",
        type=Path,
        help="Path to workbook .xlsm where ECM Data should be populated from BEPS/ES-D.",
    )
    parser.add_argument(
        "--populate-schedules",
        type=Path,
        help="Path to workbook .xlsm where eQuest_Schedule_Importer should be populated.",
    )
    parser.add_argument(
        "--list-reports",
        action="store_true",
        help="List discovered REPORT-* sections in the SIM file and exit.",
    )
    args = parser.parse_args()
    sim_text = args.sim_file.read_text(errors="ignore")
    if args.list_reports:
        print(json.dumps(detect_available_reports(sim_text), indent=args.indent))
        return
    if args.update_ecm_data:
        if not args.output_workbook:
            raise ValueError("--output-workbook is required when using --update-ecm-data.")
        model_run_type = resolve_model_run_type(args.model_run_type)
        result = populate_ecm_data_from_reports(
            sim_text=sim_text,
            workbook_path=args.update_ecm_data,
            model_run_type=model_run_type,
            output_workbook_path=args.output_workbook,
        )
        print(json.dumps(result, indent=args.indent))
        return
    if args.populate_schedules:
        if not args.output_workbook:
            raise ValueError("--output-workbook is required when using --populate-schedules.")
        result = populate_equest_schedule_importer_table(
            sim_text=sim_text,
            workbook_path=args.populate_schedules,
            output_workbook_path=args.output_workbook,
        )
        print(json.dumps(result, indent=args.indent))
        return
    if args.populate_master_room_list:
        model_run_type = resolve_model_run_type(args.model_run_type)
        normalized_model_run_type = model_run_type.upper()
        is_baseline = normalized_model_run_type == "BASELINE"
        if is_baseline:
            if not args.output_workbook:
                raise ValueError("--output-workbook is required for Baseline when using --populate-master-room-list.")
            result = populate_master_room_list_space_type_table(
                sim_text=sim_text,
                workbook_path=args.populate_master_room_list,
                output_workbook_path=args.output_workbook,
            )
            result.update(
                {
                    "model_run_type": model_run_type,
                    "is_baseline": True,
                    "space_type_table_match": True,
                }
            )
        else:
            comparison_result = check_master_room_list_space_type_table_match(
                sim_text=sim_text,
                workbook_path=args.populate_master_room_list,
            )
            result = {
                "model_run_type": model_run_type,
                "is_baseline": False,
                **comparison_result,
            }
        print(json.dumps(result, indent=args.indent))
        return
    if args.report == "beps":
        result = extract_beps_report(sim_text)
    elif args.report == "lv-b":
        result = extract_lv_b_spaces(sim_text)
    elif args.report == "lv-d":
        result = extract_lv_d_report(sim_text)
    elif args.report == "lv-i":
        result = extract_lv_i_constructions(sim_text)
    elif args.report == "ls-a":
        result = extract_ls_a_peak_loads(sim_text)
    elif args.report == "lv-m":
        result = extract_lv_m_conversions(sim_text)
    elif args.report == "es-d":
        result = extract_es_d_energy_cost_summary(sim_text)
    elif args.report == "ps-h":
        result = extract_ps_h_details(sim_text)
    elif args.report == "hourly-thermostat":
        result = extract_hourly_thermostat_setpoint_ranges(sim_text)
    else:
        lv_b_result = extract_lv_b_spaces(sim_text)
        lv_m_result = extract_lv_m_conversions(sim_text)
        result = {
            "beps": extract_beps_report(sim_text),
            "lv_b_spaces": lv_b_result,
            "lv_d": extract_lv_d_report(sim_text),
            "lv_i": extract_lv_i_constructions(sim_text),
            "ls_a_peak_loads": extract_ls_a_peak_loads(sim_text, lv_b_result=lv_b_result),
            "lv_m": lv_m_result,
            "es_d": extract_es_d_energy_cost_summary(sim_text),
            "ps_h": extract_ps_h_details(sim_text),
            "hourly_thermostat_setpoints": extract_hourly_thermostat_setpoint_ranges(sim_text),
        }
    print(json.dumps(result, indent=args.indent))
if __name__ == "__main__":
    main()
