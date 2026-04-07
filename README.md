# eQuest-Simulation-Extraction
This is a python script that can be used to output data from the simulation from eQuest simulation files.

This repository includes a Python utility (`equest_extractor.py`) to extract **BEPS**, **LV-B**, **LV-D**, **LV-I**, **LS-A**, **LV-M**, **ES-D**, and **PS-H** report data from an eQuest `.SIM` file.

## Requirements

- Python 3.9+ (Anaconda Python is fine)
- `msal` if you want to read/write files from OneDrive for Business through Microsoft Graph.
  - `pip install msal`
- Recommended: `openpyxl` for safest `.xlsm` read/write behavior.
  - `pip install openpyxl`
- Without `openpyxl`, the tool falls back to XML-level editing.


## BEPS extraction
- Extracts BEPS fuel rows and totals by fuel/end-use.

## LV-B extraction
- Extracts unique spaces and key space attributes.

## LV-D extraction
- Extracts the final LV-D orientation summary table.

## LV-I extraction
- Extracts each construction name with:
  - `u_value`
  - `number_of_response_factors`
- Returns `u_value_unit` (e.g., `BTU/HR-SQFT-F`).

## LS-A extraction
- Extracts peak load values by space from `REPORT- LS-A Space Peak Loads Summary`:
  - `cooling_load`
  - `heating_load`
- Stores load units (e.g., `KBTU/HR`).
- Associates these peak loads with the spaces extracted from LV-B (`spaces_with_peak_loads`).

## LV-M extraction
- Extracts conversion factors from `REPORT- LV-M DOE-2.2 Units Conversion Table`.
- Stores conversions in a dictionary for later use.
- Includes helper `convert_value(value, from_unit, to_unit, conversions)` to convert values now or later.

## ES-D extraction
- Extracts utility-rate summary values from `REPORT- ES-D Energy Cost Summary`.
- Returns for each utility-rate:
  - `virtual_rate` and `virtual_rate_unit` (`$/Unit`)
  - `unit` from the `METERED ENERGY UNITS/YR` field (e.g., `KWH`, `THERM`)
  - `total_charge` and `total_charge_unit` (`$`)

## PS-H extraction
- Extracts equipment/loop/pump-level PS-H details.
- Loops: heating capacity, cooling capacity, loop flow, total head, loop volume + units.
- Pumps: attached-to, flow, head, capacity control, power, mechanical efficiency, motor efficiency + units.
- Equipment (from the second PS-H instance detailed sizing): capacity, start-up, electric, heat EIR, aux elec, fuel, heat HIR + units.

## Usage

```bash
python equest_extractor.py /path/to/file.SIM
python equest_extractor.py /path/to/file.SIM --report beps
python equest_extractor.py /path/to/file.SIM --report lv-b
python equest_extractor.py /path/to/file.SIM --report lv-d
python equest_extractor.py /path/to/file.SIM --report lv-i
python equest_extractor.py /path/to/file.SIM --report ls-a
python equest_extractor.py /path/to/file.SIM --report lv-m
python equest_extractor.py /path/to/file.SIM --report es-d
python equest_extractor.py /path/to/file.SIM --report ps-h
python equest_extractor.py /path/to/file.SIM --report all

# Populate Master Room List > Space Type Table (Space Name + Area)
python equest_extractor.py /path/to/file.SIM \
  --populate-master-room-list /path/to/Building\ Performance\ Assumptions.xlsm \
  --model-run-type Baseline \
  --output-workbook /path/to/Building\ Performance\ Assumptions.updated.xlsm

# Validate against existing Baseline data (returns space_type_table_match true/false)
python equest_extractor.py /path/to/file.SIM \
  --populate-master-room-list /path/to/Building\ Performance\ Assumptions.xlsm \
  --model-run-type ECM-1

# Populate ECM Data section for model run type (Baseline/Proposed/ECM-1..ECM-7)
python equest_extractor.py /path/to/file.SIM \
  --update-ecm-data /path/to/Building\ Performance\ Assumptions.xlsm \
  --model-run-type ECM-3 \
  --output-workbook /path/to/Building\ Performance\ Assumptions.updated.xlsm
```

## OneDrive for Business (Microsoft Graph)

You can now pass OneDrive paths anywhere a file path is accepted by using the `onedrive:` prefix.

### Environment variables

Set these before running:

- `GRAPH_CLIENT_ID` (required)
- `GRAPH_TENANT_ID` (optional, default: `organizations`)
- `GRAPH_CLIENT_SECRET` (optional; if omitted, device code login is used)
- `GRAPH_USER_ID` (recommended for app-only auth; target user UPN or object ID)

### Using `run_local.py` with saved local Graph settings

Use `graph_config_path` in `local_inputs.json` to point to a local JSON file (for example, a OneDrive-synced local path) that stores Graph credentials.

`run_local.py` exports this path as `GRAPH_CONFIG_PATH`, and `ms_graph.py` will load `client_id`, `tenant_id`, `client_secret` (or `app_secret`), and `user_id` from that file.
Environment variables still override file values if both are provided.

Example:

```json
{
  "mode": "combined",
  "sim_file": "onedrive:/Projects/eQuest/MyModel.SIM",
  "workbook_path": "onedrive:/Projects/eQuest/Building Performance Assumptions.xlsm",
  "output_workbook_path": "onedrive:/Projects/eQuest/Building Performance Assumptions.updated.xlsm",
  "model_run_type": "Baseline",
  "graph_config_path": "C:/Users/you/OneDrive/local_graph_inputs.json"
}
```

`local_graph_inputs.json` example:

```json
{
  "client_id": "YOUR_APP_CLIENT_ID",
  "tenant_id": "organizations",
  "app_secret": "YOUR_APP_CLIENT_SECRET",
  "user_id": "user@yourtenant.com"
}
```

If `graph_config_path` is omitted, behavior remains env-vars-only.

### Examples

```bash
# Read SIM from OneDrive, return JSON in console.
python equest_extractor.py "onedrive:/Projects/eQuest/MyModel.SIM" --report all

# Read workbook from OneDrive and upload output workbook back to OneDrive.
python equest_extractor.py "onedrive:/Projects/eQuest/MyModel.SIM" \
  --update-ecm-data "onedrive:/Projects/eQuest/Building Performance Assumptions.xlsm" \
  --model-run-type Baseline \
  --output-workbook "onedrive:/Projects/eQuest/Building Performance Assumptions.updated.xlsm"
```

### Troubleshooting

- **`Errno 13: Permission denied` on OneDrive files**
  - Close Excel and any other applications that may currently have the workbook or output file open.
  - Close any background automation/processes tied to this workflow that may have an open file handle.
  - Re-run the command after all associated files are closed.
