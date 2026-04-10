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
- **`UnicodeDecodeError` when reading JSON config files**
  - Power Automate or editor tools may save JSON as UTF-16 with BOM.
  - `run_local.py` and Graph config loading now accept UTF-8 (with/without BOM) and UTF-16 JSON.
  - If errors persist, re-save the JSON file as UTF-8 and validate JSON syntax.

## PyCharm + OneDrive compatibility

Yes — this project works with files stored in OneDrive in two common ways:

1. **Direct Graph access** using `onedrive:/...` paths (recommended for scripted runs).
2. **Locally synced OneDrive folder** paths such as `C:/Users/<you>/OneDrive/...`.

No code changes are required specifically for PyCharm if you use valid file paths.

### Recommended setup for PyCharm

- Keep the **project folder** in OneDrive if you want it synced and available across machines.
- Prefer keeping the **virtual environment (`.venv`) outside OneDrive** to reduce file-lock/sync conflicts and improve indexing performance.
- If you do keep `.venv` inside OneDrive, pause sync while creating/updating dependencies and resume afterward.

### Move an existing PyCharm project into OneDrive

1. Close PyCharm.
2. Move the project directory to your OneDrive folder (example):
   - From: `C:\Users\<you>\PycharmProjects\eQuest-Simulation-Extraction`
   - To: `C:\Users\<you>\OneDrive\PycharmProjects\eQuest-Simulation-Extraction`
3. Re-open the project from the new OneDrive path in PyCharm.
4. In PyCharm, verify interpreter settings:
   - **Settings → Project → Python Interpreter**
   - Re-select your interpreter if the old path is no longer valid.
5. Run a quick smoke test from the PyCharm terminal:

```bash
python equest_extractor.py --help
```

### Clone directly into OneDrive so repo files live there

```bash
cd "C:\Users\<you>\OneDrive\PycharmProjects"
git clone <your-repo-url> eQuest-Simulation-Extraction
cd eQuest-Simulation-Extraction
python -m venv .venv
```

If you prefer the safest dependency behavior, create the virtualenv outside OneDrive and point PyCharm to it.

### Power Automate workflow when PyCharm project/repo is **not** in OneDrive

You can keep the repo outside OneDrive (for example `D:\dev\eQuest-Simulation-Extraction`) and still
use Power Automate to update JSON inputs and run Python.

#### Pattern that works well

1. Store your automation input JSON somewhere writable by your flow (can be outside OneDrive), e.g.:
   - `D:\automation\local_inputs.json`
2. In that JSON, point data files to OneDrive using either:
   - `onedrive:/...` paths (Graph API mode), or
   - local synced OneDrive paths (e.g., `C:/Users/<you>/OneDrive/...`).
3. Have Power Automate update only the fields it needs (`sim_file`, `workbook_path`, `model_run_type`, etc.).
4. Run `run_local.py` and pass the JSON file path explicitly.

Example command from a Power Automate “Run a command line”/Desktop step:

```bash
cd /d D:\dev\eQuest-Simulation-Extraction
python run_local.py D:\automation\local_inputs.json
```

#### Example `local_inputs.json` for this setup

```json
{
  "mode": "combined",
  "sim_file": "onedrive:/Projects/eQuest/MyModel.SIM",
  "workbook_path": "onedrive:/Projects/eQuest/Building Performance Assumptions.xlsm",
  "output_workbook_path": "onedrive:/Projects/eQuest/Building Performance Assumptions.updated.xlsm",
  "model_run_type": "Baseline",
  "graph_config_path": "D:/automation/local_graph_inputs.json"
}
```

#### Why this works

- `run_local.py` accepts a custom config-file path argument, so the JSON can live anywhere.  
- `run_local.py` also resolves relative `graph_config_path` values against that config file location and exports `GRAPH_CONFIG_PATH` for the child process.

So Power Automate can edit a JSON outside OneDrive, then launch Python in the repo folder, and the extractor will still read/write OneDrive files correctly.
