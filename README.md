# eQuest-Simulation-Extraction
This is a python script that can be used to output data from the simulation from eQuest simulation files.

This repository includes a Python utility (`equest_extractor.py`) to extract **BEPS**, **LV-B**, **LV-D**, **LV-I**, **LS-A**, **LV-M**, **ES-D**, and **PS-H** report data from an eQuest `.SIM` file.

## Requirements

- Python 3.9+ (Anaconda Python is fine)
- Recommended: `openpyxl` for safest `.xlsm` read/write behavior.
  - `pip install openpyxl`
- Without `openpyxl`, the tool falls back to XML-level editing.

## 2-minute sanity check

Run these commands in order to verify your environment before PAD:

```bash
python --version
python equest_extractor.py --help
python equest_extractor.py "<SIM_PATH>" --list-reports
python equest_extractor.py "<SIM_PATH>" --report beps
python equest_extractor.py "<SIM_PATH>" --populate-master-room-list "<XLSM_PATH>" --model-run-type Baseline --output-workbook "<OUTPUT_XLSM_PATH>"
python equest_extractor.py "<SIM_PATH>" --update-ecm-data "<XLSM_PATH>" --model-run-type ECM-1 --output-workbook "<OUTPUT_XLSM_PATH>"
```

Expected quick checks:
- `--help` prints usage
- `--list-reports` shows what REPORT sections actually exist in the SIM
- `--report beps` prints JSON
- workbook commands print JSON and produce/update output `.xlsm`

## PyCharm local testing (manual inputs now, dynamic later)

If you want to test locally in PyCharm first, use a config-driven runner:

1. Copy `local_inputs.template.json` to `local_inputs.json`.
2. Edit paths and variables in `local_inputs.json`.
3. Run:
   ```bash
   python run_local.py
   ```
   or pass an explicit config file:
   ```bash
   python run_local.py path/to/your/local_inputs.json
   ```

### `local_inputs.json` fields
- `mode`: one of `extract_report`, `master_room_list`, `ecm_data`
- `sim_file`: path to `.SIM`
- `workbook_path`: path to workbook (required for `master_room_list` and `ecm_data`)
- `output_workbook_path`: output path (required for `master_room_list` and `ecm_data`)
- `model_run_type`: `Baseline`, `Proposed`, `ECM-1`...`ECM-7`
- `report`: report name if `mode` is `extract_report` (for example `beps`, `all`)

This gives you variables you can change now in one file and later replace with user-driven inputs from UI/automation.

For a one-shot local verification with exact Windows paths, use:
- `local_test_sequence.ps1`

Run it from PowerShell (not Python):
```powershell
powershell -ExecutionPolicy Bypass -File .\local_test_sequence.ps1
```

If you run a `.ps1` file with `python ...`, you can get errors like `Unknown option: -F`.

`local_test_sequence.ps1` now checks whether the SIM contains `REPORT- BEPS`:
- if present: runs BEPS check + ECM update
- if missing: skips BEPS/ECM steps and still runs Master Room List population

It also supports Power Automate-style workbook actions against **Master Room List → Space Type Table** in `Building Performance Assumptions.xlsm` using `--model-run-type`:
- `Baseline`: writes LV-B space names and areas into the table.
- Any non-baseline type (e.g., `Proposed`, `ECM-1`): compares LV-B space names/areas against existing table values and returns a boolean match result.

It also supports updating **ECM Data** tables from BEPS + ES-D based on `--model-run-type` for:
- `Baseline`
- `Proposed`
- `ECM-1` to `ECM-7`

`Baseline-2` and `Baseline-3` are intentionally ignored.

In ECM Data population, these columns are intentionally left blank unless future mapping is provided:
`Fans Process`, `Fans Parking Garage`, `Data Centre Equipment`, `Cooking`, `Elevators/Escalators`, `CHP`, `Humidification`, `Other Processes`.

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

## Power Automate + Teams card input

If you collect **Model Run Type** from a Teams Adaptive Card, pass it directly into Python using either CLI args or an environment variable:

1. In Power Automate, add **Post adaptive card and wait for a response**.
2. Read the response field (for example: `modelRunType`).
3. In the Python execution step, pass that value as:

### Option A: CLI argument (recommended)
```bash
python equest_extractor.py "<SIM_PATH>" \
  --populate-master-room-list "<XLSM_PATH>" \
  --model-run-type "<MODEL_RUN_TYPE_FROM_TEAMS>" \
  --output-workbook "<OUTPUT_XLSM_IF_BASELINE>"
```

### Option B: Environment variable
Set environment variable `MODEL_RUN_TYPE` in the Power Automate step and run:
```bash
python equest_extractor.py "<SIM_PATH>" \
  --populate-master-room-list "<XLSM_PATH>" \
  --output-workbook "<OUTPUT_XLSM_IF_BASELINE>"
```

`equest_extractor.py` resolves model run type in this order:  
1) `--model-run-type` argument, 2) `MODEL_RUN_TYPE` env var, 3) default `"Baseline"`.

## End-to-end: GitHub -> Power Automate -> Run Python

### 1) Store your files in GitHub
Include these files in your repository:
- `equest_extractor.py`

`.SIM` files and `Building Performance Assumptions.xlsm` **do not need to be stored in GitHub**. They can be local files provided by the user at runtime (for example from a local/network folder, SharePoint sync folder, or OneDrive local path on the runner machine).

### 2) Make files reachable from Power Automate
Power Automate cannot run Python directly in cloud flows without a runner. Use one of:
- **Power Automate Desktop (PAD) on a Windows machine** (recommended)
- **Azure Automation / VM / Function** that can run Python

For PAD approach:
1. Install Git on the runner machine.
2. Clone your repo to a stable folder, for example:
   - `C:\\Automation\\eQuest-Simulation-Extraction`
3. Ensure Python is installed and available in PATH.

### 3) Build your cloud flow
1. Trigger: manual, Forms, SharePoint, email, etc.
2. Add Teams action: **Post adaptive card and wait for a response**.
3. Capture fields like:
   - `modelRunType` (Baseline / Proposed / ECM-1..ECM-7)
   - SIM file name/path (if dynamic)

### 4) Call PAD (or your runner)
In the cloud flow, add **Run a flow built with Power Automate for desktop** and pass:
- `repoPath`
- `simPath`
- `workbookPath`
- `outputWorkbookPath`
- `modelRunType`

### 5) In PAD, run Python from the repo folder
Example command for ECM Data update:
```bash
python equest_extractor.py "<SIM_PATH>" \
  --update-ecm-data "<XLSM_PATH>" \
  --model-run-type "<MODEL_RUN_TYPE>" \
  --output-workbook "<OUTPUT_XLSM_PATH>"
```

Example command for Master Room List flow:
```bash
python equest_extractor.py "<SIM_PATH>" \
  --populate-master-room-list "<XLSM_PATH>" \
  --model-run-type "<MODEL_RUN_TYPE>" \
  --output-workbook "<OUTPUT_XLSM_PATH>"
```

### 6) Capture JSON output in PAD and return to cloud flow
`equest_extractor.py` prints JSON to stdout. In PAD:
1. Capture command output to a variable.
2. Return output to cloud flow.
3. In cloud flow, parse JSON for booleans/metrics (for example `space_type_table_match`, totals, costs).

### 7) Write output workbook to your destination
After Python runs:
- Upload the updated workbook to SharePoint/OneDrive, or
- Commit/push back to GitHub via scripted step if desired.

### 8) Error handling recommendations
- If Python exit code != 0, fail the flow and notify Teams/email.
- Log command, model run type, and file paths for diagnostics.
- Keep a timestamped output filename to avoid overwriting prior runs.

## Power Automate Desktop: Run Python Script inputs

Yes — the extractor is ready for this pattern.

When using **Run Python Script** in Power Automate Desktop, create PAD variables and pass them as script arguments:

### PAD variables to create
- `%SimFilePath%` (example: `C:\Automation\inputs\project.sim`)
- `%WorkbookPath%` (example: `C:\Automation\inputs\Building Performance Assumptions.xlsm`)
- `%ModelRunType%` (example: `Baseline`, `Proposed`, `ECM-1`)
- `%OutputWorkbookPath%` (example: `C:\Automation\outputs\BPA.updated.xlsm`)
- `%RepoPath%` (example: `C:\Automation\eQuest-Simulation-Extraction`)

### Recommended PAD sequence
1. **Set Current Directory** to `%RepoPath%`.
2. **Run Python Script** (or Run application with python executable) using one of these command patterns.

#### A) Update ECM Data
```bash
python equest_extractor.py "%SimFilePath%" --update-ecm-data "%WorkbookPath%" --model-run-type "%ModelRunType%" --output-workbook "%OutputWorkbookPath%"
```

#### B) Update Master Room List / validate by model type
```bash
python equest_extractor.py "%SimFilePath%" --populate-master-room-list "%WorkbookPath%" --model-run-type "%ModelRunType%" --output-workbook "%OutputWorkbookPath%"
```

### How Python consumes these inputs
- `sim_file` positional arg -> `%SimFilePath%`
- `--update-ecm-data` or `--populate-master-room-list` -> `%WorkbookPath%`
- `--model-run-type` -> `%ModelRunType%`
- `--output-workbook` -> `%OutputWorkbookPath%`

The script prints JSON to stdout; capture this output in PAD and pass it back to your cloud flow if needed.

### PAD "Run Python Script" action (large text box)

If PAD gives you a **Python Script to Run** text box (instead of a command line), paste a wrapper script that calls `equest_extractor.py` with your variables.

`%RepoPath%` should be the local folder where you cloned/downloaded this repo on the runner machine, for example:
- `A:\Users\kyleg\OneDrive - ICR Engineering\Documents\Energy Modeling\Automation\eQuest-Simulation-Extraction`

Example wrapper script for PAD:

Paste only the code body below into PAD (do **not** include the triple-backtick lines like ```python or ```):

```python
import json
import os
import subprocess
import sys

repo_path = r"%RepoPath%"
sim_file_path = r"%SimFilePath%"
workbook_path = r"%WorkbookPath%"
model_run_type = r"%ModelRunType%"
output_workbook_path = r"%OutputWorkbookPath%"

command = [
    sys.executable,
    "equest_extractor.py",
    sim_file_path,
    "--populate-master-room-list",
    workbook_path,
    "--model-run-type",
    model_run_type,
    "--output-workbook",
    output_workbook_path,
]

result = subprocess.run(command, cwd=repo_path, capture_output=True, text=True)
print(result.stdout)
if result.returncode != 0:
    print(result.stderr, file=sys.stderr)
    raise SystemExit(result.returncode)
```

Notes:
- Keep `modelRunType` as `Baseline` (no extra quotes in the value).
- It is okay if `WorkbookPath` and `OutputWorkbookPath` are the same path if you want in-place updates.
- If you see `SyntaxError` pointing at `````python``, it means markdown code fences were pasted into PAD — remove them.

### Troubleshooting PAD runtime issues

If PAD shows:
- `ImportError: No module named 'json'`

then PAD is likely not running the expected Python 3 interpreter.

Use this approach:
1. In PAD, run a diagnostic script:
   ```python
   import sys
   print(sys.executable)
   print(sys.version)
   ```
2. Confirm the interpreter is a normal Python 3 install (for example `C:\Users\...\AppData\Local\Programs\Python\Python311\python.exe`).
3. If PAD is using a different runtime, switch to **Run application** action and call your known Python executable explicitly:
   ```bash
   "C:\Path\To\Python\python.exe" "C:\Path\To\eQuest-Simulation-Extraction\equest_extractor.py" "<SIM_PATH>" --populate-master-room-list "<XLSM_PATH>" --model-run-type "<MODEL_RUN_TYPE>" --output-workbook "<OUTPUT_XLSM_PATH>"
   ```
4. Ensure `%RepoPath%` points to the folder containing `equest_extractor.py` and not just a parent folder.
5. Capture and review both stdout and stderr in PAD; the extractor prints JSON to stdout on success.
