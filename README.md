# eQuest-Simulation-Extraction

Python tooling to extract eQuest `.SIM` report data and optionally populate/update sections of the **Building Performance Assumptions** workbook (`.xlsm`).

This README is intended as a practical handoff guide so a new teammate can set up and run the workflow in:

- Local terminal / PyCharm
- Power Automate Desktop (PAD)
- Power Automate Cloud (triggering PAD desktop flows)

---

## What this project does

The core script is `equest_extractor.py`. It can:

1. **Extract report data from a `.SIM` file**
   - Supported reports: `BEPS`, `LV-B`, `LV-D`, `LV-I`, `LS-A`, `LV-M`, `ES-D`, `PS-H`.
2. **Populate Master Room List > Space Type Table** in a target workbook.
3. **Populate ECM Data** in a target workbook.
4. Read/write files from either:
   - local disk paths, or
   - OneDrive via Microsoft Graph using `onedrive:/...` paths.

A helper runner (`run_local.py`) executes the extractor using a JSON config (default `local_inputs.json`) and supports multi-step automation mode (`combined`).

---

## Repository files

Current key files:

- `equest_extractor.py` – main CLI and workbook/report processing logic.
- `run_local.py` – local/PAD-friendly JSON-config runner.
- `ms_graph.py` – Microsoft Graph helper for OneDrive read/write.
- `local_inputs.json` – example local run configuration.
- `README.md` – setup and usage instructions.

---

## Prerequisites

## 1) Install Python

- Python **3.9+** recommended.
- Verify:

```bash
python --version
```

## 2) Install project dependencies

At minimum:

```bash
pip install msal openpyxl
```

Notes:
- `msal` is required for Graph/OneDrive API mode.
- `openpyxl` is recommended for safer `.xlsm` handling.

## 3) eQuest assets

You need:
- an input `.SIM` file,
- and for workbook workflows, a `Building Performance Assumptions.xlsm` file.

---

## PyCharm setup (Windows)

1. Install **PyCharm Community** or **Professional**.
2. Open the project folder in PyCharm.
3. Configure interpreter:
   - `File > Settings > Project > Python Interpreter`
   - Create/select a virtual environment.
4. Install dependencies in that interpreter:

```bash
pip install msal openpyxl
```

5. Verify from PyCharm terminal:

```bash
python equest_extractor.py --help
python run_local.py
```

If `run_local.py` errors with missing config, either:
- create/update `local_inputs.json`, or
- pass your own JSON path: `python run_local.py C:\path\to\inputs.json`.

---

## Local CLI usage

### A) Extract report data only

```bash
python equest_extractor.py C:\path\to\model.SIM --report all
python equest_extractor.py C:\path\to\model.SIM --report beps
python equest_extractor.py C:\path\to\model.SIM --list-reports
```

### B) Populate Master Room List

```bash
python equest_extractor.py C:\path\to\model.SIM \
  --populate-master-room-list C:\path\to\Building\ Performance\ Assumptions.xlsm \
  --model-run-type Baseline \
  --output-workbook C:\path\to\Building\ Performance\ Assumptions.updated.xlsm
```

### C) Populate ECM Data

```bash
python equest_extractor.py C:\path\to\model.SIM \
  --update-ecm-data C:\path\to\Building\ Performance\ Assumptions.xlsm \
  --model-run-type ECM-3 \
  --output-workbook C:\path\to\Building\ Performance\ Assumptions.updated.xlsm
```

---

## `run_local.py` JSON runner

`run_local.py` allows PAD/automation-friendly execution with one JSON file.

Default behavior:

```bash
python run_local.py
```

This reads `local_inputs.json` from repo root.

Custom path:

```bash
python run_local.py D:\automation\local_inputs.json
```

### Supported `mode` values

- `extract_report`
- `master_room_list`
- `ecm_data`
- `combined` (runs Master Room List then ECM Data in sequence)

### Example `local_inputs.json`

```json
{
  "mode": "combined",
  "sim_file": "C:/data/model.SIM",
  "workbook_path": "C:/data/Building Performance Assumptions.xlsm",
  "output_workbook_path": "C:/data/Building Performance Assumptions.updated.xlsm",
  "model_run_type": "Baseline",
  "report": "all",
  "graph_config_path": null
}
```

---

## OneDrive / Microsoft Graph mode (optional)

You can use `onedrive:/...` paths anywhere file paths are accepted.

Example:

```bash
python equest_extractor.py "onedrive:/Projects/eQuest/MyModel.SIM" --report all
```

Set Graph auth via environment variables or config file:

- `GRAPH_CLIENT_ID` (required)
- `GRAPH_TENANT_ID` (optional, default `organizations`)
- `GRAPH_CLIENT_SECRET` (optional)
- `GRAPH_USER_ID` (recommended for app-only auth)

You can also provide `graph_config_path` in JSON; `run_local.py` exports it to `GRAPH_CONFIG_PATH` for the child process.

---

## Power Automate Desktop (PAD) setup

This section is the most important for reliable unattended runs.

### 1) Install PAD

- Install **Power Automate Desktop** on the machine where Python runs.
- Sign in with your work account.

### 2) Create a Desktop flow

Recommended actions:

1. **Set Variable** actions for:
   - `RepoPath` (e.g., `C:\Users\<you>\PycharmProjects\eQuest-Simulation-Extraction`)
   - `ConfigPath` (e.g., `D:\automation\local_inputs.json`)

2. **Run application** action:
   - Application path:
     - Use your PyCharm venv python.exe path, example:
       `C:\Users\<you>\PycharmProjects\eQuest-Simulation-Extraction\.venv\Scripts\python.exe`
   - Command line arguments:
     - `"C:\Users\<you>\PycharmProjects\eQuest-Simulation-Extraction\run_local.py" "D:\automation\local_inputs.json"`
   - Working folder:
     - `C:\Users\<you>\PycharmProjects\eQuest-Simulation-Extraction`
   - Wait for app to complete: **Yes**

3. Add logging/error handling:
   - Capture `%ExitCode%` from PAD action.
   - If non-zero, write failure details to log/notification.

### 3) Optional: capture stdout/stderr in PAD

Use `cmd.exe` wrapper if needed:

- Application path: `C:\Windows\System32\cmd.exe`
- Arguments:

```bat
/c ""C:\...\python.exe" "C:\...\run_local.py" "D:\automation\local_inputs.json" > "C:\temp\pad_stdout.txt" 2> "C:\temp\pad_stderr.txt""
```

This helps diagnose generic exit code 1 failures.

### PAD reliability checklist

- Use **absolute paths** only.
- Use the **same interpreter** as PyCharm.
- Ensure working folder = repository root.
- Do not keep workbook open in Excel during update steps.
- If OneDrive sync locks files, retry after sync settles.

---

## Power Automate Cloud + PAD integration

Typical architecture:

1. Cloud flow is triggered (manual, schedule, SharePoint, etc.).
2. Cloud flow calls **Run a desktop flow**.
3. Desktop flow runs `python run_local.py <config>` on machine/gateway.
4. Desktop flow returns success/failure and optional output metadata.

### Cloud flow setup outline

1. Create a cloud flow (instant/scheduled/automated).
2. Add **Run a desktop flow** action.
3. Choose target machine/machine group and desktop flow name.
4. Pass inputs (if needed) that PAD writes into `local_inputs.json` before launch.
5. Add post-run actions (Teams/email/Dataverse/SharePoint updates) based on result.

### Suggested parameter pattern

- Cloud flow input parameters:
  - `sim_file`
  - `workbook_path`
  - `output_workbook_path`
  - `model_run_type`
- PAD updates JSON file with those values.
- PAD runs `run_local.py`.

This gives reusable orchestration without hardcoding each run in PAD.

---

## Troubleshooting

### Exit code 1 in PAD but works in PyCharm

Usually caused by one of:
- wrong Python interpreter path,
- wrong working directory,
- missing environment variables,
- invalid/relative file paths,
- file lock (Excel/OneDrive),
- JSON encoding issue.

Quick checks:

```bash
python --version
where python
python -c "import sys,os; print(sys.executable); print(os.getcwd())"
```

Run those in both PyCharm terminal and PAD context and compare outputs.

### `Errno 13 Permission denied`

- Close Excel/workbook.
- Ensure output file is not open or locked by another process.

### JSON parse issues

- Save JSON as UTF-8 or UTF-16 valid JSON.
- Avoid trailing commas.

---

## Recommended team handoff convention

For smoother coworker onboarding:

1. Keep project in source control.
2. Keep local machine-specific config outside repo, e.g. `D:\automation\local_inputs.json`.
3. Use `run_local.py <external_config>` in PAD.
4. Document machine-specific interpreter path in team runbook.
5. Validate once in PyCharm before enabling cloud-triggered unattended runs.

---

## Security notes

- Do not commit secrets (client secret, access tokens, user credentials).
- Prefer environment variables or secure secrets store for Graph credentials.
- Restrict file permissions for config files containing sensitive values.

---

## Quick start (copy/paste)

```bash
# 1) In repo root
pip install msal openpyxl

# 2) Edit local_inputs.json (or create your own path)
# 3) Run
python run_local.py
```

For PAD, use same interpreter and absolute paths as above.
