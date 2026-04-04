#!/usr/bin/env python3
"""Cross-shell local test sequence for Windows paths (avoids PowerShell invocation issues)."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

REPO_PATH = Path(r"A:\Users\kyleg\PycharmProjects\eQuest-Simulation-Extraction")
SIM_PATH = REPO_PATH / r"sample_data\St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM"
WORKBOOK_PATH = REPO_PATH / r"output_files\Building Performance Assumptions.xlsm"
MASTER_ROOM_OUT = REPO_PATH / r"output_files\Building Performance Assumptions.master_room.updated.xlsm"
ECM_OUT = REPO_PATH / r"output_files\Building Performance Assumptions.combined.updated.xlsm"


def run_checked(args: list[str]) -> None:
    result = subprocess.run(args, cwd=REPO_PATH, text=True, capture_output=True)
    if result.stdout:
        print(result.stdout)
    if result.returncode != 0:
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        raise SystemExit(result.returncode)


def main() -> None:
    if not SIM_PATH.exists():
        raise FileNotFoundError(f"SIM file not found: {SIM_PATH}")
    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(f"Workbook file not found: {WORKBOOK_PATH}")
    run_checked([sys.executable, "--version"])
    run_checked([sys.executable, "equest_extractor.py", "--help"])
    sim_text = SIM_PATH.read_text(errors="ignore")
    sim_has_beps = "REPORT- BEPS" in sim_text.upper()
    if sim_has_beps:
        run_checked([sys.executable, "equest_extractor.py", str(SIM_PATH), "--report", "beps"])
    else:
        print("WARNING: BEPS section not found; skipping BEPS and ECM tests.")
    run_checked(
        [
            sys.executable,
            "equest_extractor.py",
            str(SIM_PATH),
            "--populate-master-room-list",
            str(WORKBOOK_PATH),
            "--model-run-type",
            "Baseline",
            "--output-workbook",
            str(MASTER_ROOM_OUT),
        ]
    )
    if sim_has_beps:
        run_checked(
            [
                sys.executable,
                "equest_extractor.py",
                str(SIM_PATH),
                "--update-ecm-data",
                str(MASTER_ROOM_OUT),
                "--model-run-type",
                "ECM-1",
                "--output-workbook",
                str(ECM_OUT),
            ]
        )
    print("Done. Output files:")
    print(MASTER_ROOM_OUT)
    if sim_has_beps:
        print(ECM_OUT)


if __name__ == "__main__":
    main()
