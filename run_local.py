#!/usr/bin/env python3
"""Local runner for equest_extractor.py using a JSON config file."""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


def build_command(config: dict) -> list[str]:
    sim_file = config["sim_file"]
    command = [sys.executable, "equest_extractor.py", sim_file]
    mode = config.get("mode", "extract_report")
    if mode == "extract_report":
        command.extend(["--report", config.get("report", "all")])
    elif mode == "master_room_list":
        command.extend(
            [
                "--populate-master-room-list",
                config["workbook_path"],
                "--model-run-type",
                config.get("model_run_type", "Baseline"),
                "--output-workbook",
                config["output_workbook_path"],
            ]
        )
    elif mode == "ecm_data":
        command.extend(
            [
                "--update-ecm-data",
                config["workbook_path"],
                "--model-run-type",
                config.get("model_run_type", "ECM-1"),
                "--output-workbook",
                config["output_workbook_path"],
            ]
        )
    else:
        raise ValueError(f"Unsupported mode: {mode}")
    return command


def main() -> None:
    config_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("local_inputs.json")
    if not config_path.exists():
        raise FileNotFoundError(
            f"Config file not found: {config_path}. Copy local_inputs.template.json to local_inputs.json and edit paths."
        )
    config = json.loads(config_path.read_text(encoding="utf-8"))
    command = build_command(config)
    result = subprocess.run(command, text=True, capture_output=True, cwd=Path(__file__).resolve().parent)
    if result.stdout:
        print(result.stdout)
    if result.returncode != 0:
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()
