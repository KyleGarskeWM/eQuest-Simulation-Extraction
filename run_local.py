#!/usr/bin/env python3
"""Local runner for equest_extractor.py using a JSON config file."""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SUPPORTED_MODEL_RUN_TYPES = {
    "Baseline",
    "Proposed",
    "ECM-1",
    "ECM-2",
    "ECM-3",
    "ECM-4",
    "ECM-5",
    "ECM-6",
    "ECM-7",
}

GRAPH_CONFIG_TO_ENV = {
    "client_id": "GRAPH_CLIENT_ID",
    "tenant_id": "GRAPH_TENANT_ID",
    "client_secret": "GRAPH_CLIENT_SECRET",
    "user_id": "GRAPH_USER_ID",
}


def resolve_graph_config_path(config: dict, config_path: Path) -> str | None:
    graph_config_path = config.get("graph_config_path")
    if graph_config_path in (None, ""):
        return None
    path_value = Path(str(graph_config_path))
    if not path_value.is_absolute():
        path_value = (config_path.parent / path_value).resolve()
    return str(path_value)


def resolve_model_run_type(config: dict, default: str) -> str:
    model_run_type = config.get("model_run_type", default)
    if model_run_type not in SUPPORTED_MODEL_RUN_TYPES:
        supported = ", ".join(["Baseline", "Proposed", "ECM-1", "ECM-2", "ECM-3", "ECM-4", "ECM-5", "ECM-6", "ECM-7"])
        raise ValueError(f"Unsupported model_run_type: {model_run_type}. Supported options: {supported}")
    return model_run_type


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
                resolve_model_run_type(config, "Baseline"),
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
                resolve_model_run_type(config, "ECM-1"),
                "--output-workbook",
                config["output_workbook_path"],
            ]
        )
    elif mode == "schedule_importer":
        command.extend(
            [
                "--populate-schedules",
                config["workbook_path"],
                "--output-workbook",
                config["output_workbook_path"],
            ]
        )
    elif mode == "combined":
        raise ValueError("Use build_combined_commands() for mode='combined'.")
    else:
        raise ValueError(f"Unsupported mode: {mode}")
    return command


def build_combined_commands(config: dict, master_output_path: str, ecm_output_path: str) -> list[list[str]]:
    """Build three-step command list: Master Room List -> ECM Data -> Schedule Importer."""
    combined_model_run_type = resolve_model_run_type(config, "Baseline")
    sim_file = config["sim_file"]
    workbook_path = config["workbook_path"]
    output_workbook_path = config["output_workbook_path"]
    return [
        [
            sys.executable,
            "equest_extractor.py",
            sim_file,
            "--populate-master-room-list",
            workbook_path,
            "--model-run-type",
            combined_model_run_type,
            "--output-workbook",
            master_output_path,
        ],
        [
            sys.executable,
            "equest_extractor.py",
            sim_file,
            "--update-ecm-data",
            master_output_path,
            "--model-run-type",
            combined_model_run_type,
            "--output-workbook",
            ecm_output_path,
        ],
        [
            sys.executable,
            "equest_extractor.py",
            sim_file,
            "--populate-schedules",
            ecm_output_path,
            "--output-workbook",
            output_workbook_path,
        ],
    ]


def build_process_env(config: dict, config_path: Path) -> dict[str, str]:
    """Build child process env, optionally injecting Graph auth settings from config."""
    process_env = os.environ.copy()
    resolved_graph_config_path = resolve_graph_config_path(config, config_path=config_path)
    if resolved_graph_config_path:
        process_env["GRAPH_CONFIG_PATH"] = resolved_graph_config_path
    graph_config = config.get("graph")
    if not graph_config:
        return process_env
    if not isinstance(graph_config, dict):
        raise ValueError("local_inputs.json 'graph' must be an object when provided.")
    for config_key, env_key in GRAPH_CONFIG_TO_ENV.items():
        value = graph_config.get(config_key)
        if value is not None:
            process_env[env_key] = str(value)
    return process_env


def run_command(command: list[str], process_env: dict[str, str], cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(command, text=True, capture_output=True, cwd=cwd, env=process_env)


def main() -> None:
    config_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("local_inputs.json")
    if not config_path.exists():
        raise FileNotFoundError(
            f"Config file not found: {config_path}. Copy local_inputs.template.json to local_inputs.json and edit paths."
        )
    config = json.loads(config_path.read_text(encoding="utf-8"))
    process_env = build_process_env(config, config_path=config_path)
    script_dir = Path(__file__).resolve().parent
    mode = config.get("mode", "extract_report")
    if mode == "combined":
        with tempfile.TemporaryDirectory() as temp_dir:
            master_intermediate_output = str(Path(temp_dir) / "master_room_intermediate.xlsm")
            ecm_intermediate_output = str(Path(temp_dir) / "ecm_intermediate.xlsm")
            commands = build_combined_commands(
                config,
                master_output_path=master_intermediate_output,
                ecm_output_path=ecm_intermediate_output,
            )
            for command in commands:
                result = run_command(command, process_env=process_env, cwd=script_dir)
                if result.stdout:
                    print(result.stdout)
                if result.returncode != 0:
                    if result.stderr:
                        print(result.stderr, file=sys.stderr)
                    raise SystemExit(result.returncode)
    else:
        command = build_command(config)
        result = run_command(command, process_env=process_env, cwd=script_dir)
        if result.stdout:
            print(result.stdout)
        if result.returncode != 0:
            if result.stderr:
                print(result.stderr, file=sys.stderr)
            raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()
