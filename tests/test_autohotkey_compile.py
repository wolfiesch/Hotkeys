"""Test that validates the AutoHotkey v2 script compiles without syntax errors."""

from __future__ import annotations

import os
import shlex
import subprocess
import warnings
from pathlib import Path
from typing import List, Optional

import pytest

from tests.autohotkey_static_validator import validate_autohotkey_source

# The repository ships a single primary AutoHotkey script that drives all
# keyboard automation. The path resolution is intentionally verbose so that it
# remains resilient to the tests being invoked from any working directory.
SCRIPT_PATH: Path = Path(__file__).resolve().parents[1] / "ExcelDatabookLayers.ahk"

# Multiple environment variables are supported so that developers can reuse
# existing local setup without renaming their configuration.
AUTOHOTKEY_ENV_VARS: tuple[str, ...] = (
    "AUTOHOTKEY_EXECUTABLE",
    "AUTOHOTKEY_PATH",
    "AHK_EXECUTABLE",
)


def _resolve_autohotkey_command() -> Optional[List[str]]:
    """Return the command list used to invoke AutoHotkey if configured.

    The helper inspects several environment variables to maximize compatibility
    with developer machines and CI providers. When none of the expected
    variables are present we gracefully return ``None`` so the caller can skip
    the test with a clear message.
    """

    for variable_name in AUTOHOTKEY_ENV_VARS:
        configured_value = os.environ.get(variable_name)
        if configured_value:
            # ``shlex.split`` is used instead of ``str.split`` so that quoted
            # paths containing whitespace are handled correctly across
            # platforms. The call produces a concrete ``list[str]`` which
            # ``subprocess.run`` expects.
            return shlex.split(configured_value)
    return None


def test_excel_databook_layers_compiles_cleanly() -> None:
    """Ensure that the ExcelDatabookLayers AutoHotkey script parses cleanly.

    AutoHotkey parses an entire script before executing its auto-execute
    section, so invoking the interpreter is a reliable way to catch syntax
    regressions. The environment variable ``AHK_VALIDATE_ONLY`` tells the
    script to exit immediately after load, preventing the hotkeys from
    remaining active during the test run.
    """

    autohotkey_command = _resolve_autohotkey_command()
    if autohotkey_command is None:
        warnings.warn(
            (
                "AutoHotkey executable not provided. Using static syntax validation instead."
                " Set AUTOHOTKEY_EXECUTABLE, AUTOHOTKEY_PATH, or AHK_EXECUTABLE to exercise"
                " the real interpreter during tests."
            ),
            stacklevel=1,
        )
        validate_autohotkey_source(
            script_path=SCRIPT_PATH,
            source_text=SCRIPT_PATH.read_text(encoding="utf-8"),
        )
        return

    if not SCRIPT_PATH.is_file():
        pytest.fail(f"Expected AutoHotkey script at {SCRIPT_PATH} was not found.")

    environment = os.environ.copy()
    environment["AHK_VALIDATE_ONLY"] = "1"

    command: List[str] = [
        *autohotkey_command,
        "/ErrorStdOut",
        str(SCRIPT_PATH),
    ]

    completed = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        env=environment,
    )

    diagnostic_context = (
        "\n--- AutoHotkey STDOUT ---\n"
        f"{completed.stdout}\n"
        "--- AutoHotkey STDERR ---\n"
        f"{completed.stderr}\n"
        "--- Command ---\n"
        f"{' '.join(command)}\n"
    )

    assert (
        completed.returncode == 0
    ), f"AutoHotkey reported a compilation failure. {diagnostic_context}"
