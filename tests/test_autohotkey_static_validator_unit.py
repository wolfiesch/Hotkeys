"""Unit tests for the lightweight AutoHotkey static validator."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.autohotkey_static_validator import validate_autohotkey_source


def _write_script(directory: Path, name: str, contents: str) -> Path:
    """Create a temporary AutoHotkey script for validation-focused tests."""

    directory.mkdir(parents=True, exist_ok=True)
    script_path = directory / name
    script_path.write_text(contents, encoding="utf-8")
    return script_path


def test_validator_accepts_balanced_script(tmp_path: Path) -> None:
    """The validator should succeed on scripts with balanced delimiters."""

    script_path = _write_script(
        tmp_path / "ahk",
        "example.ahk",
        """
        msg := "Hello"
        if (msg != "") {
            MsgBox msg
        }
        """.strip(),
    )

    validate_autohotkey_source(
        script_path=script_path, source_text=script_path.read_text(encoding="utf-8")
    )


def test_validator_rejects_unclosed_brace(tmp_path: Path) -> None:
    """Ensure unclosed braces are detected with a helpful message."""

    script_path = _write_script(
        tmp_path / "ahk",
        "broken.ahk",
        """
        if (true) {
            MsgBox "Hello"
        """.strip(),
    )

    with pytest.raises(AssertionError) as error_info:
        validate_autohotkey_source(
            script_path=script_path, source_text=script_path.read_text(encoding="utf-8")
        )

    assert "Unclosed delimiter '{'" in str(error_info.value)


def test_validator_rejects_unterminated_string(tmp_path: Path) -> None:
    """Unterminated string literals must be surfaced as validation errors."""

    script_path = _write_script(tmp_path / "ahk", "string.ahk", 'MsgBox "Hello')

    with pytest.raises(AssertionError) as error_info:
        validate_autohotkey_source(
            script_path=script_path, source_text=script_path.read_text(encoding="utf-8")
        )

    assert "Unterminated string literal" in str(error_info.value)
