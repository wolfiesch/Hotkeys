"""Fallback AutoHotkey syntax validation utilities for test environments.

The production workflow for this repository expects AutoHotkey v2 to be
available so the full script can be parsed by the official interpreter.
However, the execution environment for automated tests in this repository is
Linux-based and does not ship with AutoHotkey. To keep ``pytest`` meaningful we
perform a structural validation pass that focuses on failure modes the authors
have historically encountered:

* Unterminated strings, block comments, or block delimiters.
* Premature closing delimiters that leave the surrounding structure malformed.

The goal of this module is not to re-implement the AutoHotkey parser. Instead
it offers a conservative safety net that complements code review by ensuring
that edits do not introduce obvious syntax errors. When an actual AutoHotkey
binary is present the main test bypasses these helpers entirely and relies on
``AutoHotkey.exe`` for definitive validation.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

__all__ = ["validate_autohotkey_source"]


@dataclass(slots=True)
class _ValidationError(Exception):
    """Represent a structural issue detected during static validation."""

    message: str
    line: int
    column: int

    def __str__(self) -> str:
        return self.message


def _validate_balanced_delimiters(source_text: str) -> None:
    """Ensure that block delimiters, strings, and comments are well-balanced."""

    opening_pairs: dict[str, str] = {"(": ")", "[": "]", "{": "}"}
    closing_to_opening: dict[str, str] = {
        value: key for key, value in opening_pairs.items()
    }
    stack: list[tuple[str, int, int]] = []

    line_number: int = 1
    column_number: int = 1

    index: int = 0
    length: int = len(source_text)
    state: str = "code"

    while index < length:
        character = source_text[index]

        if state == "string":
            if character == '"':
                if index + 1 < length and source_text[index + 1] == '"':
                    index += 2
                    column_number += 2
                    continue
                state = "code"
                index += 1
                column_number += 1
                continue
            if character == "\n":
                line_number += 1
                column_number = 1
            elif character == "\r":
                column_number = 1
            else:
                column_number += 1
            index += 1
            continue

        if state == "block_comment":
            if (
                character == "*"
                and index + 1 < length
                and source_text[index + 1] == "/"
            ):
                state = "code"
                index += 2
                column_number += 2
                continue
            if character == "\n":
                line_number += 1
                column_number = 1
            elif character == "\r":
                column_number = 1
            else:
                column_number += 1
            index += 1
            continue

        if character == '"':
            state = "string"
            index += 1
            column_number += 1
            continue

        if character == "/" and index + 1 < length and source_text[index + 1] == "*":
            state = "block_comment"
            index += 2
            column_number += 2
            continue

        if character == ";":
            while index < length and source_text[index] not in {"\n", "\r"}:
                index += 1
                column_number += 1
            continue

        if character in opening_pairs:
            stack.append((character, line_number, column_number))
        elif character in closing_to_opening:
            if not stack:
                raise _ValidationError(
                    message=(
                        f"Encountered closing delimiter '{character}' without a matching opening delimiter."
                    ),
                    line=line_number,
                    column=column_number,
                )
            opening_character, opening_line, opening_column = stack.pop()
            if opening_pairs[opening_character] != character:
                expected = opening_pairs[opening_character]
                raise _ValidationError(
                    message=(
                        "Mismatched delimiters: expected "
                        f"'{expected}' to close '{opening_character}' opened at line {opening_line}, "
                        f"column {opening_column}, but found '{character}'."
                    ),
                    line=line_number,
                    column=column_number,
                )

        if character == "\n":
            line_number += 1
            column_number = 1
        elif character == "\r":
            column_number = 1
        else:
            column_number += 1
        index += 1

    if state == "string":
        raise _ValidationError(
            message="Unterminated string literal detected.",
            line=line_number,
            column=column_number,
        )
    if state == "block_comment":
        raise _ValidationError(
            message="Unterminated block comment detected.",
            line=line_number,
            column=column_number,
        )
    if stack:
        opening_character, opening_line, opening_column = stack.pop()
        raise _ValidationError(
            message=f"Unclosed delimiter '{opening_character}' detected.",
            line=opening_line,
            column=opening_column,
        )


def validate_autohotkey_source(*, script_path: Path, source_text: str) -> None:
    """Validate AutoHotkey source structure when the interpreter is unavailable."""

    try:
        _validate_balanced_delimiters(source_text)
    except _ValidationError as error:
        raise AssertionError(
            (
                f"Static AutoHotkey validation failed for {script_path} at line {error.line}, "
                f"column {error.column}: {error.message}.\n"
                "Provide an AutoHotkey interpreter via the AUTOHOTKEY_EXECUTABLE, AUTOHOTKEY_PATH, or "
                "AHK_EXECUTABLE environment variables for definitive validation."
            )
        ) from error
