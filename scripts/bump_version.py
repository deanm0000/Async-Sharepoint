# /// script
# requires-python = ">=3.11"
# ///
"""Bump the patch version across Cargo.toml, core/Cargo.toml, and pyproject.toml."""

import os
import re
from pathlib import Path

import tomllib

ROOT_CARGO = Path("Cargo.toml")
CORE_CARGO = Path("core/Cargo.toml")
PYPROJECT = Path("pyproject.toml")
VERSION_RE = re.compile(r'^version = "(\d+)\.(\d+)\.(\d+)"', re.MULTILINE)


def _current_version() -> tuple[int, int, int]:
    data = tomllib.loads(PYPROJECT.read_text(encoding="utf-8"))
    major, minor, patch = (int(part) for part in data["project"]["version"].split("."))
    return major, minor, patch


def _bump_file(path: Path, new_version: str) -> None:
    text = path.read_text(encoding="utf-8")
    new_text, count = VERSION_RE.subn(f'version = "{new_version}"', text, count=1)
    if count != 1:
        raise SystemExit(f"Could not find a version field to bump in {path}")
    path.write_text(new_text, encoding="utf-8")


def main() -> None:
    major, minor, patch = _current_version()
    new_version = f"{major}.{minor}.{patch + 1}"
    for path in (ROOT_CARGO, CORE_CARGO, PYPROJECT):
        _bump_file(path, new_version)

    github_output = os.environ.get("GITHUB_OUTPUT")
    if github_output:
        with open(github_output, "a", encoding="utf-8") as fh:
            fh.write(f"version={new_version}\n")
            fh.write(f"tag=v{new_version}\n")
    print(new_version)


if __name__ == "__main__":
    main()
