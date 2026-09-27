from __future__ import annotations

import configparser
from pathlib import Path


def read_ini(path: Path) -> configparser.ConfigParser:
    parser = configparser.ConfigParser(interpolation=None)
    with path.open("rb") as stream:
        bom = stream.read(2)
    encoding = "utf-16" if bom in (b"\xff\xfe", b"\xfe\xff") else "utf-8"
    parser.read(path, encoding=encoding)
    return parser


def write_ini_utf16(path: Path, section: str, values: dict[str, object]) -> None:
    lines = [f"[{section}]"]
    for key, value in values.items():
        if isinstance(value, list):
            value = ",".join(map(str, value))
        elif value is None:
            value = ""
        lines.append(f"{key}={value}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-16")
