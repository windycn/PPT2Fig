"""Build the local helper on the current platform with PyInstaller."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    command = [
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--onefile",
        "--name", "ppt2fig", "--paths", str(ROOT / "src"),
        "--distpath", str(ROOT / "dist/helper"),
        "--workpath", str(ROOT / "dist/pyinstaller-work"),
        "--specpath", str(ROOT / "dist"),
        str(ROOT / "scripts/ppt2fig_entry.py"),
    ]
    subprocess.run(command, check=True, cwd=ROOT)
    binary = ROOT / "dist/helper" / ("ppt2fig.exe" if sys.platform == "win32" else "ppt2fig")
    print(binary)


if __name__ == "__main__":
    main()
