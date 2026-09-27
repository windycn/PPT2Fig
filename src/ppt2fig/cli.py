from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .core import ExportError, ExportOptions, check_poppler, export


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="ppt2fig")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("doctor", help="Check local Poppler dependencies")
    convert = sub.add_parser("convert", help="Convert a PowerPoint-generated PDF")
    convert.add_argument("job", type=Path, help="JSON job file")
    dialog = sub.add_parser("dialog", help="Open the local export settings window")
    dialog.add_argument("request", type=Path)
    dialog.add_argument("response", type=Path)
    args = parser.parse_args(argv)
    try:
        if args.command == "doctor":
            print(json.dumps({"status": "ok", "poppler": check_poppler()}, ensure_ascii=False))
        elif args.command == "convert":
            outputs = export(ExportOptions.from_file(args.job))
            print(json.dumps({"status": "ok", "files": [str(p) for p in outputs]}, ensure_ascii=False))
        else:
            from .gui import show_dialog
            show_dialog(args.request, args.response)
        return 0
    except (ExportError, ValueError, KeyError, OSError) as exc:
        print(json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
