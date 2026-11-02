from __future__ import annotations

import argparse
import json
from pathlib import Path, PurePosixPath
import re
import zipfile


def safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip().rstrip(".")
    return value or "unnamed"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("archive", type=Path)
    parser.add_argument("destination", type=Path)
    parser.add_argument("--prefix", action="append", default=[])
    parser.add_argument("--extensions", default=".xlsx,.xlsm,.xls,.proc,.json")
    args = parser.parse_args()
    allowed = {item.strip().lower() for item in args.extensions.split(",") if item.strip()}
    args.destination.mkdir(parents=True, exist_ok=True)
    manifest: list[dict[str, object]] = []
    with zipfile.ZipFile(args.archive) as archive:
        for number, info in enumerate(archive.infolist(), 1):
            if info.is_dir():
                continue
            original = info.filename
            path = PurePosixPath(original)
            if args.prefix and not any(original.startswith(prefix) for prefix in args.prefix):
                continue
            if path.suffix.lower() not in allowed:
                continue
            target = args.destination / f"{number:05d}_{safe_name(path.name)}"
            with archive.open(info) as source, target.open("wb") as output:
                while chunk := source.read(1024 * 1024):
                    output.write(chunk)
            manifest.append({"number": number, "original": original, "target": target.name, "size": info.file_size})
    (args.destination / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"extracted={len(manifest)} destination={args.destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
