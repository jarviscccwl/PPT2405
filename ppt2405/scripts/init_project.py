"""Create a classified ppt2405 project without modifying source files."""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("project", type=Path, help="New or empty project directory")
    parser.add_argument("--source", action="append", default=[], type=Path, help="Input file to copy; repeatable")
    args = parser.parse_args()
    destination = args.project.expanduser().resolve()
    sources = [p.expanduser().resolve(strict=True) for p in args.source]
    if any(not p.is_file() for p in sources):
        parser.error("Every --source must be a file")
    if destination.exists() and (not destination.is_dir() or any(destination.iterdir())):
        parser.error("Project must be new or empty; existing work is never overwritten")
    destination.mkdir(parents=True, exist_ok=True)
    directories = ["sources", "analysis", "assets/paper-figures", "assets/paper-tables", "assets/data", "assets/external-images", "assets/external-media", "assets/ai-concepts", "assets/brand", "build", "output"]
    for directory in directories:
        (destination / directory).mkdir(parents=True, exist_ok=True)
    entries = []
    for number, source in enumerate(sources, 1):
        target = destination / "sources" / f"S{number:03d}_{source.name}"
        shutil.copy2(source, target)
        entries.append({"id": f"S{number:03d}", "title": source.stem, "original_path": str(source), "local_path": target.relative_to(destination).as_posix(), "sha256": hashlib.sha256(target.read_bytes()).hexdigest()})
    write_json(destination / "sources/manifest.json", {"schema_version": 1, "sources": entries})
    write_json(destination / "assets/manifest.json", {"schema_version": 1, "assets": []})
    write_json(destination / "analysis/evidence.json", {"schema_version": 1, "evidence": []})
    write_json(destination / "analysis/storyboard.json", {"schema_version": 1, "slides": []})
    write_json(destination / "project.json", {"skill": "ppt2405", "created_at": datetime.now(timezone.utc).isoformat(), "status": "intake", "assumptions": []})
    print(json.dumps({"project": str(destination), "sources": len(entries)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
