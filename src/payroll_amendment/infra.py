from __future__ import annotations

from datetime import datetime
import hashlib
import json
from pathlib import Path
import re
import shutil
import sqlite3
from zoneinfo import ZoneInfo


SEOUL = ZoneInfo("Asia/Seoul")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip().rstrip(".")
    return re.sub(r"\s+", " ", value) or "unnamed"


def make_run_id(company_id: str) -> str:
    now = datetime.now(SEOUL).strftime("%Y%m%dT%H%M%S")
    return f"{now}_{safe_name(company_id)}_{hashlib.sha256(now.encode()).hexdigest()[:8]}"


def snapshot_file(source: Path, destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    if sha256(source) != sha256(destination):
        destination.unlink(missing_ok=True)
        raise IOError(f"스냅샷 해시가 다릅니다: {source.name}")
    return destination


class RunStore:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(path)
        self.connection.row_factory = sqlite3.Row
        self.connection.executescript(
            """
            PRAGMA journal_mode=WAL;
            CREATE TABLE IF NOT EXISTS runs(
              run_id TEXT PRIMARY KEY, company_id TEXT NOT NULL, years_json TEXT NOT NULL,
              status TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS stages(
              run_id TEXT NOT NULL, year INTEGER, stage_code TEXT NOT NULL, status TEXT NOT NULL,
              result_json TEXT NOT NULL, updated_at TEXT NOT NULL,
              PRIMARY KEY(run_id, year, stage_code)
            );
            """
        )

    def create(self, run_id: str, company_id: str, years: list[int]) -> None:
        now = datetime.now(SEOUL).isoformat()
        self.connection.execute(
            "INSERT OR REPLACE INTO runs VALUES(?,?,?,?,?,?)",
            (run_id, company_id, json.dumps(years), "RUNNING", now, now),
        )
        self.connection.commit()

    def record_stage(self, run_id: str, year: int | None, stage_code: str, status: str, result: dict) -> None:
        now = datetime.now(SEOUL).isoformat()
        self.connection.execute(
            "INSERT OR REPLACE INTO stages VALUES(?,?,?,?,?,?)",
            (run_id, year, stage_code, status, json.dumps(result, ensure_ascii=False, default=str), now),
        )
        self.connection.commit()

    def finish(self, run_id: str, status: str) -> None:
        self.connection.execute(
            "UPDATE runs SET status=?,updated_at=? WHERE run_id=?",
            (status, datetime.now(SEOUL).isoformat(), run_id),
        )
        self.connection.commit()

    def status(self, run_id: str) -> dict:
        run = self.connection.execute("SELECT * FROM runs WHERE run_id=?", (run_id,)).fetchone()
        if run is None:
            raise KeyError(run_id)
        stages = self.connection.execute(
            "SELECT year,stage_code,status,result_json FROM stages WHERE run_id=? ORDER BY year,stage_code", (run_id,)
        ).fetchall()
        return {"run": dict(run), "stages": [dict(row) for row in stages]}

