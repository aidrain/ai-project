from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import sqlite3
from typing import Any

import pandas as pd

DB_PATH = Path(__file__).parent / "curriculum.db"


def _get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """SQLite 데이터베이스와 작업 저장 테이블을 초기화합니다."""
    with _get_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS workspaces (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL UNIQUE,
                college TEXT NOT NULL,
                major TEXT NOT NULL,
                df_2026_json TEXT NOT NULL,
                changes_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def save_workspace(
    filename: str,
    meta: dict[str, str],
    df_2026: pd.DataFrame,
    changes: list[dict[str, Any]],
) -> None:
    """현재 학과의 26년 원본 데이터와 개정 내역을 SQLite에 저장(또는 갱신)합니다."""
    if not filename or df_2026 is None:
        return
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df_json = json.dumps(df_2026.to_dict(orient="records"), ensure_ascii=False)
    changes_json = json.dumps(changes, ensure_ascii=False)

    with _get_conn() as conn:
        conn.execute(
            """
            INSERT INTO workspaces (filename, college, major, df_2026_json, changes_json, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(filename) DO UPDATE SET
                college = excluded.college,
                major = excluded.major,
                df_2026_json = excluded.df_2026_json,
                changes_json = excluded.changes_json,
                updated_at = excluded.updated_at
            """,
            (
                filename,
                meta.get("college", ""),
                meta.get("major", ""),
                df_json,
                changes_json,
                now_str,
            ),
        )
        conn.commit()


def list_workspaces() -> list[dict[str, Any]]:
    """저장된 학과 작업 목록을 최신순으로 반환합니다."""
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT id, filename, college, major, changes_json, updated_at FROM workspaces ORDER BY updated_at DESC, id DESC"
        ).fetchall()

    result = []
    for row in rows:
        changes = json.loads(row["changes_json"])
        result.append(
            {
                "id": row["id"],
                "filename": row["filename"],
                "college": row["college"],
                "major": row["major"],
                "change_count": len(changes),
                "updated_at": row["updated_at"],
            }
        )
    return result


def _row_to_workspace(row: sqlite3.Row) -> dict[str, Any]:
    records = json.loads(row["df_2026_json"])
    return {
        "filename": row["filename"],
        "meta": {"college": row["college"], "major": row["major"]},
        "df_2026": pd.DataFrame(records),
        "changes": json.loads(row["changes_json"]),
    }


def load_workspace(workspace_id: int) -> dict[str, Any] | None:
    """선택한 ID의 학과 작업을 불러옵니다."""
    with _get_conn() as conn:
        row = conn.execute("SELECT * FROM workspaces WHERE id = ?", (workspace_id,)).fetchone()
    return _row_to_workspace(row) if row is not None else None


def load_latest_workspace() -> dict[str, Any] | None:
    """가장 최근에 작업한 학과 데이터를 불러옵니다."""
    with _get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM workspaces ORDER BY updated_at DESC, id DESC LIMIT 1"
        ).fetchone()
    return _row_to_workspace(row) if row is not None else None


def delete_workspace(workspace_id: int) -> None:
    """저장된 학과 작업을 삭제합니다."""
    with _get_conn() as conn:
        conn.execute("DELETE FROM workspaces WHERE id = ?", (workspace_id,))
        conn.commit()
