"""
submission_store.py
사용자가 '제출용 청구서(.xlsx)'를 다운로드한 시점(=제출)을 SQLite(submissions.db)에 기록한다.
경영지원팀이 관리자 페이지에서 접수 현황·처리 상태를 확인/변경하는 용도.
"""

import json
import os
import sqlite3
from datetime import datetime

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "submissions.db")

STATUSES = ("received", "reviewing", "approved", "rejected", "paid")


def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS submissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dept TEXT,
                name TEXT,
                title TEXT,
                count INTEGER NOT NULL,
                total_claim INTEGER NOT NULL,
                filename TEXT,
                status TEXT NOT NULL DEFAULT 'received',
                note TEXT,
                created_at TEXT NOT NULL,
                rows_json TEXT
            )
        """)
        cols = {r[1] for r in conn.execute("PRAGMA table_info(submissions)").fetchall()}
        if "rows_json" not in cols:
            conn.execute("ALTER TABLE submissions ADD COLUMN rows_json TEXT")


def add_submission(dept, name, title, count, total_claim, filename, rows=None):
    """rows: 매크로 검토(정렬·매핑) 완료된 행 목록(list[dict]) — 제출 시점의 최종 표를 그대로 보관."""
    rows_json = json.dumps(rows, ensure_ascii=False) if rows is not None else None
    with _conn() as conn:
        cur = conn.execute(
            "INSERT INTO submissions "
            "(dept, name, title, count, total_claim, filename, status, created_at, rows_json) "
            "VALUES (?, ?, ?, ?, ?, ?, 'received', ?, ?)",
            (dept, name, title, count, total_claim, filename,
             datetime.now().isoformat(timespec="seconds"), rows_json),
        )
        return cur.lastrowid


def list_submissions():
    """목록용 요약 — rows_json은 무거우므로 제외하고 보유 여부만 내려준다."""
    with _conn() as conn:
        rows = conn.execute(
            "SELECT id, dept, name, title, count, total_claim, filename, "
            "status, note, created_at, (rows_json IS NOT NULL) AS has_rows "
            "FROM submissions ORDER BY id DESC"
        ).fetchall()
        return [dict(r) for r in rows]


def get_rows(sid):
    """단건의 매크로 검토 최종 표(행 목록)를 반환. 없으면 None."""
    with _conn() as conn:
        row = conn.execute("SELECT rows_json FROM submissions WHERE id = ?", (sid,)).fetchone()
        if row is None or row["rows_json"] is None:
            return None
        return json.loads(row["rows_json"])


def set_status(sid, status, note=None):
    if status not in STATUSES:
        raise ValueError("잘못된 상태값입니다.")
    with _conn() as conn:
        if note is None:
            conn.execute("UPDATE submissions SET status = ? WHERE id = ?", (status, sid))
        else:
            conn.execute("UPDATE submissions SET status = ?, note = ? WHERE id = ?",
                         (status, note, sid))


def delete_submission(sid):
    with _conn() as conn:
        conn.execute("DELETE FROM submissions WHERE id = ?", (sid,))
