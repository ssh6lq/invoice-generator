"""
submission_store.py
사용자가 '제출용 청구서(.xlsx)'를 다운로드한 시점(=제출)을 SQLite(submissions.db)에 기록한다.
경영지원팀이 관리자 페이지에서 접수 현황·처리 상태를 확인/변경하는 용도.
"""

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
                created_at TEXT NOT NULL
            )
        """)


def add_submission(dept, name, title, count, total_claim, filename):
    with _conn() as conn:
        cur = conn.execute(
            "INSERT INTO submissions "
            "(dept, name, title, count, total_claim, filename, status, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, 'received', ?)",
            (dept, name, title, count, total_claim, filename,
             datetime.now().isoformat(timespec="seconds")),
        )
        return cur.lastrowid


def list_submissions():
    with _conn() as conn:
        rows = conn.execute("SELECT * FROM submissions ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]


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
