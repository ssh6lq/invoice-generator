"""
feedback_store.py
사용자 문의/이슈를 SQLite(로컬 파일 하나, feedback.db)에 저장한다.
별도 DB 서버 없이 이 프로젝트 규모에 맞춘 가장 가벼운 저장 방식.
"""

import os
import sqlite3
from datetime import datetime

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "feedback.db")


def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                screenshot BLOB,
                screenshot_mime TEXT,
                status TEXT NOT NULL DEFAULT 'open',
                created_at TEXT NOT NULL
            )
        """)


def add_feedback(title, content, screenshot_bytes=None, screenshot_mime=None):
    with _conn() as conn:
        cur = conn.execute(
            "INSERT INTO feedback (title, content, screenshot, screenshot_mime, status, created_at) "
            "VALUES (?, ?, ?, ?, 'open', ?)",
            (title, content, screenshot_bytes, screenshot_mime,
             datetime.now().isoformat(timespec="seconds")),
        )
        return cur.lastrowid


def list_feedback():
    """최신순. 스크린샷 본문(BLOB)은 목록에서 제외하고 첨부 여부만 내려준다."""
    with _conn() as conn:
        rows = conn.execute(
            "SELECT id, title, content, status, created_at, "
            "(screenshot IS NOT NULL) AS has_screenshot "
            "FROM feedback ORDER BY id DESC"
        ).fetchall()
        out = [dict(r) for r in rows]
        for r in out:
            r["has_screenshot"] = bool(r["has_screenshot"])
        return out


def get_screenshot(fid):
    with _conn() as conn:
        row = conn.execute(
            "SELECT screenshot, screenshot_mime FROM feedback WHERE id = ?", (fid,)
        ).fetchone()
        if not row or row["screenshot"] is None:
            return None, None
        return row["screenshot"], row["screenshot_mime"]


def set_status(fid, status):
    with _conn() as conn:
        conn.execute("UPDATE feedback SET status = ? WHERE id = ?", (status, fid))


def delete_feedback(fid):
    with _conn() as conn:
        conn.execute("DELETE FROM feedback WHERE id = ?", (fid,))
