#!/usr/bin/env python3
"""ย้ายข้อมูลจาก SQLite เดิมไป PostgreSQL โดยไม่ลบข้อมูลปลายทาง

ใช้งาน:
  export DATABASE_URL='postgresql://...'
  python3 migrate_sqlite_to_postgres.py

กำหนดไฟล์ SQLite อื่นได้ด้วย:
  SQLITE_DB_PATH=/path/to/tournament_events.db python3 migrate_sqlite_to_postgres.py
"""
from __future__ import annotations

import os
import sqlite3
from pathlib import Path

import psycopg
from psycopg import sql

BASE_DIR = Path(__file__).resolve().parent
SQLITE_DB_PATH = Path(os.environ.get("SQLITE_DB_PATH", BASE_DIR / "tournament_events.db"))
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = "postgresql://" + DATABASE_URL[len("postgres://"):]
if DATABASE_URL.startswith("postgresql+psycopg://"):
    DATABASE_URL = "postgresql://" + DATABASE_URL[len("postgresql+psycopg://"):]

TABLES = [
    "users",
    "tournaments",
    "events",
    "registrations",
    "registration_members",
    "certificates",
    "import_logs",
]


def sqlite_columns(conn: sqlite3.Connection, table: str) -> list[str]:
    return [row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()]


def postgres_columns(conn: psycopg.Connection, table: str) -> list[str]:
    rows = conn.execute(
        """SELECT column_name
           FROM information_schema.columns
           WHERE table_schema='public' AND table_name=%s
           ORDER BY ordinal_position""",
        (table,),
    ).fetchall()
    return [row[0] for row in rows]


def copy_table(source: sqlite3.Connection, target: psycopg.Connection, table: str) -> int:
    source_cols = sqlite_columns(source, table)
    target_cols = postgres_columns(target, table)
    cols = [c for c in source_cols if c in target_cols]
    if not cols:
        return 0

    rows = source.execute(
        f"SELECT {','.join(cols)} FROM {table} ORDER BY id"
    ).fetchall()
    if not rows:
        return 0

    query = sql.SQL("INSERT INTO {} ({}) VALUES ({}) ON CONFLICT DO NOTHING").format(
        sql.Identifier(table),
        sql.SQL(", ").join(map(sql.Identifier, cols)),
        sql.SQL(", ").join(sql.Placeholder() for _ in cols),
    )
    inserted = 0
    with target.cursor() as cur:
        for row in rows:
            cur.execute(query, tuple(row[c] for c in cols))
            inserted += max(cur.rowcount, 0)
    return inserted


def reset_sequence(conn: psycopg.Connection, table: str) -> None:
    conn.execute(
        """SELECT setval(
               pg_get_serial_sequence(%s, 'id'),
               COALESCE((SELECT MAX(id) FROM {}), 1),
               EXISTS(SELECT 1 FROM {})
           )""".format(table, table),
        (table,),
    )


def main() -> None:
    if not DATABASE_URL.startswith("postgresql://"):
        raise SystemExit("กรุณาตั้งค่า DATABASE_URL ของ PostgreSQL ก่อน")
    if not SQLITE_DB_PATH.exists():
        raise SystemExit(f"ไม่พบไฟล์ SQLite: {SQLITE_DB_PATH}")

    source = sqlite3.connect(SQLITE_DB_PATH)
    source.row_factory = sqlite3.Row
    with psycopg.connect(DATABASE_URL) as target:
        total = 0
        for table in TABLES:
            count = copy_table(source, target, table)
            total += count
            print(f"{table}: ย้ายเพิ่ม {count} รายการ")
        for table in TABLES:
            reset_sequence(target, table)
        target.commit()
    source.close()
    print(f"เสร็จแล้ว: ย้ายเพิ่มรวม {total} รายการ")
    print("หมายเหตุ: ไฟล์รูปภาพใน static/uploads ต้องคัดลอกหรือใช้ Volume แยกต่างหาก")


if __name__ == "__main__":
    main()
