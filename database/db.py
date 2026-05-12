# database/db.py
import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "absensi.db")

def get_conn():
    """Koneksi dengan foreign key enforcement dan WAL mode untuk performa."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row   # hasil query bisa diakses by nama kolom
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")  # aman untuk concurrent read
    return conn

def init_db():
    """Buat tabel jika belum ada. Aman dipanggil berulang kali."""
    with get_conn() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS karyawan (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            account   TEXT UNIQUE NOT NULL,
            nama      TEXT NOT NULL,
            rules     TEXT,
            department TEXT
        );

        CREATE TABLE IF NOT EXISTS absensi_harian (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            karyawan_id         INTEGER NOT NULL REFERENCES karyawan(id),
            tanggal             TEXT NOT NULL,          -- format: YYYY-MM-DD
            shift               TEXT,
            tipe_shift          TEXT,                   -- S1 / S2 / H
            jam_masuk           TEXT,
            jam_keluar          TEXT,
            jam_kerja           REAL DEFAULT 0,
            status_absensi      TEXT,
            status_klasifikasi  TEXT,                   -- Normal / Late / K / NULL
            periode             TEXT NOT NULL,          -- format: YYYY-MM
            UNIQUE(karyawan_id, tanggal)
        );

        CREATE INDEX IF NOT EXISTS idx_periode
            ON absensi_harian(periode);

        CREATE INDEX IF NOT EXISTS idx_karyawan_tanggal
            ON absensi_harian(karyawan_id, tanggal);
        """)

def save_periode(df_raw, periode: str):
    """
    file_bytes : bytes mentah dari uploaded.read()
    df_result  : DataFrame hasil process_file (untuk ambil daftar Account valid)
    """
    with get_conn() as conn:
        # Hapus data periode ini jika sudah pernah diimport
        conn.execute("""
            DELETE FROM absensi_harian
            WHERE periode = ?
        """, (periode,))

        for _, r in df_raw.iterrows():
            account = str(r.get("Account", "")).strip()
            nama    = str(r.get("Name", "")).strip()
            rules   = str(r.get("Rules", "")).strip()
            dept    = str(r.get("Department", "")).strip()

            if not account or account in ("", "--"):
                continue

            # Upsert karyawan (bisa saja nama/rules berubah)
            conn.execute("""
                INSERT INTO karyawan(account, nama, rules, department)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(account) DO UPDATE SET
                    nama       = excluded.nama,
                    rules      = excluded.rules,
                    department = excluded.department
            """, (account, nama, rules, dept))

            karyawan_id = conn.execute(
                "SELECT id FROM karyawan WHERE account = ?", (account,)
            ).fetchone()["id"]

            # Parse tanggal dari "2026/04/30 星期四" → "2026-04-30"
            import re
            raw_time = str(r.get("Time", ""))
            m = re.search(r'(\d{4})/(\d{2})/(\d{2})', raw_time)
            tanggal = f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else None
            if not tanggal:
                continue

            conn.execute("""
                INSERT OR REPLACE INTO absensi_harian
                    (karyawan_id, tanggal, shift, tipe_shift, jam_masuk, jam_keluar,
                     jam_kerja, status_absensi, status_klasifikasi, periode)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                karyawan_id,
                tanggal,
                str(r.get("Shift", "")).strip(),
                r.get("_tipe_shift"),          # kolom hasil classify_shift_type
                str(r.get("Earliest", "")).strip(),
                str(r.get("Latest", "")).strip(),
                (lambda v: float(v) if str(v).replace('.','',1).lstrip('-').isdigit() else 0.0)(r.get("Actual working hours(Hour)", 0) or 0),
                str(r.get("Attendance results", "")).strip(),
                r.get("_status_klasifikasi"),  # kolom hasil classify()
                periode,
            ))


def get_periodes():
    """Kembalikan list periode yang sudah diimport, urut terbaru dulu."""
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT DISTINCT periode
            FROM absensi_harian
            ORDER BY periode DESC
        """).fetchall()
    return [r["periode"] for r in rows]


def get_rekap(periode: str):
    """Rekap Normal/Late/K per karyawan untuk satu periode."""
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT
                k.nama, k.account, k.rules,
                SUM(CASE WHEN a.status_klasifikasi = 'Normal' THEN 1 ELSE 0 END) AS normal,
                SUM(CASE WHEN a.status_klasifikasi = 'Late'   THEN 1 ELSE 0 END) AS late,
                SUM(CASE WHEN a.status_klasifikasi = 'K'      THEN 1 ELSE 0 END) AS k
            FROM karyawan k
            JOIN absensi_harian a ON a.karyawan_id = k.id
            WHERE a.periode = ?
            GROUP BY k.id
            ORDER BY k.rules, k.nama
        """, (periode,)).fetchall()
    import pandas as pd
    return pd.DataFrame([dict(r) for r in rows])


def get_daily(account: str, periode: str):
    """Rincian harian satu karyawan untuk satu periode."""
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT
                a.tanggal, a.shift, a.tipe_shift,
                a.jam_masuk, a.jam_keluar, a.jam_kerja,
                a.status_absensi, a.status_klasifikasi
            FROM absensi_harian a
            JOIN karyawan k ON k.id = a.karyawan_id
            WHERE k.account = ? AND a.periode = ?
            ORDER BY a.tanggal
        """, (account, periode)).fetchall()
    import pandas as pd
    return pd.DataFrame([dict(r) for r in rows])