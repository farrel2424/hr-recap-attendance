# database/db.py
import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "absensi.db")

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn

def init_db():
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
            tanggal             TEXT NOT NULL,
            shift               TEXT,
            tipe_shift          TEXT,
            jam_masuk           TEXT,
            jam_keluar          TEXT,
            jam_kerja           REAL DEFAULT 0,
            status_absensi      TEXT,
            status_klasifikasi  TEXT,
            leave_app           TEXT,
            periode             TEXT NOT NULL,
            UNIQUE(karyawan_id, tanggal)
        );

        CREATE INDEX IF NOT EXISTS idx_periode
            ON absensi_harian(periode);

        CREATE INDEX IF NOT EXISTS idx_karyawan_tanggal
            ON absensi_harian(karyawan_id, tanggal);
        """)

        # Migrasi: tambah kolom leave_app jika belum ada (untuk DB lama)
        try:
            conn.execute("ALTER TABLE absensi_harian ADD COLUMN leave_app TEXT")
        except Exception:
            pass  # Kolom sudah ada, abaikan


def save_periode(df_raw, periode: str):
    with get_conn() as conn:
        conn.execute("DELETE FROM absensi_harian WHERE periode = ?", (periode,))

        for _, r in df_raw.iterrows():
            account = str(r.get("Account", "")).strip()
            nama    = str(r.get("Name", "")).strip()
            rules   = str(r.get("Rules", "")).strip()
            dept    = str(r.get("Department", "")).strip()

            if not account or account in ("", "--"):
                continue

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

            import re
            raw_time = str(r.get("Time", ""))
            m = re.search(r'(\d{4})/(\d{2})/(\d{2})', raw_time)
            tanggal = f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else None
            if not tanggal:
                continue

            leave_val = str(r.get("Leave & Overtime Application", "") or "").strip()

            conn.execute("""
                INSERT OR REPLACE INTO absensi_harian
                    (karyawan_id, tanggal, shift, tipe_shift, jam_masuk, jam_keluar,
                     jam_kerja, status_absensi, status_klasifikasi, leave_app, periode)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                karyawan_id,
                tanggal,
                str(r.get("Shift", "")).strip(),
                r.get("_tipe_shift"),
                str(r.get("Earliest", "")).strip(),
                str(r.get("Latest", "")).strip(),
                (lambda v: float(v) if str(v).replace('.','',1).lstrip('-').isdigit() else 0.0)(r.get("Actual working hours(Hour)", 0) or 0),
                str(r.get("Attendance results", "")).strip(),
                r.get("_status_klasifikasi"),
                leave_val,
                periode,
            ))


def get_periodes():
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT DISTINCT periode FROM absensi_harian ORDER BY periode DESC
        """).fetchall()
    return [r["periode"] for r in rows]


def get_rekap(periode: str):
    """
    Rekap per karyawan untuk satu periode.
    status_klasifikasi menggunakan format baru (separator '|'):
      S, Late, 1/2 UL, UL, 1/2 AL, AL, WFA, DW, K, Off
    """
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT
                k.nama, k.account, k.rules,
                SUM(CASE WHEN a.status_klasifikasi = 'S'
                         THEN 1 ELSE 0 END) AS normal,
                SUM(CASE WHEN a.status_klasifikasi = 'Late'
                         THEN 1 ELSE 0 END) AS late,
                SUM(CASE WHEN a.status_klasifikasi = '1/2 UL'
                         THEN 1 ELSE 0 END) AS half_ul,
                SUM(CASE WHEN a.status_klasifikasi = 'UL'
                         THEN 1 ELSE 0 END) AS ul_count,
                SUM(CASE WHEN a.status_klasifikasi = '1/2 AL'
                         THEN 1 ELSE 0 END) AS half_al,
                SUM(CASE WHEN a.status_klasifikasi = 'AL'
                         THEN 1 ELSE 0 END) AS al,
                SUM(CASE WHEN a.status_klasifikasi = 'WFA'
                         THEN 1 ELSE 0 END) AS wfa,
                SUM(CASE WHEN a.status_klasifikasi = 'DW'
                         THEN 1 ELSE 0 END) AS dw,
                SUM(CASE WHEN a.status_klasifikasi = 'K'
                         THEN 1 ELSE 0 END) AS k_sick,
                SUM(CASE WHEN a.status_klasifikasi = 'Off'
                         THEN 1 ELSE 0 END) AS off_count
            FROM karyawan k
            JOIN absensi_harian a ON a.karyawan_id = k.id
            WHERE a.periode = ?
            GROUP BY k.id
            ORDER BY k.rules, k.nama
        """, (periode,)).fetchall()
    import pandas as pd
    return pd.DataFrame([dict(r) for r in rows])


def get_daily(account: str, periode: str):
    """Ambil semua data harian dari DB untuk satu karyawan + periode."""
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT
                a.tanggal, a.shift, a.tipe_shift,
                a.jam_masuk, a.jam_keluar, a.jam_kerja,
                a.status_absensi, a.status_klasifikasi, a.leave_app
            FROM absensi_harian a
            JOIN karyawan k ON k.id = a.karyawan_id
            WHERE k.account = ? AND a.periode = ?
            ORDER BY a.tanggal
        """, (account, periode)).fetchall()
    import pandas as pd
    return pd.DataFrame([dict(r) for r in rows])