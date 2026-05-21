# classifiers/base.py
"""
Konstanta dan helper parse yang dipakai bersama oleh semua classifier.
"""

import re
import pandas as pd
from datetime import time, datetime

# ──────────────────────────────────────────────────────────────
# Konstanta
# ──────────────────────────────────────────────────────────────

SKIP_SHIFTS     = {"Rest", "Not scheduled", "--", ""}
K_THRESHOLD_MIN = 120   # 2 jam — batas Late vs ½UL
_NOT_PUNCHED    = {"not punched", "--", ""}

# Nilai att_result yang langsung menghasilkan S (exact match)
S_ATT_RESULTS = {"Normal", "Normal（Correction of missed punch）"}


# ──────────────────────────────────────────────────────────────
# Helpers Parse
# ──────────────────────────────────────────────────────────────

def parse_shift_start(shift_text) -> int | None:
    """Ambil jam mulai shift dalam menit (misal 08:00 → 480). Return None jika tidak valid."""
    if not isinstance(shift_text, str):
        return None
    s = shift_text.strip()
    if s in SKIP_SHIFTS:
        return None
    m = re.search(r'(\d{1,2}):(\d{2})', s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def parse_shift_end(shift_text) -> int | None:
    """
    Ambil jam selesai shift dalam menit.
    Mendukung format 'WD：08:30-17:00' dan overnight 'Night：19:00-Next day 05:00'.
    Return None jika tidak valid atau shift tidak terjadwal.
    """
    if not isinstance(shift_text, str):
        return None
    s = shift_text.strip()
    if s in SKIP_SHIFTS:
        return None
    # Ambil semua pasangan HH:MM, pilih yang terakhir (jam selesai)
    matches = re.findall(r'(\d{1,2}):(\d{2})', s)
    if len(matches) >= 2:
        h, m = int(matches[-1][0]), int(matches[-1][1])
        return h * 60 + m
    return None


def parse_time_to_minutes(val) -> int | None:
    """Konversi berbagai format waktu ke menit sejak tengah malam."""
    if val is None:
        return None
    if isinstance(val, str):
        v = val.strip()
        if v.lower() in _NOT_PUNCHED:
            return None
        m = re.match(r'^(\d{1,2}):(\d{2})', v)
        if m:
            return int(m.group(1)) * 60 + int(m.group(2))
        return None
    if isinstance(val, time):
        return val.hour * 60 + val.minute
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.hour * 60 + val.minute
    if isinstance(val, pd.Timedelta):
        return (int(val.total_seconds()) % 86400) // 60
    if isinstance(val, float):
        if pd.isna(val):
            return None
        return round(val * 1440) % 1440
    return None


def has_punch(val) -> bool:
    """True jika nilai punch bukan 'not punched' / '--' / kosong."""
    return parse_time_to_minutes(val) is not None


def has_status(raw, status: str) -> bool:
    """Cek apakah status tertentu ada di list Klasifikasi_raw."""
    return isinstance(raw, list) and status in raw


def classify_shift_type(shift_text) -> str | None:
    """
    Tentukan tipe shift: 'Normal' (hari kerja) atau 'Off' (hari libur/rest).
    Return None jika shift tidak terjadwal / tidak dikenal.

    Mapping:
      "Rest"           → "Off"
      ""  / "--" / "Not scheduled" → None  (dilewati)
      Semua shift kerja lainnya (S1, S2, Night, dll.) → "Normal"
    """
    if not isinstance(shift_text, str):
        return None
    s = shift_text.strip()
    if s == "Rest":
        return "Off"
    if s in ("", "--", "Not scheduled"):
        return None
    # Semua shift kerja — termasuk S1, S2, Night, malam, dll. — dianggap "Normal"
    return "Normal"


def is_zero_or_dash(val) -> bool:
    """
    True jika nilai kolom count dianggap nol/kosong:
    "--", "0", "0.0", "" atau NaN.
    Digunakan untuk cek kolom 'Number of absences(Count)' dan 'K-Sick W Letter'.
    """
    if val is None:
        return True
    if isinstance(val, float):
        if pd.isna(val):
            return True
        return val == 0.0
    s = str(val).strip()
    return s in {"", "--", "0", "0.0", "nan"}


def is_dash_or_empty(val) -> bool:
    """
    True jika nilai dianggap tidak berisi data bermakna:
    None, NaN, "" (kosong), atau "--".
    Berbeda dari is_zero_or_dash — angka "0" atau "0.0" TIDAK dianggap kosong.
    Digunakan untuk cek kolom 'Offsite(Hour)' yang cukup diisi nilai apapun ≠ "--".
    """
    if val is None:
        return True
    if isinstance(val, float):
        return pd.isna(val)
    s = str(val).strip()
    return s in {"", "--", "nan"}


def parse_day_value(val) -> float | None:
    """
    Parse nilai day count dari kolom AL / UL / WFH (e.g. 0.5, 1, 1.0).
    Return None jika nol / kosong / tidak valid.
    Mendukung koma sebagai pemisah desimal (locale Indonesia: '0,5').
    """
    if val is None:
        return None
    if isinstance(val, float):
        return None if (pd.isna(val) or val == 0.0) else val
    if isinstance(val, int):
        return None if val == 0 else float(val)
    s = str(val).strip().replace(",", ".")
    if s in {"", "--", "0", "0.0", "nan"}:
        return None
    try:
        v = float(s)
        return None if v == 0.0 else v
    except ValueError:
        return None


def parse_duration_minutes(val) -> int | None:
    """
    Parse nilai durasi dalam menit dari kolom Duration of late arrival /
    Duration of early departure.
    Return None jika nol / kosong / tidak valid.
    """
    if val is None:
        return None
    if isinstance(val, float):
        if pd.isna(val) or val == 0.0:
            return None
        return int(val)
    if isinstance(val, int):
        return None if val == 0 else val
    s = str(val).strip()
    if s in {"", "--", "0", "0.0", "nan"}:
        return None
    try:
        v = float(s)
        return None if v == 0.0 else int(v)
    except ValueError:
        return None