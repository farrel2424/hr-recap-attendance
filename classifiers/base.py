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
    Tentukan tipe shift: 'S1', 'S2', atau 'H' (hari libur/rest).
    Return None jika shift tidak terjadwal.
    """
    if not isinstance(shift_text, str):
        return None
    s = shift_text.strip()
    if s == "Rest":
        return "H"
    if s in ("", "--", "Not scheduled"):
        return None
    s_lower = s.lower()
    if any(kw in s_lower for kw in ["s2", "night", "malam"]):
        return "S2"
    return "S1"