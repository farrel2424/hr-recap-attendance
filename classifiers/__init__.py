# classifiers/__init__.py
"""
Orchestrator — titik masuk utama untuk semua logika klasifikasi absensi.

Fungsi publik:
  classify()      → list[str] | None
  classify_str()  → str | None   (versi joined untuk disimpan ke DB)

Flow:
                         ┌─ "annualleave" ──► annual_leave.classify()  → ["Normal","AL"/"1/2 AL"]
  att_result             │
  mengandung "normal" ───┤─ "wfh/workfromhome" ► wfa.classify()        → ["Normal","WFA"/"Normal"]
                         │
                         └─ leave lain / tanpa leave ► normal.classify() → ["Normal"]

  att_result TIDAK
  mengandung "normal" ──────────────────────────► late.classify()       → ["Late"/"1/2 UL"/"Normal"]
"""

import pandas as pd

from .base          import parse_shift_start, SKIP_SHIFTS
from .normal        import classify as _classify_normal
from .late          import classify as _classify_late
from .annual_leave  import classify as _classify_annual_leave
from .wfa           import classify as _classify_wfa

# Re-export helpers yang dipakai di app.py / db.py
from .base import (         # noqa: F401
    parse_shift_start,
    parse_time_to_minutes,
    has_punch,
    has_status,
    classify_shift_type,
    SKIP_SHIFTS,
    K_THRESHOLD_MIN,
    _NOT_PUNCHED,
)


def classify(
    earliest_raw,
    shift_text,
    att_result,
    latest_raw=None,
    leave_app=None,
) -> list[str] | None:
    """
    Klasifikasi satu baris absensi.

    Returns:
        list of str  — e.g. ["Normal"], ["Normal","WFA"], ["Late"], ["1/2 UL"]
        None         — shift dilewati (Rest / Not scheduled / dll)
    """
    att_lower   = str(att_result).strip().lower() if pd.notna(att_result) else ""
    shift_start = parse_shift_start(shift_text)
    shift_clean = str(shift_text).strip() if isinstance(shift_text, str) else ""
    leave_str   = str(leave_app).strip() if (leave_app is not None and pd.notna(leave_app)) else ""

    # Lewati shift Rest / Not scheduled / kosong
    if shift_clean in SKIP_SHIFTS or shift_start is None:
        return None

    # ── Attendance result mengandung "normal" ───────────────
    if "normal" in att_lower:
        if "leave" in att_lower:
            leave_lower = leave_str.lower()

            if "annualleave" in leave_lower:
                return _classify_annual_leave(earliest_raw, latest_raw)

            if "workfromhome" in leave_lower or "wfh" in leave_lower:
                return _classify_wfa(earliest_raw, latest_raw)

        # Normal biasa (tepat waktu, atau leave jenis lain)
        return _classify_normal()

    # ── Tidak ada "normal" → cek keterlambatan ──────────────
    return _classify_late(earliest_raw, shift_start)


def classify_str(
    earliest_raw,
    shift_text,
    att_result,
    latest_raw=None,
    leave_app=None,
) -> str | None:
    """
    Versi string dari classify() — untuk disimpan ke DB (dipisah '/').
    Contoh: "Normal/WFA", "Normal/AL", "Late", None
    """
    result = classify(earliest_raw, shift_text, att_result, latest_raw, leave_app)
    return "/".join(result) if result else None