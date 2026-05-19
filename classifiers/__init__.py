# classifiers/__init__.py
"""
Orchestrator — titik masuk utama untuk semua logika klasifikasi absensi.

Fungsi publik:
  classify()      → list[str] | None
  classify_str()  → str | None   (versi joined untuk disimpan ke DB)

Flow:
  att_result = "Normal (rest)" / "Normal (not scheduled)"
    → Off                                                    ["Off"]

  att_result mengandung "Absence"
    → DW                                                     ["DW"]

  leave_app mengandung "K-Sick W Letter"
    → K (dual-count jika att juga Normal)                    ["Normal","K"] / ["K"]

  att_result mengandung "normal"
    ├─ leave_app = AnnualLeave  → AL / ½AL                   ["Normal","AL"] / ["Normal","1/2 AL"]
    ├─ leave_app = WFH/WorkFromHome → WFA                    ["Normal","WFA"]
    └─ lain-lain (Normal biasa)
         → cek Punch In vs shift start:
              terlambat ≤ 120 mnt → dual-count Late          ["Normal","Late"]
              terlambat  > 120 mnt → dual-count ½UL          ["Normal","1/2 UL"]
              tepat/lebih awal    → Normal murni              ["Normal"]

  att_result TIDAK mengandung "normal" (Early Departure, Missed Punch, dll)
    → Berdasarkan Punch Out vs shift end:
        hanya satu punch            → ["1/2 UL"]
        tidak ada punch             → None
        punch-out lebih cepat >2j   → ["1/2 UL"]
        punch-out lebih cepat ≤2j   → ["Late"]
        punch-out tepat/lebih lama  → ["Normal"]
"""

import pandas as pd

from .base          import parse_shift_start, parse_shift_end, SKIP_SHIFTS
from .normal        import classify as _classify_normal
from .late          import classify as _classify_late
from .late_in       import classify as _classify_late_in
from .annual_leave  import classify as _classify_annual_leave
from .wfa           import classify as _classify_wfa
from .dw            import classify as _classify_dw
from .k_sick        import classify as _classify_k_sick
from .off           import classify as _classify_off, OFF_RESULTS

# Re-export helpers yang dipakai di app.py / db.py
from .base import (         # noqa: F401
    parse_shift_start,
    parse_shift_end,
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
        list of str  — e.g. ["Normal"], ["Normal","Late"], ["Normal","1/2 UL"],
                            ["Normal","WFA"], ["Late"], ["1/2 UL"], ["DW"], ["Off"]
        None         — shift dilewati (Rest / Not scheduled / dll) atau tidak ada data
    """
    att_str   = str(att_result).strip() if pd.notna(att_result) else ""
    att_lower = att_str.lower()
    leave_str = str(leave_app).strip() if (leave_app is not None and pd.notna(leave_app)) else ""
    leave_lower = leave_str.lower()
    shift_clean = str(shift_text).strip() if isinstance(shift_text, str) else ""

    # ── Off: att_result bernilai "Normal (rest)" / "Normal (not scheduled)" ─
    if att_str in OFF_RESULTS:
        return _classify_off()

    # ── Lewati shift Rest / Not scheduled / kosong ──────────────────────────
    shift_start = parse_shift_start(shift_text)
    shift_end   = parse_shift_end(shift_text)

    if shift_clean in SKIP_SHIFTS or shift_start is None:
        return None

    # ── DW: att_result mengandung "Absence" ─────────────────────────────────
    if "absence" in att_lower:
        return _classify_dw()

    # ── K-Sick W Letter (cek sebelum Normal agar tidak tertukar AL/WFA) ─────
    if "k-sick w letter" in leave_lower:
        return _classify_k_sick(has_normal="normal" in att_lower)

    # ── Attendance result mengandung "normal" ───────────────────────────────
    if "normal" in att_lower:
        if "leave" in att_lower:
            if "annualleave" in leave_lower:
                return _classify_annual_leave(earliest_raw, latest_raw)
            if "workfromhome" in leave_lower or "wfh" in leave_lower:
                return _classify_wfa(earliest_raw, latest_raw)

        # Normal biasa — cek tambahan keterlambatan Punch In vs shift start
        late_in = _classify_late_in(earliest_raw, shift_start)
        if late_in:
            return ["Normal"] + late_in   # ["Normal","Late"] atau ["Normal","1/2 UL"]
        return _classify_normal()

    # ── Tidak ada "normal" → cek pulang lebih awal / hanya satu punch ───────
    if shift_end is not None:
        return _classify_late(earliest_raw, latest_raw, shift_end, shift_start)

    # Fallback: shift_end tidak bisa di-parse → tidak bisa diklasifikasi
    return None


def classify_str(
    earliest_raw,
    shift_text,
    att_result,
    latest_raw=None,
    leave_app=None,
) -> str | None:
    """
    Versi string dari classify() — untuk disimpan ke DB (dipisah '/').
    Contoh: "Normal/Late", "Normal/1/2 UL", "Normal/WFA", "Late", "DW", "Off", None
    """
    result = classify(earliest_raw, shift_text, att_result, latest_raw, leave_app)
    return "/".join(result) if result else None