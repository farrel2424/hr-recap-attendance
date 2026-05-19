# classifiers/__init__.py
"""
Orchestrator — titik masuk utama untuk semua logika klasifikasi absensi.

Fungsi publik:
  classify()      → list[str] | None
  classify_str()  → str | None   (versi joined untuk disimpan ke DB)

Flow (Urutan Prioritas):
  1. att_result = "Normal (rest)" / "Normal (not scheduled)"
       → Off                                                   ["Off"]

  2. Shift = Rest / Not scheduled / kosong / "--"
       → dilewati (None)

  3. att_result mengandung "Absence"
       → DW                                                    ["DW"]

  4. leave_app mengandung "K-Sick W Letter"
       → K (dual-count jika att juga Normal)                   ["Normal","K"] / ["K"]

  5. att_result mengandung "normal" + leave
       ├─ leave = AnnualLeave → AL / ½AL                       ["Normal","AL"] / ["Normal","1/2 AL"]
       └─ leave = WFH / WorkFromHome → WFA                     ["Normal","WFA"]

  6. [SUMBER A] Keterlambatan Punch In vs shift start
     Berlaku tanpa memandang att_result; dicek setelah Off/DW/K/AL/WFA
       ├─ terlambat 1–120 mnt  → Late                          ["Late"]    (standalone)
       └─ terlambat  > 120 mnt → 1/2 UL                       ["1/2 UL"]  (standalone)

  7. att_result mengandung "normal" + tidak ada keterlambatan
       → Normal                                                ["Normal"]

  8. [SUMBER B] att_result TIDAK mengandung "normal" + Punch In tepat/lebih awal
     Berdasarkan Punch Out vs shift end:
       hanya satu punch            → ["1/2 UL"]
       tidak ada punch             → None
       punch-out lebih cepat >2j   → ["1/2 UL"]
       punch-out lebih cepat ≤2j   → ["Late"]
       punch-out tepat/lebih lama  → ["Normal"]
"""

import pandas as pd

from .base         import parse_shift_start, parse_shift_end, SKIP_SHIFTS
from .normal       import classify as _classify_normal
from .late         import classify as _classify_late
from .late_in      import classify as _classify_late_in
from .annual_leave import classify as _classify_annual_leave
from .wfa          import classify as _classify_wfa
from .dw           import classify as _classify_dw
from .k_sick       import classify as _classify_k_sick
from .off          import classify as _classify_off, OFF_RESULTS

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
        list of str  — e.g. ["Normal"], ["Late"], ["1/2 UL"],
                            ["Normal","WFA"], ["Normal","AL"], ["Normal","1/2 AL"],
                            ["Normal","K"], ["K"], ["DW"], ["Off"]
        None         — shift dilewati atau tidak bisa diklasifikasi
    """
    att_str     = str(att_result).strip() if pd.notna(att_result) else ""
    att_lower   = att_str.lower()
    leave_str   = str(leave_app).strip() if (leave_app is not None and pd.notna(leave_app)) else ""
    leave_lower = leave_str.lower()
    shift_clean = str(shift_text).strip() if isinstance(shift_text, str) else ""

    # ── 1. Off ──────────────────────────────────────────────────────────────
    if att_str in OFF_RESULTS:
        return _classify_off()

    # ── 2. Lewati shift Rest / Not scheduled / kosong ───────────────────────
    shift_start = parse_shift_start(shift_text)

    if shift_clean in SKIP_SHIFTS or shift_start is None:
        return None

    # ── 3. DW ───────────────────────────────────────────────────────────────
    if "absence" in att_lower:
        return _classify_dw()

    # ── 4. K-Sick W Letter ──────────────────────────────────────────────────
    if "k-sick w letter" in leave_lower:
        return _classify_k_sick(has_normal="normal" in att_lower)

    # ── 5. AL / WFA (hanya saat att Normal + mengandung leave) ─────────────
    if "normal" in att_lower and "leave" in att_lower:
        if "annualleave" in leave_lower:
            return _classify_annual_leave(earliest_raw, latest_raw)
        if "workfromhome" in leave_lower or "wfh" in leave_lower:
            return _classify_wfa(earliest_raw, latest_raw)

    # ── 6. Keterlambatan Punch In — berlaku tanpa memandang att_result ───────
    #   1–120 mnt terlambat → ["Late"]
    #   > 120 mnt terlambat → ["1/2 UL"]
    late_in = _classify_late_in(earliest_raw, shift_start)
    if late_in:
        return late_in      # ["Late"] atau ["1/2 UL"] — tanpa prefix "Normal"

    # ── 7. Normal ───────────────────────────────────────────────────────────
    if "normal" in att_lower:
        return _classify_normal()

    # ── 8. Punch In tepat/lebih awal tapi att tidak normal → tidak terdefinisi
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
    Contoh: "Normal", "Late", "1/2 UL", "Normal/WFA", "Normal/AL", "DW", "Off", None
    """
    result = classify(earliest_raw, shift_text, att_result, latest_raw, leave_app)
    return "/".join(result) if result else None