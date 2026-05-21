# classifiers/__init__.py
"""
Orchestrator — titik masuk utama untuk semua logika klasifikasi absensi.

Fungsi publik:
  classify()      → list[str] | None
  classify_str()  → str | None   (versi joined dengan '|' untuk disimpan ke DB)

Flow (Urutan Prioritas):
  1. att_result = "Normal (rest)" / "Normal (not scheduled)"
       → Off                                                   ["Off"]

  2. Shift = Rest / Not scheduled / kosong / "--"
       → dilewati (None)

  3. Kolom K-Sick W Letter ≠ "0" / "--" / kosong
       → K                                                     ["K"]

  4. Kolom Number of absences(Count) ≠ "0" / "--" / kosong
       → DW                                                    ["DW"]

  5. att_result mengandung "normal" + "leave"
       ├─ leave = AnnualLeave → AL / ½AL                       ["AL"] / ["1/2 AL"]
       └─ leave = WFH / WorkFromHome → WFA                     ["WFA"]

  6a. [SUMBER A] Keterlambatan Punch In vs shift start
      Berlaku tanpa memandang att_result:
        ├─ terlambat 1–120 mnt  → Late                         ["Late"]    (standalone)
        └─ terlambat  > 120 mnt → 1/2 UL                      ["1/2 UL"]  (standalone)

  6b. [SUMBER B] Kepulangan Lebih Awal — Punch Out vs shift end
      Hanya berjalan jika langkah 6a tidak menghasilkan klasifikasi:
        ├─ lebih awal  1–120 mnt dari shift end → Late         ["Late"]    (standalone)
        └─ lebih awal  > 120 mnt dari shift end → 1/2 UL      ["1/2 UL"]  (standalone)

  7. att_result TEPAT "Normal" atau "Normal（Correction of missed punch）"
       + tidak ada keterlambatan
       → S (Shift)                                             ["S"]

  8. Selain itu → tidak diklasifikasi (None)

Catatan penting:
  - Semua status bersifat STANDALONE — tidak ada dual-count.
  - K dan DW ditentukan oleh KOLOM spesifik, bukan att_result.
  - S hanya untuk att_result yang TEPAT sama, bukan mengandung kata "Normal".
"""

import pandas as pd

from .base         import parse_shift_start, parse_shift_end, SKIP_SHIFTS, S_ATT_RESULTS, is_zero_or_dash
from .normal       import classify as _classify_normal
from .late_in      import classify as _classify_late_in
from .late_out     import classify as _classify_late_out
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
    is_zero_or_dash,
)


def classify(
    earliest_raw,
    shift_text,
    att_result,
    latest_raw=None,
    leave_app=None,
    absences_count=None,
    k_sick_count=None,
) -> list[str] | None:
    """
    Klasifikasi satu baris absensi.

    Args:
        earliest_raw   : nilai mentah jam masuk (kolom Earliest)
        shift_text     : teks shift (kolom Shift)
        att_result     : hasil absensi (kolom Attendance results)
        latest_raw     : nilai mentah jam keluar (kolom Latest)
        leave_app      : aplikasi leave (kolom Leave & Overtime Application)
        absences_count : nilai kolom "Number of absences(Count)"
        k_sick_count   : nilai kolom "K-Sick W Letter-病假有信(Day(s))"

    Returns:
        list of str  — e.g. ["S"], ["Late"], ["1/2 UL"],
                            ["WFA"], ["AL"], ["1/2 AL"],
                            ["K"], ["DW"], ["Off"]
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

    # ── 3. K-Sick W Letter (kolom khusus) ───────────────────────────────────
    #   Cek SEBELUM DW agar sakit-dengan-surat tidak tertimpa DW
    if not is_zero_or_dash(k_sick_count):
        return _classify_k_sick()

    # ── 4. DW — Number of absences(Count) ───────────────────────────────────
    if not is_zero_or_dash(absences_count):
        return _classify_dw()

    # ── 5. AL / WFA (att Normal + mengandung leave) ─────────────────────────
    if "normal" in att_lower and "leave" in att_lower:
        if "annualleave" in leave_lower:
            return _classify_annual_leave(earliest_raw, latest_raw)
        if "workfromhome" in leave_lower or "wfh" in leave_lower:
            return _classify_wfa(earliest_raw, latest_raw)

    # ── 6a. Keterlambatan Punch In — berlaku tanpa memandang att_result ──────
    late_in = _classify_late_in(earliest_raw, shift_start)
    if late_in:
        return late_in      # ["Late"] atau ["1/2 UL"]

    # ── 6b. Kepulangan Lebih Awal — Punch Out vs shift end ───────────────────
    shift_end = parse_shift_end(shift_text)
    if shift_end is not None:
        late_out = _classify_late_out(latest_raw, shift_end, shift_start)
        if late_out:
            return late_out  # ["Late"] atau ["1/2 UL"]

    # ── 7. S (Shift) — att_result TEPAT "Normal" atau "Normal（Correction…）" ─
    if att_str in S_ATT_RESULTS:
        return _classify_normal()   # returns ["S"]

    # ── 8. Tidak diklasifikasi ───────────────────────────────────────────────
    return None


def classify_str(
    earliest_raw,
    shift_text,
    att_result,
    latest_raw=None,
    leave_app=None,
    absences_count=None,
    k_sick_count=None,
) -> str | None:
    """
    Versi string dari classify() — untuk disimpan ke DB (dipisah '|').
    Menggunakan '|' bukan '/' agar tidak bertabrakan dengan '1/2 AL' / '1/2 UL'.
    Contoh: "S", "Late", "1/2 UL", "WFA", "AL", "DW", "Off", None
    """
    result = classify(
        earliest_raw, shift_text, att_result,
        latest_raw=latest_raw,
        leave_app=leave_app,
        absences_count=absences_count,
        k_sick_count=k_sick_count,
    )
    return "|".join(result) if result else None