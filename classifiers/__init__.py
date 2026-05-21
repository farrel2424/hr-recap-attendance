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

  5. Kolom AnnualLeave - 印尼员工年假(Day(s))
       ├─ nilai = 0.5 → 1/2 AL                                 ["1/2 AL"]
       └─ nilai = 1   → AL                                     ["AL"]

  6. Kolom UL-Unpaid Leave-事假(Day(s))
       ├─ nilai = 1   → UL                                     ["UL"]
       └─ nilai = 0.5 → 1/2 UL                                 ["1/2 UL"]

  7. att_result mengandung "normal" + "leave" + leave = WFH
       → WFA                                                   ["WFA"]

  8. Kolom "Duration of late arrival(分钟)" (keterlambatan masuk)
       ├─ 1–120 menit → Late                                   ["Late"]    (standalone)
       └─  > 120 menit → 1/2 UL                               ["1/2 UL"]  (standalone)

  9. Kolom "Duration of early departure(分钟)" (pulang lebih awal)
       ├─ 1–120 menit → Late                                   ["Late"]    (standalone)
       └─  > 120 menit → 1/2 UL                               ["1/2 UL"]  (standalone)

  10. att_result TEPAT "Normal" atau "Normal（Correction of missed punch）"
        → S (Shift)                                            ["S"]

  11. Selain itu → tidak diklasifikasi (None)

Catatan penting:
  - Semua status bersifat STANDALONE — tidak ada dual-count.
  - K dan DW ditentukan oleh KOLOM spesifik, bukan att_result.
  - AL / ½AL ditentukan oleh kolom AnnualLeave, bukan punch presence.
  - UL / ½UL ditentukan oleh kolom UL-Unpaid Leave.
  - Late / ½UL dari keterlambatan ditentukan oleh kolom Duration (bukan hitung punch).
  - S hanya untuk att_result yang TEPAT sama, bukan mengandung kata "Normal".
"""

import pandas as pd

from .base         import parse_shift_start, parse_shift_end, SKIP_SHIFTS, S_ATT_RESULTS, is_zero_or_dash
from .normal       import classify as _classify_normal
from .late_in      import classify as _classify_late_in
from .late_out     import classify as _classify_late_out
from .annual_leave import classify as _classify_annual_leave
from .ul           import classify as _classify_ul
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
    parse_day_value,
    parse_duration_minutes,
)


def classify(
    earliest_raw,
    shift_text,
    att_result,
    latest_raw=None,
    leave_app=None,
    absences_count=None,
    k_sick_count=None,
    al_count=None,        # Kolom "AnnualLeave - 印尼员工年假(Day(s))"
    ul_count=None,        # Kolom "UL-Unpaid Leave-事假(Day(s))"
    duration_late=None,   # Kolom "Duration of late arrival(分钟)"
    duration_early=None,  # Kolom "Duration of early departure(分钟)"
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
        al_count       : nilai kolom "AnnualLeave - 印尼员工年假(Day(s))"
        ul_count       : nilai kolom "UL-Unpaid Leave-事假(Day(s))"
        duration_late  : nilai kolom "Duration of late arrival(分钟)"
        duration_early : nilai kolom "Duration of early departure(分钟)"

    Returns:
        list of str  — e.g. ["S"], ["Late"], ["1/2 UL"],
                            ["UL"], ["WFA"], ["AL"], ["1/2 AL"],
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
        # Jika ada punch valid → tampilkan tanpa klasifikasi ([] = "-")
        if has_punch(earliest_raw) and has_punch(latest_raw):
            return []
        return None

    # ── 3. K-Sick W Letter (kolom khusus) ───────────────────────────────────
    #   Cek SEBELUM DW agar sakit-dengan-surat tidak tertimpa DW
    if not is_zero_or_dash(k_sick_count):
        return _classify_k_sick()

    # ── 4. DW — Number of absences(Count) ───────────────────────────────────
    if not is_zero_or_dash(absences_count):
        return _classify_dw()

    # ── 5. AL — Kolom AnnualLeave ───────────────────────────────────────────
    #   nilai 0.5 → 1/2 AL   |   nilai 1 → AL
    al_result = _classify_annual_leave(al_count)
    if al_result:
        return al_result

    # ── 6. UL — Kolom UL-Unpaid Leave ───────────────────────────────────────
    #   nilai 1 → UL   |   nilai 0.5 → 1/2 UL
    ul_result = _classify_ul(ul_count)
    if ul_result:
        return ul_result

    # ── 7. WFA (att Normal + leave mengandung WFH) ──────────────────────────
    if "normal" in att_lower and "leave" in att_lower:
        if "workfromhome" in leave_lower or "wfh" in leave_lower:
            return _classify_wfa(earliest_raw, latest_raw)

    # ── 8. Keterlambatan masuk — Duration of late arrival ───────────────────
    #   1–120 mnt → Late   |   > 120 mnt → 1/2 UL
    late_in = _classify_late_in(duration_late)
    if late_in:
        return late_in

    # ── 9. Pulang lebih awal — Duration of early departure ──────────────────
    #   1–120 mnt → Late   |   > 120 mnt → 1/2 UL
    late_out = _classify_late_out(duration_early)
    if late_out:
        return late_out

    # ── 10. S (Shift) — att_result TEPAT "Normal" atau "Normal（Correction…）" ─
    if att_str in S_ATT_RESULTS:
        return _classify_normal()   # returns ["S"]

    # ── 11. Tidak diklasifikasi ──────────────────────────────────────────────
    return None


def classify_str(
    earliest_raw,
    shift_text,
    att_result,
    latest_raw=None,
    leave_app=None,
    absences_count=None,
    k_sick_count=None,
    al_count=None,
    ul_count=None,
    duration_late=None,
    duration_early=None,
) -> str | None:
    """
    Versi string dari classify() — untuk disimpan ke DB (dipisah '|').
    Menggunakan '|' bukan '/' agar tidak bertabrakan dengan '1/2 AL' / '1/2 UL'.
    Contoh: "S", "Late", "1/2 UL", "UL", "WFA", "AL", "DW", "Off", None
    """
    result = classify(
        earliest_raw, shift_text, att_result,
        latest_raw=latest_raw,
        leave_app=leave_app,
        absences_count=absences_count,
        k_sick_count=k_sick_count,
        al_count=al_count,
        ul_count=ul_count,
        duration_late=duration_late,
        duration_early=duration_early,
    )
    if result is None:
        return None
    return "|".join(result)  # [] → "" (empty string, disimpan ke DB sebagai bukan NULL)