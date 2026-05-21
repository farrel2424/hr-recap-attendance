# classifiers/__init__.py
"""
Orchestrator — titik masuk utama untuk semua logika klasifikasi absensi.

Fungsi publik:
  classify()      → list[str] | None
  classify_str()  → str | None   (versi joined dengan '|' untuk disimpan ke DB)

Flow (Urutan Prioritas):
  1. att_result = "Normal (rest)" / "Normal (not scheduled)"
       → Off                                                   ["Off"]

  2. att_result = "Normal (Offsite)" DAN kolom Offsite(Hour) ≠ "--" / kosong
       → WFS                                                   ["WFS"]

  3. Shift = Rest / Not scheduled / kosong / "--"
       → dilewati (None)

  4. Kolom K-Sick W Letter ≠ "0" / "--" / kosong
       → K                                                     ["K"]

  5. Kolom Number of absences(Count) ≠ "0" / "--" / kosong
       → DW                                                    ["DW"]

  6. Kolom AnnualLeave - 印尼员工年假(Day(s))
       ├─ nilai = 0.5 → 1/2 AL                                 ["1/2 AL"]
       └─ nilai = 1   → AL                                     ["AL"]

  7. Kolom UL-Unpaid Leave-事假(Day(s))
       ├─ nilai = 1   → UL                                     ["UL"]
       └─ nilai = 0.5 → 1/2 UL                                 ["1/2 UL"]

  8. Kolom WFH-WorkFromHome-家办公(Day(s))
       ├─ nilai = 0.5 → 1/2 WFA                                ["1/2 WFA"]
       └─ nilai = 1   → WFA                                    ["WFA"]

  9. Kolom "Duration of late arrival(分钟)" (keterlambatan masuk)
       ├─ 1–120 menit → Late                                   ["Late"]    (standalone)
       └─  > 120 menit → 1/2 UL                               ["1/2 UL"]  (standalone)

  10. Kolom "Duration of early departure(分钟)" (pulang lebih awal)
       ├─ 1–120 menit → Late                                   ["Late"]    (standalone)
       └─  > 120 menit → 1/2 UL                               ["1/2 UL"]  (standalone)

  11. att_result TEPAT "Normal" atau "Normal（Correction of missed punch）"
        → S (Shift)                                            ["S"]

  12. Selain itu → tidak diklasifikasi (None)

Catatan penting:
  - Semua status bersifat STANDALONE — tidak ada dual-count.
  - K dan DW ditentukan oleh KOLOM spesifik, bukan att_result.
  - AL / ½AL ditentukan oleh kolom AnnualLeave, bukan punch presence.
  - UL / ½UL ditentukan oleh kolom UL-Unpaid Leave.
  - WFA / ½WFA ditentukan oleh kolom WFH-WorkFromHome (bukan att_result + leave string).
  - WFS ditentukan oleh att_result = "Normal (Offsite)" DAN kolom Offsite(Hour) ≠ "--".
  - Late / ½UL dari keterlambatan ditentukan oleh kolom Duration (bukan hitung punch).
  - S hanya untuk att_result yang TEPAT sama, bukan mengandung kata "Normal".
"""

import pandas as pd

from .base         import parse_shift_start, parse_shift_end, SKIP_SHIFTS, S_ATT_RESULTS, is_zero_or_dash, is_dash_or_empty
from .normal       import classify as _classify_normal
from .late_in      import classify as _classify_late_in
from .late_out     import classify as _classify_late_out
from .annual_leave import classify as _classify_annual_leave
from .ul           import classify as _classify_ul
from .wfa          import classify as _classify_wfa
from .wfs          import classify as _classify_wfs
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
    is_dash_or_empty,
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
    wfh_count=None,       # Kolom "WFH-WorkFromHome-家办公(Day(s))"
    offsite_hour=None,    # Kolom "Offsite(Hour)"
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
        wfh_count      : nilai kolom "WFH-WorkFromHome-家办公(Day(s))"
        offsite_hour   : nilai kolom "Offsite(Hour)"

    Returns:
        list of str  — e.g. ["S"], ["Late"], ["1/2 UL"],
                            ["UL"], ["WFA"], ["1/2 WFA"], ["WFS"],
                            ["AL"], ["1/2 AL"], ["K"], ["DW"], ["Off"]
        None         — shift dilewati atau tidak bisa diklasifikasi
    """
    att_str     = str(att_result).strip() if pd.notna(att_result) else ""
    shift_clean = str(shift_text).strip() if isinstance(shift_text, str) else ""

    # ── 1. Off ──────────────────────────────────────────────────────────────
    if att_str in OFF_RESULTS:
        return _classify_off()

    # ── 2. WFS — Normal (Offsite) ────────────────────────────────────────────
    #   att_result mengandung "Offsite" + kolom Offsite(Hour) terisi (bukan "--"/kosong)
    #   (Mendukung variasi tanda kurung biasa '()' maupun full-width '（）' khas Excel)
    if "offsite" in att_str.lower() and not is_dash_or_empty(offsite_hour):
        return _classify_wfs()

    # ── 3. Lewati shift Rest / Not scheduled / kosong ───────────────────────
    shift_start = parse_shift_start(shift_text)
    if shift_clean in SKIP_SHIFTS or shift_start is None:
        if has_punch(earliest_raw) and has_punch(latest_raw):
            return []
        return None

    # ── 4. K-Sick W Letter (kolom khusus) ───────────────────────────────────
    #   Cek SEBELUM DW agar sakit-dengan-surat tidak tertimpa DW
    if not is_zero_or_dash(k_sick_count):
        return _classify_k_sick()

    # ── 5. DW — Number of absences(Count) ───────────────────────────────────
    if not is_zero_or_dash(absences_count):
        return _classify_dw()

    # ── 6. AL — Kolom AnnualLeave ───────────────────────────────────────────
    #   nilai 0.5 → 1/2 AL   |   nilai 1 → AL
    al_result = _classify_annual_leave(al_count)
    if al_result:
        return al_result

    # ── 7. UL — Kolom UL-Unpaid Leave ───────────────────────────────────────
    #   nilai 1 → UL   |   nilai 0.5 → 1/2 UL
    ul_result = _classify_ul(ul_count)
    if ul_result:
        return ul_result

    # ── 8. WFA / 1/2 WFA — Kolom WFH-WorkFromHome ───────────────────────────
    #   nilai 1 → WFA   |   nilai 0.5 → 1/2 WFA
    wfa_result = _classify_wfa(wfh_count)
    if wfa_result:
        return wfa_result

    # ── 9. Keterlambatan masuk — Duration of late arrival ───────────────────
    #   1–120 mnt → Late   |   > 120 mnt → 1/2 UL
    late_in = _classify_late_in(duration_late)
    if late_in:
        return late_in

    # ── 10. Pulang lebih awal — Duration of early departure ─────────────────
    #   1–120 mnt → Late   |   > 120 mnt → 1/2 UL
    late_out = _classify_late_out(duration_early)
    if late_out:
        return late_out

    # ── 11. S (Shift) — att_result TEPAT "Normal" atau "Normal（Correction…）" ─
    if att_str in S_ATT_RESULTS:
        return _classify_normal()   # returns ["S"]

    # ── 12. Tidak diklasifikasi ──────────────────────────────────────────────
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
    wfh_count=None,
    offsite_hour=None,
) -> str | None:
    """
    Versi string dari classify() — untuk disimpan ke DB (dipisah '|').
    Menggunakan '|' bukan '/' agar tidak bertabrakan dengan '1/2 AL' / '1/2 UL' / '1/2 WFA'.
    Contoh: "S", "Late", "1/2 UL", "UL", "WFA", "1/2 WFA", "WFS", "AL", "DW", "Off", None
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
        wfh_count=wfh_count,
        offsite_hour=offsite_hour,
    )
    if result is None:
        return None
    return "|".join(result)  # [] → "" (empty string, disimpan ke DB sebagai bukan NULL)