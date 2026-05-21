# classifiers/annual_leave.py
"""
Klasifikasi Annual Leave (AL / ½AL).

Aturan:
  - Jika ada punch in DAN punch out → "1/2 AL"  (hadir setengah hari)
  - Jika tidak ada punch sama sekali → "AL"      (cuti penuh)

Tidak ada dual-count dengan S — status bersifat standalone.

Return:
  ["1/2 AL"]  — ada punch in & out
  ["AL"]      — tidak ada punch

Dipanggil oleh __init__.classify() ketika:
  - att_result mengandung "normal" + "leave"
  - leave_app mengandung "annualleave"
"""

from .base import has_punch


def classify(earliest_raw, latest_raw) -> list[str]:
    """
    Args:
        earliest_raw : nilai mentah jam masuk
        latest_raw   : nilai mentah jam keluar

    Returns:
        ["AL"] atau ["1/2 AL"]
    """
    has_in  = has_punch(earliest_raw)
    has_out = has_punch(latest_raw)

    if has_in and has_out:
        return ["1/2 AL"]
    return ["AL"]