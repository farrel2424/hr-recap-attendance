# classifiers/annual_leave.py
"""
Klasifikasi Annual Leave (AL / ½AL).

Aturan dual-count:
  - Selalu masuk Normal (karena att_result mengandung "Normal")
  - Jika ada punch in DAN punch out → tambah "1/2 AL"  (hadir setengah hari)
  - Jika tidak ada punch sama sekali → tambah "AL"      (cuti penuh)

Return:
  ["Normal", "1/2 AL"]  — ada punch in & out
  ["Normal", "AL"]      — tidak ada punch

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
        ["Normal", "AL"] atau ["Normal", "1/2 AL"]
    """
    has_in  = has_punch(earliest_raw)
    has_out = has_punch(latest_raw)

    if has_in and has_out:
        return ["Normal", "1/2 AL"]
    return ["Normal", "AL"]