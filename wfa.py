# classifiers/wfa.py
"""
Klasifikasi WFA — Work From Anywhere / Work From Home.

Aturan dual-count:
  - Selalu masuk Normal (karena att_result mengandung "Normal")
  - Jika ada punch in DAN punch out → tambah "WFA"
  - Jika tidak ada punch           → cukup Normal (WFH tanpa presensi = Normal biasa)

Return:
  ["Normal", "WFA"]  — ada punch in & out
  ["Normal"]         — tidak ada punch

Dipanggil oleh __init__.classify() ketika:
  - att_result mengandung "normal" + "leave"
  - leave_app mengandung "workfromhome" atau "wfh"
"""

from .base import has_punch


def classify(earliest_raw, latest_raw) -> list[str]:
    """
    Args:
        earliest_raw : nilai mentah jam masuk
        latest_raw   : nilai mentah jam keluar

    Returns:
        ["Normal", "WFA"] atau ["Normal"]
    """
    if has_punch(earliest_raw) and has_punch(latest_raw):
        return ["Normal", "WFA"]
    return ["Normal"]