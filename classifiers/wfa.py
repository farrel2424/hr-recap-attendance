# classifiers/wfa.py
"""
Klasifikasi WFA — Work From Anywhere / Work From Home.

Aturan:
  - Jika leave_app mengandung "WFH" atau "WorkFromHome" → selalu WFA
  - Jika att_result juga mengandung "Normal" → double count: Normal + WFA
    (pengecekan Normal dilakukan di __init__.py sebelum memanggil fungsi ini)
  - Tidak perlu ada punch in / punch out

Return:
  ["Normal", "WFA"]  — selalu, karena dipanggil hanya saat att_result mengandung "Normal"
"""


def classify(earliest_raw, latest_raw) -> list[str]:
    """
    Args:
        earliest_raw : nilai mentah jam masuk (tidak digunakan, dipertahankan agar signature konsisten)
        latest_raw   : nilai mentah jam keluar (tidak digunakan)

    Returns:
        ["Normal", "WFA"] — selalu
    """
    return ["Normal", "WFA"]