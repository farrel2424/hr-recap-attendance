# classifiers/wfa.py
"""
Klasifikasi WFA — Work From Anywhere / Work From Home.

Return:
  ["Normal", "WFA"]  — selalu, karena dipanggil hanya saat att_result mengandung "Normal"
"""


def classify(earliest_raw, latest_raw) -> list[str]:
    return ["Normal", "WFA"]