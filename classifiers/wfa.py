# classifiers/wfa.py
"""
Klasifikasi WFA — Work From Anywhere / Work From Home.

Return:
  ["WFA"]  — selalu, bersifat standalone (tidak dual-count dengan S).
"""


def classify(earliest_raw, latest_raw) -> list[str]:
    return ["WFA"]