# classifiers/late.py
"""
Klasifikasi keterlambatan standar.

Return:
  ["Normal"]   — tidak terlambat (diff ≤ 0)
  ["Late"]     — terlambat ≤ K_THRESHOLD_MIN (2 jam)
  ["1/2 UL"]   — terlambat > K_THRESHOLD_MIN (lebih dari 2 jam)
  None         — tidak ada data punch in

Dipanggil oleh __init__.classify() ketika att_result TIDAK mengandung "normal".
"""

from .base import parse_time_to_minutes, K_THRESHOLD_MIN


def classify(earliest_raw, shift_start: int) -> list[str] | None:
    """
    Args:
        earliest_raw : nilai mentah jam masuk (string / time / float / dll)
        shift_start  : jam mulai shift dalam menit (sudah di-parse)

    Returns:
        list satu elemen, atau None jika tidak ada punch in.
    """
    earliest = parse_time_to_minutes(earliest_raw)
    if earliest is None:
        return None

    diff = earliest - shift_start
    if diff <= 0:
        return ["Normal"]
    elif diff <= K_THRESHOLD_MIN:
        return ["Late"]
    else:
        return ["1/2 UL"]