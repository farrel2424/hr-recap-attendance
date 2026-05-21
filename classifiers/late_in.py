# classifiers/late_in.py
"""
Klasifikasi keterlambatan Punch In (Earliest) vs jam mulai shift.

Aturan (menggunakan K_THRESHOLD_MIN = 120 menit):
  Punch in terlambat ≤ 120 menit dari shift start → ["Late"]
  Punch in terlambat  > 120 menit dari shift start → ["1/2 UL"]
  Punch in tepat waktu / lebih awal / tidak ada punch → None

Return: list satu elemen atau None
"""

from .base import parse_time_to_minutes, has_punch, K_THRESHOLD_MIN


def classify(earliest_raw, shift_start: int) -> list[str] | None:
    if not has_punch(earliest_raw):
        return None

    earliest = parse_time_to_minutes(earliest_raw)
    if earliest is None:
        return None

    diff = earliest - shift_start
    if diff <= 0:
        return None
    elif diff <= K_THRESHOLD_MIN:
        return ["Late"]
    else:
        return ["1/2 UL"]