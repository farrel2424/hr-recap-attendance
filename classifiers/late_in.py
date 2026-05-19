# classifiers/late_in.py
"""
Klasifikasi keterlambatan Punch In (Earliest) vs jam mulai shift.

Dipanggil oleh __init__.classify() saat att_result mengandung "normal"
dan tidak ada kondisi leave khusus (bukan AL / WFA / K-Sick),
sehingga tidak terjadi triple-count.

Aturan (menggunakan K_THRESHOLD_MIN = 120 menit):
  Punch in terlambat ≤ 120 menit dari shift start → ["Late"]
  Punch in terlambat  > 120 menit dari shift start → ["1/2 UL"]
  Punch in tepat waktu / lebih awal / tidak ada punch → None

Return: list satu elemen atau None
"""

from .base import parse_time_to_minutes, has_punch, K_THRESHOLD_MIN


def classify(earliest_raw, shift_start: int) -> list[str] | None:
    """
    Args:
        earliest_raw : nilai mentah jam masuk (kolom Earliest)
        shift_start  : jam mulai shift dalam menit sejak tengah malam
                       (hasil parse_shift_start)

    Returns:
        ["Late"]    — terlambat masuk 1–120 menit
        ["1/2 UL"]  — terlambat masuk > 120 menit
        None        — tepat waktu / lebih awal / tidak ada punch
    """
    if not has_punch(earliest_raw):
        return None

    earliest = parse_time_to_minutes(earliest_raw)
    if earliest is None:
        return None

    diff = earliest - shift_start   # positif = terlambat masuk
    if diff <= 0:
        return None                 # hadir tepat waktu atau lebih awal
    elif diff <= K_THRESHOLD_MIN:
        return ["Late"]
    else:
        return ["1/2 UL"]