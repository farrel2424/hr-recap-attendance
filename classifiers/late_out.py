# classifiers/late_out.py
"""
Klasifikasi pulang lebih awal berdasarkan kolom
"Duration of early departure(分钟)".

Aturan (menggunakan K_THRESHOLD_MIN = 120 menit):
  duration  1–120 menit → ["Late"]
  duration  > 120 menit → ["1/2 UL"]
  duration 0 / kosong   → None

Return: list satu elemen atau None
"""

from .base import parse_duration_minutes, K_THRESHOLD_MIN


def classify(duration_early) -> list[str] | None:
    """
    Args:
        duration_early : nilai kolom "Duration of early departure(分钟)"

    Returns:
        ["Late"], ["1/2 UL"], atau None
    """
    minutes = parse_duration_minutes(duration_early)
    if minutes is None or minutes <= 0:
        return None
    if minutes <= K_THRESHOLD_MIN:
        return ["Late"]
    return ["1/2 UL"]