# classifiers/late_out.py
"""
Klasifikasi keterlambatan Punch Out (Latest) vs jam selesai shift.
"""

from .base import parse_time_to_minutes, has_punch, K_THRESHOLD_MIN


def _normalize_overnight(shift_start: int, shift_end: int, latest_min: int) -> tuple[int, int]:
    if shift_end < shift_start:
        shift_end += 1440
        if latest_min < shift_start:
            latest_min += 1440
    return shift_end, latest_min


def classify(latest_raw, shift_end: int, shift_start: int = 0) -> list[str] | None:
    if not has_punch(latest_raw):
        return None

    latest = parse_time_to_minutes(latest_raw)
    if latest is None:
        return None

    norm_end, norm_latest = _normalize_overnight(shift_start, shift_end, latest)

    diff = norm_end - norm_latest
    if diff <= 0:
        return None
    elif diff <= K_THRESHOLD_MIN:
        return ["Late"]
    else:
        return ["1/2 UL"]