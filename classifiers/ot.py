# classifiers/ot.py
"""
Klasifikasi OT — Cuti Lainnya (Others Leave).

Dipanggil oleh __init__.classify() ketika kolom
"OT - Others - 其他(Day(s))" bernilai ≥ 1.

Return: ["OT"] atau ["1/2 OT"]
"""

from .base import parse_day_value


def classify(ot_count) -> list[str] | None:
    val = parse_day_value(ot_count)
    if val is None:
        return None
    if abs(val - 0.5) < 0.01:
        return ["1/2 OT"]
    if val >= 0.99:
        return ["OT"]
    return None