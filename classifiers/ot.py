# classifiers/ot.py
"""
Klasifikasi OT — Cuti Lainnya (Others Leave).

Dipanggil oleh __init__.classify() ketika kolom
"OT - Others - 其他(Day(s))" bernilai ≥ 1.

Return: ["OT"]
"""

from .base import parse_day_value


def classify(ot_count) -> list[str] | None:
    val = parse_day_value(ot_count)
    if val is not None and val >= 0.99:
        return ["OT"]
    return None