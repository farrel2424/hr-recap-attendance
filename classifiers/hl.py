# classifiers/hl.py
"""
Klasifikasi HL — Cuti Pernikahan (Happy/Marry Leave).

Dipanggil oleh __init__.classify() ketika kolom
"HL-Happy(Marry)-婚假(Day(s))" bernilai ≥ 1.

Return: ["HL"]
"""

from .base import parse_day_value


def classify(hl_count) -> list[str] | None:
    val = parse_day_value(hl_count)
    if val is not None and val >= 0.99:
        return ["HL"]
    return None