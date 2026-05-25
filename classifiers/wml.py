# classifiers/wml.py
"""
Klasifikasi WML — Cuti Istri Melahirkan (Wife Maternity Leave).

Dipanggil oleh __init__.classify() ketika kolom
"WML-WifeMater-妻产假(Day(s))" bernilai ≥ 1.

Return: ["WML"]
"""

from .base import parse_day_value


def classify(wml_count) -> list[str] | None:
    val = parse_day_value(wml_count)
    if val is not None and val >= 0.99:
        return ["WML"]
    return None