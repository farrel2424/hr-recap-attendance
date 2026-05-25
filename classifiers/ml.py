# classifiers/ml.py
"""
Klasifikasi ML — Cuti Melahirkan (Maternity Leave).

Dipanggil oleh __init__.classify() ketika kolom
"ML-MaternityLeave-产假(Day(s))" berisi nilai selain '--' / kosong / NaN.
(Nilai 0 / "0" dianggap tidak aktif dan TIDAK memicu ML.)

Return: ["ML"]
"""

from .base import is_zero_or_dash


def classify(ml_count) -> list[str] | None:
    if not is_zero_or_dash(ml_count):
        return ["ML"]
    return None