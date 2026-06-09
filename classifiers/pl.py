# classifiers/pl.py
"""
Klasifikasi PL — Personal Leave (TKA).

Dipanggil oleh __init__.classify() ketika kolom
"PL-Personal(TKA)-私假(Day(s))" berisi nilai selain '--' / kosong / NaN.
(Nilai 0 / "0" dianggap tidak aktif dan TIDAK memicu PL.)

Return: ["PL"]
"""

from .base import is_zero_or_dash


def classify(pl_count) -> list[str] | None:
    if not is_zero_or_dash(pl_count):
        return ["PL"]
    return None