# classifiers/rl.py
"""
Klasifikasi RL — Roster Leave.

Dipanggil oleh __init__.classify() ketika kolom
"RL - Roster Leave(Day(s))" berisi nilai selain '--' / kosong / NaN.
(Nilai 0 / "0" dianggap tidak aktif dan TIDAK memicu RL.)

Return: ["RL"]
"""

from .base import is_dash_or_empty


def classify(rl_count) -> list[str] | None:
    if not is_dash_or_empty(rl_count):
        return ["RL"]
    return None