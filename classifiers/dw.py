# classifiers/dw.py
"""
Klasifikasi DW — karyawan tidak hadir (Absence).

Dipanggil oleh __init__.classify() ketika att_result mengandung "Absence".

Return: ["DW"]
"""


def classify() -> list[str]:
    """Kembalikan status DW (tidak hadir / Absence)."""
    return ["DW"]