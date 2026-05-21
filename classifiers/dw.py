# classifiers/dw.py
"""
Klasifikasi DW — karyawan tidak hadir (Absence).

Dipanggil oleh __init__.classify() ketika kolom
"Number of absences(Count)" bernilai BUKAN "0", "--", atau kosong.

Return: ["DW"]
"""


def classify() -> list[str]:
    """Kembalikan status DW (tidak hadir / Absence)."""
    return ["DW"]