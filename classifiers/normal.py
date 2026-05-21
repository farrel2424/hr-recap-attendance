# classifiers/normal.py
"""
Klasifikasi S (Shift) — karyawan hadir tepat waktu atau lebih awal.

Return: ["S"]
Dipanggil oleh __init__.classify() ketika:
  - att_result bernilai TEPAT "Normal" atau "Normal（Correction of missed punch）"
  - punch in ≤ jam mulai shift (tidak terlambat)
"""


def classify() -> list[str]:
    """Kembalikan status S (Shift)."""
    return ["S"]