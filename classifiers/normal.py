# classifiers/normal.py
"""
Klasifikasi Normal — karyawan hadir tepat waktu atau lebih awal.

Return: ["Normal"]
Dipanggil oleh __init__.classify() ketika:
  - att_result mengandung "normal" tanpa leave, ATAU
  - jam masuk ≤ jam mulai shift
"""


def classify() -> list[str]:
    """Kembalikan status Normal."""
    return ["Normal"]