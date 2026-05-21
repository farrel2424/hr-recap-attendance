# classifiers/shift.py
"""
Klasifikasi S (Shift) — karyawan hadir penuh sesuai jadwal shift.

Dipanggil oleh __init__.classify() ketika att_result bernilai TEPAT:
  - "Normal"
  - "Normal（Correction of missed punch）"

Return: ["S"]
"""

# Nilai att_result yang diklasifikasikan sebagai Shift (S)
SHIFT_RESULTS = {"Normal", "Normal（Correction of missed punch）"}


def classify() -> list[str]:
    """Kembalikan status S (Shift)."""
    return ["S"]