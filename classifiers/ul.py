# classifiers/ul.py
"""
Klasifikasi UL — Unpaid Leave (cuti tidak dibayar).

Aturan (berbasis kolom):
  - Kolom "UL-Unpaid Leave-事假(Day(s))" bernilai 1   → "UL"   (cuti tidak dibayar penuh)
  - Kolom "UL-Unpaid Leave-事假(Day(s))" bernilai 0.5 → "1/2 UL" (setengah hari)

Output bersifat standalone — tidak ada dual-count dengan status lain.

Return:
  ["UL"]     — nilai kolom ≥ 1
  ["1/2 UL"] — nilai kolom = 0.5
  None       — kolom kosong / 0 / "--"

Dipanggil oleh __init__.classify() setelah cek AL.
"""

from .base import parse_day_value


def classify(ul_count) -> list[str] | None:
    """
    Args:
        ul_count : nilai kolom "UL-Unpaid Leave-事假(Day(s))"

    Returns:
        ["UL"], ["1/2 UL"], atau None
    """
    val = parse_day_value(ul_count)
    if val is None:
        return None
    # nilai ≥ 1 → UL penuh
    if val >= 0.99:
        return ["UL"]
    # nilai = 0.5 → setengah hari
    if abs(val - 0.5) < 0.01:
        return ["1/2 UL"]
    return None