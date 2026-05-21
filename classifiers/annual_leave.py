# classifiers/annual_leave.py
"""
Klasifikasi Annual Leave (AL / ½AL).

Aturan (berbasis kolom):
  - Kolom "AnnualLeave - 印尼员工年假(Day(s))" bernilai 0.5 → "1/2 AL"
  - Kolom "AnnualLeave - 印尼员工年假(Day(s))" bernilai 1   → "AL"

Tidak ada dual-count dengan S — status bersifat standalone.

Return:
  ["AL"]     — nilai kolom ≥ 1
  ["1/2 AL"] — nilai kolom = 0.5
  None       — kolom kosong / 0 / "--"

Dipanggil oleh __init__.classify() setelah cek DW.
"""

from .base import parse_day_value


def classify(al_count) -> list[str] | None:
    """
    Args:
        al_count : nilai kolom "AnnualLeave - 印尼员工年假(Day(s))"

    Returns:
        ["AL"], ["1/2 AL"], atau None
    """
    val = parse_day_value(al_count)
    if val is None:
        return None
    # 0.5 → setengah hari cuti
    if abs(val - 0.5) < 0.01:
        return ["1/2 AL"]
    # 1 atau lebih → cuti penuh
    if val >= 0.99:
        return ["AL"]
    return None