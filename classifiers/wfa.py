# classifiers/wfa.py
"""
Klasifikasi WFA / ½ WFA — Work From Anywhere (Work From Home).

Aturan (berbasis kolom):
  - Kolom "WFH-WorkFromHome-家办公(Day(s))" bernilai 1   → "WFA"     (WFH penuh)
  - Kolom "WFH-WorkFromHome-家办公(Day(s))" bernilai 0.5 → "1/2 WFA" (WFH setengah hari)

Output bersifat standalone — tidak ada dual-count dengan status lain.

Return:
  ["WFA"]     — nilai kolom ≥ 1
  ["1/2 WFA"] — nilai kolom = 0.5
  None        — kolom kosong / 0 / "--"

Dipanggil oleh __init__.classify() setelah cek UL.
"""

from .base import parse_day_value


def classify(wfh_count) -> list[str] | None:
    """
    Args:
        wfh_count : nilai kolom "WFH-WorkFromHome-家办公(Day(s))"

    Returns:
        ["WFA"], ["1/2 WFA"], atau None
    """
    val = parse_day_value(wfh_count)
    if val is None:
        return None
    # nilai = 0.5 → WFH setengah hari
    if abs(val - 0.5) < 0.01:
        return ["1/2 WFA"]
    # nilai ≥ 1 → WFH penuh
    if val >= 0.99:
        return ["WFA"]
    return None
