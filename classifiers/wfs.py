# classifiers/wfs.py
"""
Klasifikasi WFS — Work From Offsite.

Dipanggil oleh __init__.classify() ketika:
  - att_result bernilai TEPAT "Normal (Offsite)"  DAN
  - kolom "Offsite(Hour)" bernilai bukan "--", kosong, atau NaN

Return: ["WFS"]
"""


def classify() -> list[str]:
    """Kembalikan status WFS (Work From Offsite)."""
    return ["WFS"]


    