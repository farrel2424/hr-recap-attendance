# classifiers/off.py
"""
Klasifikasi Off — hari libur / tidak terjadwal yang tercatat Normal.

Dipanggil oleh __init__.classify() ketika att_result bernilai:
  - "Normal (rest)"
  - "Normal (not scheduled)"

Return: ["Off"]
"""

# Nilai att_result yang diklasifikasikan sebagai Off
OFF_RESULTS = {"Normal (rest)", "Normal (not scheduled)"}


def classify() -> list[str]:
    """Kembalikan status Off."""
    return ["Off"]