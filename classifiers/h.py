# classifiers/h.py
"""
Klasifikasi H — Hari Libur Nasional (Tanggal Merah).

Tidak dihasilkan oleh engine otomatis — hanya diterapkan melalui
fitur Bulk Correction (Tanggal Merah) di UI.

Return: ["H"]
"""


def classify() -> list[str]:
    """Kembalikan status H (Hari Libur Nasional)."""
    return ["H"]