# classifiers/k_sick.py
"""
Klasifikasi K — sakit dengan surat (K-Sick W Letter).

Dipanggil oleh __init__.classify() ketika kolom "K-Sick W Letter"
(K-Sick W Letter-病假有信) bernilai BUKAN "0", "--", atau kosong.

Output selalu bersifat standalone — tidak ada dual-count dengan S.

Return: ["K"]
"""


def classify() -> list[str]:
    """Kembalikan status K (Sakit dengan Surat)."""
    return ["K"]