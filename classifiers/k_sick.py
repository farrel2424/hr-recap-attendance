# classifiers/k_sick.py
"""
Klasifikasi K — sakit dengan surat (K-Sick W Letter).

Dipanggil oleh __init__.classify() ketika leave_app mengandung "K-Sick W Letter".

Aturan dual-count:
  - Jika att_result mengandung "Normal" → ["Normal", "K"]
  - Jika att_result tidak mengandung "Normal" → ["K"]

Return: list satu atau dua elemen.
"""


def classify(has_normal: bool) -> list[str]:
    """
    Args:
        has_normal : True jika att_result mengandung kata "normal"

    Returns:
        ["Normal", "K"] atau ["K"]
    """
    if has_normal:
        return ["Normal", "K"]
    return ["K"]