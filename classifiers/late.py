# classifiers/late.py
"""
Klasifikasi keterlambatan / pulang lebih awal.

Logika baru (berbasis Punch Out vs jam selesai shift):
  ½UL  — hanya ada satu punch (in tanpa out, atau out tanpa in)
  ½UL  — punch out lebih cepat > K_THRESHOLD_MIN (120 menit) dari shift end
  Late — punch out lebih cepat ≤ K_THRESHOLD_MIN dari shift end
  Normal — punch out ≥ shift end (pulang tepat waktu atau lebih lama)
  None — tidak ada punch sama sekali (tidak bisa diklasifikasi)

Dipanggil oleh __init__.classify() ketika att_result TIDAK mengandung "normal".
"""

from .base import parse_time_to_minutes, has_punch, K_THRESHOLD_MIN


def _normalize_overnight(shift_start: int, shift_end: int, latest_min: int) -> tuple[int, int]:
    """
    Normalisasi untuk shift overnight agar shift_end dan latest_min
    bisa dibandingkan secara linear.

    Contoh: shift 19:00-05:00, punch out 04:30
      shift_end < shift_start → overnight
      latest_min (270) < shift_start (1140) → punch out setelah tengah malam
      → latest_min += 1440 = 1710, shift_end += 1440 = 1740
      diff = 1740 - 1710 = 30 mnt → Late ✓
    """
    if shift_end < shift_start:          # overnight shift
        shift_end += 1440
        if latest_min < shift_start:     # punch out ada di hari berikutnya
            latest_min += 1440
    return shift_end, latest_min


def classify(earliest_raw, latest_raw, shift_end: int, shift_start: int = 0) -> list[str] | None:
    """
    Args:
        earliest_raw : nilai mentah jam masuk
        latest_raw   : nilai mentah jam keluar
        shift_end    : jam selesai shift dalam menit
        shift_start  : jam mulai shift dalam menit (untuk normalisasi overnight)

    Returns:
        list satu elemen, atau None jika tidak ada punch sama sekali.
    """
    has_in  = has_punch(earliest_raw)
    has_out = has_punch(latest_raw)

    # Hanya satu punch (in tanpa out, atau out tanpa in) → ½UL
    if has_in ^ has_out:
        return ["1/2 UL"]

    # Tidak ada punch sama sekali → tidak bisa diklasifikasi
    if not has_in and not has_out:
        return None

    # Ada kedua punch → bandingkan punch-out dengan shift end
    latest = parse_time_to_minutes(latest_raw)
    if latest is None:
        return None

    # Normalisasi untuk shift overnight
    norm_end, norm_latest = _normalize_overnight(shift_start, shift_end, latest)

    diff = norm_end - norm_latest   # positif = pulang lebih awal dari jadwal
    if diff <= 0:
        return ["Normal"]
    elif diff <= K_THRESHOLD_MIN:
        return ["Late"]
    else:
        return ["1/2 UL"]