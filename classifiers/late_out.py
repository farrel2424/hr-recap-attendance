# classifiers/late_out.py
"""
Klasifikasi keterlambatan Punch Out (Latest) vs jam selesai shift.

Dipanggil oleh __init__.classify() SETELAH late_in tidak menemukan keterlambatan
(artinya punch in sudah tepat waktu / lebih awal).

Aturan (menggunakan K_THRESHOLD_MIN = 120 menit):
  Punch out lebih cepat  1–120 menit dari shift end  → ["Late"]
  Punch out lebih cepat  > 120 menit dari shift end  → ["1/2 UL"]
  Punch out tepat waktu / lebih lama / tidak ada punch → None

Catatan:
  - Tidak berlaku jika row sudah diklasifikasi sebelumnya (Off / DW / K / AL / WFA / Late_in).
  - Output bersifat standalone — tidak ada dual-count dengan Normal.
  - Mendukung shift overnight (mis. 19:00–05:00).

Return: list satu elemen atau None
"""

from .base import parse_time_to_minutes, has_punch, K_THRESHOLD_MIN


def _normalize_overnight(shift_start: int, shift_end: int, latest_min: int) -> tuple[int, int]:
    """
    Normalisasi untuk shift overnight agar shift_end dan latest_min
    bisa dibandingkan secara linear.

    Contoh: shift 19:00–05:00, punch out 04:30
      shift_end (300) < shift_start (1140) → overnight
      latest_min (270) < shift_start (1140) → punch out setelah tengah malam
      → latest_min += 1440 = 1710, shift_end += 1440 = 1740
      diff = 1740 - 1710 = 30 mnt → Late ✓
    """
    if shift_end < shift_start:          # overnight shift
        shift_end += 1440
        if latest_min < shift_start:     # punch out ada di hari berikutnya
            latest_min += 1440
    return shift_end, latest_min


def classify(latest_raw, shift_end: int, shift_start: int = 0) -> list[str] | None:
    """
    Args:
        latest_raw  : nilai mentah jam keluar (kolom Latest)
        shift_end   : jam selesai shift dalam menit sejak tengah malam
                      (hasil parse_shift_end)
        shift_start : jam mulai shift dalam menit (untuk normalisasi overnight)

    Returns:
        ["Late"]    — pulang lebih awal 1–120 menit dari shift end
        ["1/2 UL"]  — pulang lebih awal > 120 menit dari shift end
        None        — tepat waktu / lebih lama / tidak ada punch
    """
    if not has_punch(latest_raw):
        return None

    latest = parse_time_to_minutes(latest_raw)
    if latest is None:
        return None

    norm_end, norm_latest = _normalize_overnight(shift_start, shift_end, latest)

    diff = norm_end - norm_latest   # positif = pulang lebih awal dari jadwal
    if diff <= 0:
        return None                 # tepat waktu atau lebih lama
    elif diff <= K_THRESHOLD_MIN:
        return ["Late"]
    else:
        return ["1/2 UL"]