"""
Absensi Rekap — Streamlit App
Jalankan dengan: streamlit run app.py
"""

import re
import io
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import time, datetime
from database.db import init_db, save_periode, get_periodes, get_rekap, get_daily

# ──────────────────────────────────────────────────────────────
# Konfigurasi Halaman
# ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Absensi Rekap",
    page_icon="🗓️",
    layout="wide",
    initial_sidebar_state="collapsed",
)
init_db()
# ──────────────────────────────────────────────────────────────
# Custom CSS
# ──────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Halaman */
.main .block-container {
    padding: 2rem 3rem 4rem;
    max-width: 1300px;
}

/* Header */
.app-header {
    background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
    border-radius: 16px;
    padding: 2.5rem 3rem;
    margin-bottom: 2.5rem;
    position: relative;
    overflow: hidden;
}
.app-header::before {
    content: '';
    position: absolute;
    top: -50%;
    right: -10%;
    width: 400px;
    height: 400px;
    background: radial-gradient(circle, rgba(255,255,255,0.04) 0%, transparent 70%);
    border-radius: 50%;
}
.app-header h1 {
    color: #ffffff;
    font-size: 2.1rem;
    font-weight: 700;
    margin: 0 0 0.3rem 0;
    letter-spacing: -0.03em;
}
.app-header p {
    color: rgba(255,255,255,0.60);
    font-size: 0.95rem;
    margin: 0;
    font-weight: 400;
}
.badge {
    display: inline-block;
    background: rgba(255,255,255,0.12);
    color: #7dd3fc;
    padding: 0.2rem 0.7rem;
    border-radius: 20px;
    font-size: 0.78rem;
    font-family: 'DM Mono', monospace;
    margin-bottom: 1rem;
    letter-spacing: 0.05em;
}

/* Upload area */
.upload-section {
    background: #f8fafc;
    border: 2px dashed #cbd5e1;
    border-radius: 14px;
    padding: 2.5rem;
    text-align: center;
    transition: border-color 0.2s;
}
.upload-section:hover { border-color: #3b82f6; }

/* Metric cards */
.metric-row {
    display: flex;
    gap: 1rem;
    margin: 1.5rem 0;
}
.metric-card {
    flex: 1;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    font-weight: 600;
}
.metric-normal { background: #f0fdf4; border-left: 4px solid #22c55e; }
.metric-late   { background: #fffbeb; border-left: 4px solid #f59e0b; }
.metric-k      { background: #fef2f2; border-left: 4px solid #ef4444; }
.metric-total  { background: #eff6ff; border-left: 4px solid #3b82f6; }

.metric-card .label { font-size: 0.78rem; color: #64748b; font-weight: 500; text-transform: uppercase; letter-spacing: 0.07em; }
.metric-card .value { font-size: 2rem; font-weight: 700; margin-top: 0.2rem; font-family: 'DM Mono', monospace; }
.metric-normal .value { color: #16a34a; }
.metric-late   .value { color: #d97706; }
.metric-k      .value { color: #dc2626; }
.metric-total  .value { color: #2563eb; }

/* Rules pills */
.rules-pill {
    display: inline-block;
    background: #e0f2fe;
    color: #0369a1;
    padding: 0.15rem 0.6rem;
    border-radius: 999px;
    font-size: 0.78rem;
    font-family: 'DM Mono', monospace;
}

/* Status styling in table */
.status-normal { color: #16a34a; font-weight: 600; }
.status-late   { color: #d97706; font-weight: 600; }
.status-k      { color: #dc2626; font-weight: 700; }

/* Download button */
.stDownloadButton button {
    background: linear-gradient(135deg, #1e40af, #3b82f6) !important;
    color: white !important;
    border: none !important;
    padding: 0.6rem 1.8rem !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    letter-spacing: 0.02em !important;
    transition: all 0.2s !important;
}
.stDownloadButton button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(59,130,246,0.4) !important;
}

/* Divider */
.section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #1e293b;
    letter-spacing: -0.01em;
    margin: 0 0 1rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

/* Info box */
.info-box {
    background: #eff6ff;
    border-radius: 10px;
    padding: 1rem 1.3rem;
    font-size: 0.88rem;
    color: #1e40af;
    border: 1px solid #bfdbfe;
    line-height: 1.7;
}

/* Sidebar / expander */
.streamlit-expanderHeader { font-weight: 600 !important; }

/* Hide streamlit branding */
#MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# Logic Absensi (sama seperti absensi_rekap.py)
# ──────────────────────────────────────────────────────────────

SKIP_SHIFTS       = {"Rest", "Not scheduled", "--", ""}
K_THRESHOLD_MIN   = 120  # 2 jam

def parse_shift_start(shift_text):
    if not isinstance(shift_text, str):
        return None
    s = shift_text.strip()
    if s in SKIP_SHIFTS:
        return None
    m = re.search(r'(\d{1,2}):(\d{2})', s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None

def parse_time_to_minutes(val):
    if val is None:
        return None
    if isinstance(val, str):
        v = val.strip()
        if v.lower() in ("not punched", "--", ""):
            return None
        m = re.match(r'^(\d{1,2}):(\d{2})', v)
        if m:
            return int(m.group(1)) * 60 + int(m.group(2))
        return None
    if isinstance(val, time):
        return val.hour * 60 + val.minute
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.hour * 60 + val.minute
    if isinstance(val, pd.Timedelta):
        return (int(val.total_seconds()) % 86400) // 60
    if isinstance(val, float):
        if pd.isna(val):
            return None
        return round(val * 1440) % 1440
    return None

def classify(earliest_raw, shift_text, att_result):
    att_lower = str(att_result).strip().lower() if pd.notna(att_result) else ""
    shift_start = parse_shift_start(shift_text)
    shift_clean = str(shift_text).strip() if isinstance(shift_text, str) else ""
    if shift_clean in SKIP_SHIFTS or shift_start is None:
        return None
    earliest = parse_time_to_minutes(earliest_raw)
    if earliest is None:
        return "Normal" if att_lower.startswith("normal") else None
    diff = earliest - shift_start
    if diff <= 0:
        return "Normal"
    elif diff <= K_THRESHOLD_MIN:
        return "Late"
    else:
        return "K"

# ──────────────────────────────────────────────────────────────
# Helpers: Rincian Harian
# ──────────────────────────────────────────────────────────────

def parse_date_from_time(val):
    """Ekstrak tanggal dari format '2026/04/30 星期四' → '2026/04/30'."""
    if not isinstance(val, str):
        return str(val) if val else ""
    m = re.search(r'(\d{4}/\d{2}/\d{2})', val)
    return m.group(1) if m else val.strip()

def classify_shift_type(shift_text):
    """Klasifikasi shift: H (Rest), S2 (Malam/Night), S1 (lainnya), None (skip)."""
    if not isinstance(shift_text, str):
        return None
    s = shift_text.strip()
    if s == "Rest":
        return "H"
    if s in ("", "--", "Not scheduled"):
        return None
    s_lower = s.lower()
    # Night / malam keywords → S2
    if any(kw in s_lower for kw in ["s2", "night", "malam"]):
        return "S2"
    return "S1"

@st.cache_data(show_spinner=False)
def get_employee_daily(file_bytes, account):
    """Kembalikan (detail_df, summary_df) rincian harian satu karyawan."""
    buf = io.BytesIO(file_bytes)
    df_all = pd.read_excel(
        buf,
        sheet_name="General statistics and attendan",
        header=4,
        dtype={"Earliest": str, "Latest": str},
    )
    df_emp = df_all[df_all["Account"].astype(str).str.strip() == account].copy()

    def _parse_hours(val):
        if val is None:
            return 0.0
        if isinstance(val, (int, float)):
            return 0.0 if pd.isna(val) else float(val)
        try:
            return float(str(val).strip())
        except Exception:
            return 0.0

    df_emp["Tanggal"]   = df_emp["Time"].astype(str).apply(parse_date_from_time)
    df_emp["Tipe"]      = df_emp["Shift"].apply(classify_shift_type)
    df_emp["Jam Kerja"] = df_emp["Actual working hours(Hour)"].apply(_parse_hours)

    # Nilai yang dianggap "tidak absen sama sekali"
    _NOT_PUNCHED = {"not punched", "--", ""}

    rows = []
    for _, r in df_emp.iterrows():
        tipe = r["Tipe"]
        if tipe is None:
            continue

        # Override → H jika Earliest DAN Latest keduanya "Not punched"
        earliest_str = str(r["Earliest"]).strip().lower() if pd.notna(r["Earliest"]) else ""
        latest_str   = str(r["Latest"]).strip().lower()   if pd.notna(r["Latest"])   else ""
        if earliest_str in _NOT_PUNCHED and latest_str in _NOT_PUNCHED:
            tipe = "H"

        rows.append({
            "Tanggal"     : r["Tanggal"],
            "Shift"       : str(r["Shift"]).strip(),
            "Tipe"        : tipe,
            "Jam Masuk"   : str(r["Earliest"]).strip() if pd.notna(r["Earliest"]) else "--",
            "Jam Keluar"  : str(r["Latest"]).strip() if pd.notna(r["Latest"]) else "--",  # ← BARU
            "Status"      : str(r["Attendance results"]).strip() if pd.notna(r["Attendance results"]) else "--",
            "Jam Kerja"   : r["Jam Kerja"],
            "Klasifikasi" : classify(r["Earliest"], r["Shift"], r["Attendance results"]),
        })

    detail_df = pd.DataFrame(rows).sort_values("Tanggal").reset_index(drop=True)
    detail_df.insert(0, "No.", range(1, len(detail_df) + 1))

    summary_df = (
        detail_df.groupby("Tipe")
        .agg(Hari=("Tanggal", "count"), Total_Jam=("Jam Kerja", "sum"))
        .reset_index()
        .sort_values("Tipe")
    )
    return detail_df, summary_df


@st.dialog("📋 Rincian Harian Karyawan", width="large")
def show_daily_detail(account, nama, rules, file_bytes):
    """Modal dialog rincian harian per karyawan."""
    with st.spinner("Memuat rincian harian…"):
        detail_df, summary_df = get_employee_daily(file_bytes, account)

    # ── Info karyawan ──
    st.markdown(f"""
<div style="background:#f8fafc;border-radius:10px;padding:1rem 1.4rem;
            border-left:4px solid #3b82f6;margin-bottom:1.2rem;">
  <div style="font-size:1.05rem;font-weight:700;color:#1e293b;">{nama}</div>
  <div style="font-size:0.82rem;color:#64748b;margin-top:0.2rem;">
    <code>{account}</code> &nbsp;·&nbsp; {rules}
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Kartu Akumulasi S1 / S2 / H ──
    tipe_cfg = {
        "S1": ("🌅", "S1",  "#f0fdf4", "#22c55e", "#166534"),
        "S2": ("🌙", "S2", "#faf5ff", "#a855f7", "#6b21a8"),
        "H" : ("🏖️", "H",  "#fff7ed", "#fb923c", "#9a3412"),
    }
    cols = st.columns(3)
    for i, tipe in enumerate(["S1", "S2", "H"]):
        row = summary_df[summary_df["Tipe"] == tipe]
        hari = int(row["Hari"].values[0])      if len(row) else 0
        jam  = float(row["Total_Jam"].values[0]) if len(row) else 0.0
        icon, label, bg, border_c, text_c = tipe_cfg[tipe]
        with cols[i]:
            st.markdown(f"""
<div style="background:{bg};border-left:4px solid {border_c};border-radius:10px;
            padding:1rem 1.2rem;text-align:center;">
  <div style="font-size:0.75rem;color:#64748b;font-weight:500;text-transform:uppercase;
              letter-spacing:.06em;">{icon} {label}</div>
  <div style="font-size:1.9rem;font-weight:700;color:{text_c};font-family:'DM Mono',monospace;">
    {hari}<span style="font-size:1rem"> hari</span>
  </div>
  <div style="font-size:0.8rem;color:#64748b;margin-top:0.1rem;">
    {jam:.1f} jam kerja
  </div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:1.2rem'></div>", unsafe_allow_html=True)

    # ── Helper: durasi terlambat ──
    def _menit_terlambat(row):
        shift_start = parse_shift_start(row["Shift"])
        jam_masuk   = parse_time_to_minutes(row["Jam Masuk"])
        if shift_start is None or jam_masuk is None:
            return "--"
        diff = jam_masuk - shift_start
        if diff <= 0:
            return "--"
        h, m = divmod(diff, 60)
        return f"{h}j {m}m" if h > 0 else f"{m} mnt"

    KLAS_EMOJI = {"Normal": "✅ Normal", "Late": "🟡 Late", "K": "🔴 K"}

    late_df = detail_df[detail_df["Klasifikasi"] == "Late"].copy()
    k_df    = detail_df[detail_df["Klasifikasi"] == "K"].copy()
    n_late  = len(late_df)
    n_k     = len(k_df)

    # ── Expander 1: Rincian Keterlambatan ──
    late_label = f"⚠️ Rincian Keterlambatan  —  🟡 {n_late} Late  ·  🔴 {n_k} K"
    with st.expander(late_label, expanded=False):
        if n_late == 0 and n_k == 0:
            st.success("✅ Tidak ada keterlambatan pada periode ini.")
        else:
            mc1, mc2 = st.columns(2)
            with mc1:
                st.markdown(f"""
<div style="background:#fffbeb;border-left:4px solid #f59e0b;border-radius:10px;
            padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">
  <div style="font-size:0.72rem;color:#92400e;font-weight:600;text-transform:uppercase;
              letter-spacing:.06em;">🟡 Late</div>
  <div style="font-size:1.7rem;font-weight:700;color:#d97706;font-family:'DM Mono',monospace;">
    {n_late}<span style="font-size:.9rem"> hari</span>
  </div>
</div>""", unsafe_allow_html=True)
            with mc2:
                st.markdown(f"""
<div style="background:#fef2f2;border-left:4px solid #ef4444;border-radius:10px;
            padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">
  <div style="font-size:0.72rem;color:#991b1b;font-weight:600;text-transform:uppercase;
              letter-spacing:.06em;">🔴 K (&gt;2 jam)</div>
  <div style="font-size:1.7rem;font-weight:700;color:#dc2626;font-family:'DM Mono',monospace;">
    {n_k}<span style="font-size:.9rem"> hari</span>
  </div>
</div>""", unsafe_allow_html=True)

            late_k_combined = pd.concat([late_df, k_df]).sort_values("Tanggal").reset_index(drop=True)
            late_k_combined["No."]       = range(1, len(late_k_combined) + 1)
            late_k_combined["Terlambat"] = late_k_combined.apply(_menit_terlambat, axis=1)
            late_k_combined["Klasifikasi"] = late_k_combined["Klasifikasi"].map(
                lambda x: KLAS_EMOJI.get(x, x) if x else "—"
            )

            st.dataframe(
                late_k_combined[["No.", "Tanggal", "Klasifikasi", "Shift", "Jam Masuk", "Jam Keluar", "Terlambat"]],
                use_container_width=True,
                height=min(60 + len(late_k_combined) * 35, 380),
                hide_index=True,
                column_config={
                    "No."        : st.column_config.NumberColumn("No.", width="small"),
                    "Tanggal"    : st.column_config.TextColumn("Tanggal", width="medium"),
                    "Klasifikasi": st.column_config.TextColumn("Status", width="small"),
                    "Shift"      : st.column_config.TextColumn("Shift", width="large"),
                    "Jam Masuk"  : st.column_config.TextColumn("Jam Masuk", width="small"),
                    "Jam Keluar" : st.column_config.TextColumn("Jam Keluar", width="small"),  # ← BARU
                    "Terlambat"  : st.column_config.TextColumn("Terlambat", width="small"),
                },
            )

    # ── Expander 2: Rincian per Tipe Shift (S1 / S2 / H) ──
    tipe_counts = detail_df["Tipe"].value_counts()
    s1_n = tipe_counts.get("S1", 0)
    s2_n = tipe_counts.get("S2", 0)
    h_n  = tipe_counts.get("H",  0)
    shift_label = f"📅 Rincian per Tipe Shift  —  🌅 S1: {s1_n}  ·  🌙 S2: {s2_n}  ·  🏖️ H: {h_n}"
    with st.expander(shift_label, expanded=False):
        tab_s1, tab_s2, tab_h = st.tabs(["🌅 S1", "🌙 S2", "🏖️ H"])

        def _render_tipe_tab(tipe_key):
            df_tipe = detail_df[detail_df["Tipe"] == tipe_key].copy().reset_index(drop=True)
            if df_tipe.empty:
                st.info(f"Tidak ada data tipe {tipe_key} pada periode ini.")
                return
            df_tipe["No."]         = range(1, len(df_tipe) + 1)
            df_tipe["Klasifikasi"] = df_tipe["Klasifikasi"].map(
                lambda x: KLAS_EMOJI.get(x, "—") if x else "—"
            )
            df_tipe["Jam Kerja"]   = df_tipe["Jam Kerja"].apply(
                lambda x: f"{x:.1f} jam" if x > 0 else "—"
            )
            total_jam = detail_df[detail_df["Tipe"] == tipe_key]["Jam Kerja"].sum()
            st.caption(f"{len(df_tipe)} hari  ·  Total jam kerja: {total_jam:.1f} jam")
            st.dataframe(
                df_tipe[["No.", "Tanggal", "Shift", "Jam Masuk", "Jam Keluar", "Klasifikasi", "Jam Kerja"]],
                use_container_width=True,
                height=min(60 + len(df_tipe) * 35, 420),
                hide_index=True,
                column_config={
                    "No."        : st.column_config.NumberColumn("No.", width="small"),
                    "Tanggal"    : st.column_config.TextColumn("Tanggal", width="medium"),
                    "Shift"      : st.column_config.TextColumn("Shift", width="large"),
                    "Jam Masuk"  : st.column_config.TextColumn("Jam Masuk", width="small"),
                    "Jam Keluar" : st.column_config.TextColumn("Jam Keluar", width="small"),  # ← BARU
                    "Klasifikasi": st.column_config.TextColumn("Klasifikasi", width="small"),
                    "Jam Kerja"  : st.column_config.TextColumn("Jam Kerja", width="small"),
                },
            )

        with tab_s1: _render_tipe_tab("S1")
        with tab_s2: _render_tipe_tab("S2")
        with tab_h:  _render_tipe_tab("H")

    # ── Expander 3: Detail Lengkap per Hari ──
    with st.expander(f"🗂️ Detail Lengkap per Hari  —  {len(detail_df)} hari tercatat", expanded=False):
        TIPE_EMOJI = {"S1": "🌅 S1", "S2": "🌙 S2", "H": "🏖️ H"}
        detail_display = detail_df.copy()
        detail_display["Tipe"]        = detail_display["Tipe"].map(lambda x: TIPE_EMOJI.get(x, x))
        detail_display["Klasifikasi"] = detail_display["Klasifikasi"].map(
            lambda x: KLAS_EMOJI.get(x, "—") if x else "—"
        )
        detail_display["Jam Kerja"]   = detail_display["Jam Kerja"].apply(
            lambda x: f"{x:.1f} jam" if x > 0 else "—"
        )
        st.dataframe(
            detail_display[["No.", "Tanggal", "Tipe", "Shift", "Jam Masuk", "Jam Keluar", "Status", "Klasifikasi", "Jam Kerja"]],
            use_container_width=True,
            height=420,
            hide_index=True,
            column_config={
                "No."        : st.column_config.NumberColumn("No.", width="small"),
                "Tanggal"    : st.column_config.TextColumn("Tanggal", width="medium"),
                "Tipe"       : st.column_config.TextColumn("Tipe", width="small"),
                "Shift"      : st.column_config.TextColumn("Shift", width="large"),
                "Jam Masuk"  : st.column_config.TextColumn("Jam Masuk", width="small"),
                "Jam Keluar" : st.column_config.TextColumn("Jam Keluar", width="small"),  # ← BARU
                "Status"     : st.column_config.TextColumn("Status Absensi", width="large"),
                "Klasifikasi": st.column_config.TextColumn("Klasifikasi", width="small"),
                "Jam Kerja"  : st.column_config.TextColumn("Jam Kerja", width="small"),
            },
        )

    st.caption("Klik di luar kotak ini untuk menutup")


@st.cache_data(show_spinner=False)
def process_file(file_bytes):
    """Proses file Excel dan kembalikan DataFrame summary."""
    buf = io.BytesIO(file_bytes)
    df = pd.read_excel(
        buf,
        sheet_name="General statistics and attendan",
        header=4,
        dtype={"Earliest": str, "Latest": str}
    )

    required = ["Name", "Account", "Rules", "Shift", "Earliest", "Attendance results"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom tidak ditemukan: {missing}")

    df = df[required].dropna(subset=["Account", "Rules"])
    df = df[~df["Account"].astype(str).str.strip().isin(["", "--"])]

    df["Status"] = df.apply(
        lambda r: classify(r["Earliest"], r["Shift"], r["Attendance results"]), axis=1
    )

    # Rules diambil nilai terbanyak per Account (handle karyawan ganti Rules)
    all_employees = (
        df.groupby("Account")["Rules"]
        .agg(lambda x: x.mode()[0])
        .reset_index()
    )

    classified = df[df["Status"].notna()]
    pivot = classified.pivot_table(
        index="Account",
        columns="Status",
        values="Shift",
        aggfunc="count",
        fill_value=0
    ).reset_index()
    pivot.columns.name = None
    for col in ["Normal", "Late", "K"]:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot = all_employees.merge(pivot, on="Account", how="left").fillna(0)
    for col in ["Normal", "Late", "K"]:
        pivot[col] = pivot[col].astype(int)

    # Ambil Name lengkap dari baris pertama per Account
    name_map = df.groupby("Account")["Name"].first()
    pivot["Nama"] = pivot["Account"].map(name_map)

    pivot = pivot.sort_values(["Rules", "Nama"]).reset_index(drop=True)
    pivot.insert(0, "No.", range(1, len(pivot) + 1))
    result = pivot[["No.", "Nama", "Account", "Rules", "Normal", "Late", "K"]].copy()

    stats = {
        "total_rows": len(df),
        "classified": len(classified),
        "skipped": len(df) - len(classified),
        "employees": len(result),
        "dist": df["Status"].value_counts(dropna=False).to_dict()
    }
    return result, stats

def to_excel_bytes(df, time_range=""):
    """Konversi DataFrame ke bytes Excel dengan formatting."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap"

    DARK   = PatternFill("solid", fgColor="0F2027")
    BLUE   = PatternFill("solid", fgColor="2E75B6")
    ALT    = PatternFill("solid", fgColor="EFF6FF")
    WHITE  = PatternFill("solid", fgColor="FFFFFF")
    GREEN  = PatternFill("solid", fgColor="F0FDF4")
    YELLOW = PatternFill("solid", fgColor="FFFBEB")
    RED    = PatternFill("solid", fgColor="FEF2F2")

    thin   = Side(style="thin", color="CBD5E1")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center")
    LEFT   = Alignment(horizontal="left",   vertical="center")

    # Title
    ws.merge_cells("A1:G1")
    ws["A1"] = "Summary Attendance"
    ws["A1"].font = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    ws["A1"].fill = DARK
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 30

    # Subtitle
    ws.merge_cells("A2:G2")
    ws["A2"] = f"Time Range: {time_range}" if time_range else "Rekap Absensi"
    ws["A2"].font = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    ws["A2"].fill = BLUE
    ws["A2"].alignment = CENTER
    ws.row_dimensions[2].height = 16

    # Header
    headers = ["No.", "Nama", "Account", "Rules", "Normal", "Late", "K"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        c.fill  = BLUE
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[3].height = 20

    # Data
    for ri, row in df.iterrows():
        er = ri + 4
        base_fill = ALT if ri % 2 == 0 else WHITE
        vals = [row["No."], row["Nama"], row["Account"], row["Rules"], row["Normal"], row["Late"], row["K"]]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=er, column=ci, value=v)
            c.border = BORDER
            c.font = Font(name="Calibri", size=10)
            if ci <= 4:
                c.fill = base_fill
                c.alignment = LEFT if ci > 1 else CENTER
            elif ci == 5:  # Normal
                c.fill = GREEN if v > 0 else base_fill
                c.font = Font(name="Calibri", size=10, bold=(v > 0), color="166534" if v > 0 else "000000")
                c.alignment = CENTER
            elif ci == 6:  # Late
                c.fill = YELLOW if v > 0 else base_fill
                c.font = Font(name="Calibri", size=10, bold=(v > 0), color="92400E" if v > 0 else "000000")
                c.alignment = CENTER
            else:  # K
                c.fill = RED if v > 0 else base_fill
                c.font = Font(name="Calibri", size=10, bold=(v > 0), color="991B1B" if v > 0 else "000000")
                c.alignment = CENTER
        ws.row_dimensions[er].height = 17

    # Total row
    tr = len(df) + 4
    ws.merge_cells(f"A{tr}:D{tr}")
    ws[f"A{tr}"] = "TOTAL"
    ws[f"A{tr}"].font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    ws[f"A{tr}"].fill = DARK
    ws[f"A{tr}"].alignment = CENTER
    ws[f"A{tr}"].border = BORDER
    for ci, col in enumerate(["E", "F", "G"], 5):
        c = ws.cell(row=tr, column=ci)
        c.value = f"=SUM({get_column_letter(ci)}4:{get_column_letter(ci)}{tr-1})"
        c.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        c.fill  = DARK
        c.alignment = CENTER
        c.border = BORDER

    for ci, w in enumerate([6, 32, 24, 28, 10, 10, 8], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:G{tr-1}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────

# Header
st.markdown("""
<div class="app-header">
  <div class="badge">HR TOOLS</div>
  <h1>🗓️ Absensi Rekap Generator</h1>
  <p>Upload file Excel absensi → Hitung Normal / Late / K per karyawan → Download hasil</p>
</div>
""", unsafe_allow_html=True)

# ── Upload ──
col_up, col_info = st.columns([2, 1], gap="large")

with col_up:
    st.markdown('<p class="section-title">📂 Upload File Excel</p>', unsafe_allow_html=True)
    periodes_tersedia = get_periodes()
    
    if periodes_tersedia:
        st.markdown("**📅 Atau pilih periode yang sudah tersimpan:**")
        periode_dipilih = st.selectbox(
            label="",
            options=["— Upload file baru —"] + periodes_tersedia,
            label_visibility="collapsed"
        )
    else:
        periode_dipilih = "— Upload file baru —"

    uploaded = st.file_uploader(
        label="",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
        help="File Excel dari sistem absensi. Sheet 'General statistics and attendan' harus ada."
    )

with col_info:
    st.markdown("""
<div class="info-box">
<strong>Logic Klasifikasi:</strong><br>
🟢 <strong>Normal</strong> — datang tepat waktu atau lebih awal<br>
🟡 <strong>Late</strong> — terlambat ≤ 2 jam dari jam shift<br>
🔴 <strong>K</strong> — terlambat > 2 jam dari jam shift<br><br>
<strong>Catatan:</strong><br>
• K <em>tidak</em> double-count ke kolom Late<br>
• Not punched + Leave/Offsite → Normal<br>
• Shift Rest/Not scheduled → dilewati
</div>
""", unsafe_allow_html=True)

# ── Proses ──
if uploaded is not None or periode_dipilih != "— Upload file baru —":
    if uploaded is not None:
        file_bytes = uploaded.read()

        with st.spinner("Memproses data absensi…"):
            try:
                df_result, stats = process_file(file_bytes)
            except Exception as e:
                st.error(f"❌ Gagal memproses file: {e}")
                st.stop()

        # Simpan ke database
        try:
            import io as _io
            import re as _re
            import pandas as _pd
            _buf = _io.BytesIO(file_bytes)
            df_raw = _pd.read_excel(
                _buf,
                sheet_name="General statistics and attendan",
                header=4,
                dtype={"Earliest": str, "Latest": str},
            )
            # Ekstrak periode "YYYY-MM" dari baris Time pertama yang valid
            _periode = None
            for _val in df_raw["Time"].astype(str):
                _m = _re.search(r'(\d{4})/(\d{2})/(\d{2})', _val)
                if _m:
                    _periode = f"{_m.group(1)}-{_m.group(2)}"
                    break
            if _periode is None:
                _periode = "unknown"

            # Samakan filter dengan process_file() agar data DB konsisten
            df_raw = df_raw.dropna(subset=["Account", "Rules"])
            df_raw = df_raw[~df_raw["Account"].astype(str).str.strip().isin(["", "--"])]

            df_raw["_tipe_shift"] = df_raw["Shift"].apply(classify_shift_type)
            df_raw["_status_klasifikasi"] = df_raw.apply(
                lambda r: classify(r["Earliest"], r["Shift"], r["Attendance results"]), axis=1
            )
            save_periode(df_raw, _periode)
        except Exception as e:
            st.warning(f"⚠️ Gagal simpan ke database: {e}")
    else:
        # Ambil dari database
        df_result = get_rekap(periode_dipilih)
        df_result = df_result.rename(columns={
            "nama": "Nama", "account": "Account", "rules": "Rules",
            "normal": "Normal", "late": "Late", "k": "K",
        })
        file_bytes = None
        stats = {
            "total_rows": len(df_result),
            "classified": int(df_result[["Normal", "Late", "K"]].sum().sum()),
            "skipped": 0,
            "employees": len(df_result),
            "dist": {}
        }
        df_result.insert(0, "No.", range(1, len(df_result) + 1))

    # ── Metric Cards ──
    total_n = int(df_result["Normal"].sum())
    total_l = int(df_result["Late"].sum())
    total_k = int(df_result["K"].sum())
    total_e = stats["employees"]

    st.markdown(f"""
<div class="metric-row">
  <div class="metric-card metric-normal">
    <div class="label">Total Normal</div>
    <div class="value">{total_n:,}</div>
  </div>
  <div class="metric-card metric-late">
    <div class="label">Total Late</div>
    <div class="value">{total_l:,}</div>
  </div>
  <div class="metric-card metric-k">
    <div class="label">Total K (&gt;2 jam)</div>
    <div class="value">{total_k:,}</div>
  </div>
  <div class="metric-card metric-total">
    <div class="label">Jumlah Karyawan</div>
    <div class="value">{total_e:,}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Filter & Tabel ──
    st.markdown('<p class="section-title">📊 Hasil Summary per Karyawan</p>', unsafe_allow_html=True)

    # Filter
    fcol1, fcol2, fcol3 = st.columns([2, 2, 1])
    with fcol1:
        all_rules = sorted(df_result["Rules"].unique().tolist())
        sel_rules = st.multiselect("Filter Rules", options=all_rules, placeholder="Semua Rules")
    with fcol2:
        search = st.text_input("Cari Nama / Account", placeholder="Ketik nama atau account…")
    with fcol3:
        show_late_only = st.checkbox("Hanya Late/K", value=False)

    # Apply filter
    df_show = df_result.copy()
    if sel_rules:
        df_show = df_show[df_show["Rules"].isin(sel_rules)]
    if search:
        mask = (
            df_show["Nama"].str.contains(search, case=False, na=False) |
            df_show["Account"].str.contains(search, case=False, na=False)
        )
        df_show = df_show[mask]
    if show_late_only:
        df_show = df_show[(df_show["Late"] > 0) | (df_show["K"] > 0)]

    # Re-number setelah filter
    df_show = df_show.copy()
    df_show["No."] = range(1, len(df_show) + 1)

    st.caption("💡 **Klik baris karyawan** untuk melihat rincian harian & tanggal keterlambatan")
    sel_event = st.dataframe(
        df_show,
        use_container_width=True,
        height=520,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        column_config={
            "No."     : st.column_config.NumberColumn("No.", width="small"),
            "Nama"    : st.column_config.TextColumn("Nama", width="large"),
            "Account" : st.column_config.TextColumn("Account", width="medium"),
            "Rules"   : st.column_config.TextColumn("Rules", width="medium"),
            "Normal"  : st.column_config.NumberColumn("Normal ✅", format="%d", width="small"),
            "Late"    : st.column_config.NumberColumn("Late 🟡", format="%d", width="small"),
            "K"       : st.column_config.NumberColumn("K 🔴", format="%d", width="small"),
        },
    )

    # ── Trigger dialog rincian harian ──
    sel_rows = sel_event.selection.rows if sel_event and sel_event.selection else []
    if sel_rows and file_bytes is not None:
        idx = sel_rows[0]
        if idx < len(df_show):
            emp = df_show.iloc[idx]
            show_daily_detail(
                account=emp["Account"],
                nama=emp["Nama"],
                rules=emp["Rules"],
                file_bytes=file_bytes,
            )
    elif sel_rows and file_bytes is None:
        st.info("ℹ️ Rincian harian hanya tersedia saat file Excel diupload langsung.")

    st.caption(f"Menampilkan {len(df_show):,} dari {len(df_result):,} karyawan  •  "
               f"Total baris diproses: {stats['total_rows']:,}  •  "
               f"Diklasifikasikan: {stats['classified']:,}  •  "
               f"Dilewati (Rest/Absence/dll): {stats['skipped']:,}")

    # ── Download ──
    st.markdown("---")
    dcol1, dcol2, dcol3 = st.columns([1, 1, 2])

    # Ambil time range dari file
    try:
        buf = io.BytesIO(file_bytes)
        raw = pd.read_excel(buf, sheet_name="General statistics and attendan",
                            header=None, nrows=2)
        tr_text = str(raw.iloc[1, 0])
        m = re.search(r'Time Range[:\s]*([\d/–\-\s]+)', tr_text)
        time_range = m.group(1).strip() if m else ""
    except Exception:
        time_range = ""

    with dcol1:
        xlsx_bytes = to_excel_bytes(df_result, time_range)
        fname = f"Rekap_Absensi_{time_range.replace(' ', '_').replace('–','s/d')}.xlsx" \
                if time_range else "Rekap_Absensi.xlsx"
        st.download_button(
            label="⬇️  Download Rekap Lengkap (.xlsx)",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with dcol2:
        if len(df_show) < len(df_result):
            xlsx_filtered = to_excel_bytes(df_show, time_range)
            st.download_button(
                label="⬇️  Download Hasil Filter (.xlsx)",
                data=xlsx_filtered,
                file_name=f"Rekap_Filter_{fname}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # ── Detail per Rules (expandable) ──
    with st.expander("📈 Ringkasan per Rules"):
        grp = df_result.groupby("Rules").agg(
            Karyawan=("Account", "count"),
            Normal=("Normal", "sum"),
            Late=("Late", "sum"),
            K=("K", "sum")
        ).reset_index().sort_values("K", ascending=False)
        grp["Late Rate"] = (grp["Late"] / (grp["Normal"] + grp["Late"] + grp["K"]) * 100).round(1)
        grp["K Rate"]    = (grp["K"]    / (grp["Normal"] + grp["Late"] + grp["K"]) * 100).round(1)
        st.dataframe(
            grp,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Rules"    : st.column_config.TextColumn("Rules"),
                "Karyawan" : st.column_config.NumberColumn("Karyawan", format="%d"),
                "Normal"   : st.column_config.NumberColumn("Normal ✅", format="%d"),
                "Late"     : st.column_config.NumberColumn("Late 🟡", format="%d"),
                "K"        : st.column_config.NumberColumn("K 🔴", format="%d"),
                "Late Rate": st.column_config.NumberColumn("% Late", format="%.1f%%"),
                "K Rate"   : st.column_config.NumberColumn("% K", format="%.1f%%"),
            }
        )

else:
    # Placeholder saat belum upload
    st.markdown("""
<div style="text-align:center; padding: 3rem 2rem; color: #94a3b8;">
    <div style="font-size: 3.5rem; margin-bottom: 1rem;">📁</div>
    <div style="font-size: 1.1rem; font-weight: 600; color: #64748b; margin-bottom: 0.5rem;">
        Belum ada file yang diupload
    </div>
    <div style="font-size: 0.9rem;">
        Upload file <code>.xlsx</code> absensi di atas untuk memulai
    </div>
</div>
""", unsafe_allow_html=True)