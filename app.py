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
from database.db import init_db, save_periode, get_periodes, get_rekap, get_daily
from classifiers import (
    classify,
    classify_str,
    classify_shift_type,
    parse_shift_start,
    parse_shift_end,
    parse_time_to_minutes,
    has_status,
    SKIP_SHIFTS,
    _NOT_PUNCHED,
)

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
if "dialog_target" not in st.session_state:
    st.session_state.dialog_target = None
if "dialog_emp" not in st.session_state:
    st.session_state.dialog_emp = None

# ──────────────────────────────────────────────────────────────
# Custom CSS
# ──────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.main .block-container { padding: 2rem 3rem 4rem; max-width: 1400px; }

/* ── Header ── */
.app-header {
    background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
    border-radius: 16px; padding: 2.5rem 3rem; margin-bottom: 2.5rem;
    position: relative; overflow: hidden;
}
.app-header::before {
    content: ''; position: absolute; top: -50%; right: -10%;
    width: 400px; height: 400px;
    background: radial-gradient(circle, rgba(255,255,255,0.04) 0%, transparent 70%);
    border-radius: 50%;
}
.app-header h1 {
    color: #ffffff; font-size: 2.1rem; font-weight: 700;
    margin: 0 0 0.3rem 0; letter-spacing: -0.03em;
    display: flex; align-items: center; gap: 0.6rem;
}
.app-header p { color: rgba(255,255,255,0.60); font-size: 0.95rem; margin: 0; }
.badge {
    display: inline-flex; align-items: center; gap: 0.3rem;
    background: rgba(255,255,255,0.12); color: #7dd3fc;
    padding: 0.2rem 0.7rem; border-radius: 20px; font-size: 0.78rem;
    font-family: 'DM Mono', monospace; margin-bottom: 1rem; letter-spacing: 0.05em;
}

/* ── Metric Cards ── */
.metric-row {
    display: grid; grid-template-columns: repeat(5, 1fr);
    gap: 1.25rem; margin: 2rem 0 2.5rem;
}
.metric-card {
    border-radius: 14px; padding: 1.5rem 1.75rem 1.4rem;
    font-weight: 600; position: relative; overflow: hidden;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(0,0,0,0.08);
}
.metric-card::after {
    content: ''; position: absolute; bottom: -18px; right: -18px;
    width: 70px; height: 70px; border-radius: 50%; opacity: 0.07; background: currentColor;
}

.metric-normal  { background: #f0fdf4; border-left: 4px solid #22c55e; }
.metric-late    { background: #fffbeb; border-left: 4px solid #f59e0b; }
.metric-k       { background: #fef2f2; border-left: 4px solid #ef4444; }
.metric-total   { background: #eff6ff; border-left: 4px solid #3b82f6; }
.metric-al      { background: #fdf4ff; border-left: 4px solid #a855f7; }
.metric-half-al { background: #fff1f2; border-left: 4px solid #fb7185; }
.metric-wfa     { background: #f0f9ff; border-left: 4px solid #0ea5e9; }
.metric-dw      { background: #fff7ed; border-left: 4px solid #f97316; }
.metric-ksick   { background: #fdf2f8; border-left: 4px solid #ec4899; }
.metric-off     { background: #f8fafc; border-left: 4px solid #94a3b8; }

.metric-card .icon {
    font-size: 1.4rem; margin-bottom: 0.4rem; display: block; line-height: 1;
}
.metric-card .label {
    font-size: 0.75rem; color: #64748b; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 0.35rem;
}
.metric-card .value {
    font-size: 2.2rem; font-weight: 700;
    font-family: 'DM Mono', monospace; line-height: 1;
}
.metric-normal  .value { color: #16a34a; }
.metric-late    .value { color: #d97706; }
.metric-k       .value { color: #dc2626; }
.metric-total   .value { color: #2563eb; }
.metric-al      .value { color: #9333ea; }
.metric-half-al .value { color: #e11d48; }
.metric-wfa     .value { color: #0284c7; }
.metric-dw      .value { color: #ea580c; }
.metric-ksick   .value { color: #db2777; }
.metric-off     .value { color: #64748b; }
.metric-card .sub { font-size: 0.72rem; color: #94a3b8; font-weight: 400; margin-top: 0.35rem; }

/* ── Download Button ── */
.stDownloadButton button {
    background: linear-gradient(135deg, #1e40af, #3b82f6) !important;
    color: white !important; border: none !important;
    padding: 0.6rem 1.8rem !important; border-radius: 8px !important;
    font-weight: 600 !important; font-size: 0.9rem !important;
    letter-spacing: 0.02em !important; transition: all 0.2s !important;
}
.stDownloadButton button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(59,130,246,0.4) !important;
}

/* ── Section titles ── */
.section-title {
    font-size: 1rem; font-weight: 700; color: #1e293b;
    letter-spacing: -0.01em; margin: 0 0 1rem 0;
    display: flex; align-items: center; gap: 0.5rem;
}

/* ── Info chips ── */
.chip {
    display: inline-flex; align-items: center; gap: 0.3rem;
    background: #f1f5f9; color: #475569;
    padding: 0.2rem 0.65rem; border-radius: 20px;
    font-size: 0.76rem; font-weight: 500; margin-right: 0.35rem;
}
.chip-blue  { background: #eff6ff; color: #2563eb; }
.chip-green { background: #f0fdf4; color: #16a34a; }
.chip-amber { background: #fffbeb; color: #d97706; }
.chip-red   { background: #fef2f2; color: #dc2626; }

/* ── Status row separator ── */
.status-divider {
    display: flex; align-items: center; gap: 0.5rem;
    font-size: 0.78rem; color: #94a3b8; font-weight: 500;
    margin: 0.5rem 0 1rem; text-transform: uppercase; letter-spacing: 0.06em;
}
.status-divider::after {
    content: ''; flex: 1; height: 1px; background: #e2e8f0;
}

.streamlit-expanderHeader { font-weight: 600 !important; }
#MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# Helpers: Rincian Harian
# ──────────────────────────────────────────────────────────────

def parse_date_from_time(val):
    if not isinstance(val, str):
        return str(val) if val else ""
    m = re.search(r'(\d{4}/\d{2}/\d{2})', val)
    return m.group(1) if m else val.strip()


@st.cache_data(show_spinner=False)
def get_employee_daily(file_bytes, account):
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

    rows = []
    for _, r in df_emp.iterrows():
        tipe = r["Tipe"]
        if tipe is None:
            continue

        earliest_str = str(r["Earliest"]).strip().lower() if pd.notna(r["Earliest"]) else ""
        latest_str   = str(r["Latest"]).strip().lower()   if pd.notna(r["Latest"])   else ""
        if earliest_str in _NOT_PUNCHED and latest_str in _NOT_PUNCHED:
            tipe = "H"

        _klas_raw = classify(
            r["Earliest"], r["Shift"], r["Attendance results"],
            latest_raw=r["Latest"],
            leave_app=r.get("Leave & Overtime Application"),
        )
        _klas_display = " / ".join(_klas_raw) if _klas_raw else None

        rows.append({
            "Tanggal"        : r["Tanggal"],
            "Shift"          : str(r["Shift"]).strip(),
            "Tipe"           : tipe,
            "Jam Masuk"      : str(r["Earliest"]).strip() if pd.notna(r["Earliest"]) else "--",
            "Jam Keluar"     : str(r["Latest"]).strip()   if pd.notna(r["Latest"])   else "--",
            "Status"         : str(r["Attendance results"]).strip() if pd.notna(r["Attendance results"]) else "--",
            "Jam Kerja"      : r["Jam Kerja"],
            "Klasifikasi"    : _klas_display,
            "Klasifikasi_raw": _klas_raw,
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
    with st.spinner("Memuat rincian harian..."):
        detail_df, summary_df = get_employee_daily(file_bytes, account)

    st.markdown(
        '<div style="background:#f8fafc;border-radius:10px;padding:1rem 1.4rem;'
        'border-left:4px solid #3b82f6;margin-bottom:1.2rem;">'
        f'<div style="font-size:1.05rem;font-weight:700;color:#1e293b;">👤 {nama}</div>'
        f'<div style="font-size:0.82rem;color:#64748b;margin-top:0.2rem;">'
        f'<code>{account}</code> &nbsp;·&nbsp; 📋 {rules}</div></div>',
        unsafe_allow_html=True,
    )

    tipe_cfg = {
        "S1": ("☀️ S1 — Shift Pagi",  "#f0fdf4", "#22c55e", "#166534"),
        "S2": ("🌙 S2 — Shift Malam", "#faf5ff", "#a855f7", "#6b21a8"),
        "H" : ("🏖️ H — Libur/Rest",   "#fff7ed", "#fb923c", "#9a3412"),
    }
    cols = st.columns(3)
    for i, tipe in enumerate(["S1", "S2", "H"]):
        row = summary_df[summary_df["Tipe"] == tipe]
        hari = int(row["Hari"].values[0])        if len(row) else 0
        jam  = float(row["Total_Jam"].values[0]) if len(row) else 0.0
        label, bg, border_c, text_c = tipe_cfg[tipe]
        with cols[i]:
            st.markdown(
                f'<div style="background:{bg};border-left:4px solid {border_c};'
                f'border-radius:10px;padding:1rem 1.2rem;text-align:center;">'
                f'<div style="font-size:0.75rem;color:#64748b;font-weight:500;'
                f'text-transform:uppercase;letter-spacing:.06em;">{label}</div>'
                f'<div style="font-size:1.9rem;font-weight:700;color:{text_c};">'
                f'{hari}<span style="font-size:1rem"> hari</span></div>'
                f'<div style="font-size:0.8rem;color:#64748b;margin-top:0.1rem;">'
                f'⏱ {jam:.1f} jam kerja</div></div>',
                unsafe_allow_html=True,
            )

    st.markdown("<div style='margin-top:1.2rem'></div>", unsafe_allow_html=True)

    def _menit_lebih_awal(row):
        s_end      = parse_shift_end(row["Shift"])
        s_start    = parse_shift_start(row["Shift"])
        jam_keluar = parse_time_to_minutes(row["Jam Keluar"])
        if s_end is None or jam_keluar is None:
            return "--"
        if s_start is not None and s_end < s_start:
            s_end += 1440
            if jam_keluar < s_start:
                jam_keluar += 1440
        diff = s_end - jam_keluar
        if diff <= 0:
            return "--"
        h, m = divmod(diff, 60)
        return f"{h}j {m}m" if h > 0 else f"{m} mnt"

    late_df   = detail_df[detail_df["Klasifikasi_raw"].apply(lambda x: has_status(x, "Late"))].copy()
    k_df      = detail_df[detail_df["Klasifikasi_raw"].apply(lambda x: has_status(x, "1/2 UL"))].copy()
    al_df     = detail_df[detail_df["Klasifikasi_raw"].apply(
        lambda x: has_status(x, "AL") or has_status(x, "1/2 AL"))].copy()
    wfa_df    = detail_df[detail_df["Klasifikasi_raw"].apply(lambda x: has_status(x, "WFA"))].copy()
    dw_df     = detail_df[detail_df["Klasifikasi_raw"].apply(lambda x: has_status(x, "DW"))].copy()
    k_sick_df = detail_df[detail_df["Klasifikasi_raw"].apply(lambda x: has_status(x, "K"))].copy()

    n_late   = len(late_df)
    n_k      = len(k_df)
    n_al     = len(al_df)
    n_wfa    = len(wfa_df)
    n_dw     = len(dw_df)
    n_k_sick = len(k_sick_df)

    # Expander 1: Pelanggaran Jam Kerja
    with st.expander(
        f"⚠️ Pelanggaran Jam Kerja  —  Late: {n_late}  |  1/2 UL: {n_k}",
        expanded=False,
    ):
        if n_late == 0 and n_k == 0:
            st.success("✅ Tidak ada pelanggaran jam kerja pada periode ini.")
        else:
            mc1, mc2 = st.columns(2)
            with mc1:
                st.markdown(
                    f'<div style="background:#fffbeb;border-left:4px solid #f59e0b;border-radius:10px;'
                    f'padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">'
                    f'<div style="font-size:1.3rem;margin-bottom:.2rem;">🕐</div>'
                    f'<div style="font-size:0.72rem;color:#92400e;font-weight:600;text-transform:uppercase;">'
                    f'Late (terlambat masuk 1-120 mnt)</div>'
                    f'<div style="font-size:1.7rem;font-weight:700;color:#d97706;">'
                    f'{n_late}<span style="font-size:.9rem"> hari</span></div></div>',
                    unsafe_allow_html=True,
                )
            with mc2:
                st.markdown(
                    f'<div style="background:#fef2f2;border-left:4px solid #ef4444;border-radius:10px;'
                    f'padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">'
                    f'<div style="font-size:1.3rem;margin-bottom:.2rem;">🚫</div>'
                    f'<div style="font-size:0.72rem;color:#991b1b;font-weight:600;text-transform:uppercase;">'
                    f'1/2 UL (terlambat masuk >120 mnt)</div>'
                    f'<div style="font-size:1.7rem;font-weight:700;color:#dc2626;">'
                    f'{n_k}<span style="font-size:.9rem"> hari</span></div></div>',
                    unsafe_allow_html=True,
                )

            combined = pd.concat([late_df, k_df]).sort_values("Tanggal").reset_index(drop=True)
            combined["No."]         = range(1, len(combined) + 1)
            combined["Lebih Awal"]  = combined.apply(_menit_lebih_awal, axis=1)
            combined["Klasifikasi"] = combined["Klasifikasi"].fillna("-")
            st.dataframe(
                combined[["No.", "Tanggal", "Klasifikasi", "Shift", "Jam Masuk", "Jam Keluar", "Lebih Awal"]],
                width="stretch",
                height=min(60 + len(combined) * 35, 380),
                hide_index=True,
                column_config={
                    "Lebih Awal": st.column_config.TextColumn("Pulang Lebih Awal", width="medium"),
                },
            )

    # Expander 2: DW & K-Sick
    with st.expander(
        f"🏥 DW & Sakit  —  DW: {n_dw}  |  K-Sick: {n_k_sick}",
        expanded=False,
    ):
        if n_dw == 0 and n_k_sick == 0:
            st.info("ℹ️ Tidak ada data DW / K-Sick pada periode ini.")
        else:
            dc1, dc2 = st.columns(2)
            with dc1:
                st.markdown(
                    f'<div style="background:#fff7ed;border-left:4px solid #f97316;border-radius:10px;'
                    f'padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">'
                    f'<div style="font-size:1.3rem;margin-bottom:.2rem;">🚷</div>'
                    f'<div style="font-size:0.72rem;color:#9a3412;font-weight:600;text-transform:uppercase;">DW (Absence)</div>'
                    f'<div style="font-size:1.7rem;font-weight:700;color:#ea580c;">'
                    f'{n_dw}<span style="font-size:.9rem"> hari</span></div></div>',
                    unsafe_allow_html=True,
                )
            with dc2:
                st.markdown(
                    f'<div style="background:#fdf2f8;border-left:4px solid #ec4899;border-radius:10px;'
                    f'padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">'
                    f'<div style="font-size:1.3rem;margin-bottom:.2rem;">🩺</div>'
                    f'<div style="font-size:0.72rem;color:#9d174d;font-weight:600;text-transform:uppercase;">K-Sick W Letter</div>'
                    f'<div style="font-size:1.7rem;font-weight:700;color:#db2777;">'
                    f'{n_k_sick}<span style="font-size:.9rem"> hari</span></div></div>',
                    unsafe_allow_html=True,
                )

            if n_dw > 0:
                st.markdown("**🚷 DW — Tidak Hadir (Absence)**")
                dw_df["No."] = range(1, len(dw_df) + 1)
                st.dataframe(
                    dw_df[["No.", "Tanggal", "Shift", "Status"]],
                    width="stretch",
                    height=min(60 + len(dw_df) * 35, 280),
                    hide_index=True,
                    column_config={"Status": st.column_config.TextColumn("Attendance Results", width="large")},
                )
            if n_k_sick > 0:
                st.markdown("**🩺 K-Sick — Sakit dengan Surat**")
                k_sick_df["No."] = range(1, len(k_sick_df) + 1)
                st.dataframe(
                    k_sick_df[["No.", "Tanggal", "Shift", "Status", "Klasifikasi"]],
                    width="stretch",
                    height=min(60 + len(k_sick_df) * 35, 280),
                    hide_index=True,
                    column_config={"Status": st.column_config.TextColumn("Attendance Results", width="large")},
                )

    # Expander 3: Leave (AL / 1/2 AL / WFA)
    with st.expander(
        f"🏖️ Rincian Leave  —  AL: {n_al}  |  WFA: {n_wfa}",
        expanded=False,
    ):
        if n_al == 0 and n_wfa == 0:
            st.info("ℹ️ Tidak ada data AL / 1/2 AL / WFA pada periode ini.")
        else:
            lc1, lc2, lc3 = st.columns(3)
            n_full_al = len(detail_df[detail_df["Klasifikasi_raw"].apply(lambda x: has_status(x, "AL"))])
            n_half_al = len(detail_df[detail_df["Klasifikasi_raw"].apply(lambda x: has_status(x, "1/2 AL"))])
            for (col, label, icon, val, bg, bc, tc) in [
                (lc1, "AL (Full)",        "🌴", n_full_al, "#fdf4ff", "#a855f7", "#7e22ce"),
                (lc2, "1/2 AL (Setengah)","🌅", n_half_al, "#fff1f2", "#fb7185", "#be123c"),
                (lc3, "WFA",              "🏠", n_wfa,     "#f0f9ff", "#0ea5e9", "#0369a1"),
            ]:
                with col:
                    st.markdown(
                        f'<div style="background:{bg};border-left:4px solid {bc};border-radius:10px;'
                        f'padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">'
                        f'<div style="font-size:1.3rem;margin-bottom:.2rem;">{icon}</div>'
                        f'<div style="font-size:0.72rem;color:{tc};font-weight:600;text-transform:uppercase;">{label}</div>'
                        f'<div style="font-size:1.7rem;font-weight:700;color:{tc};">'
                        f'{val}<span style="font-size:.9rem"> hari</span></div></div>',
                        unsafe_allow_html=True,
                    )

            leave_combined = pd.concat([al_df, wfa_df]).sort_values("Tanggal").reset_index(drop=True)
            leave_combined = leave_combined.drop_duplicates(subset=["Tanggal"]).reset_index(drop=True)
            leave_combined["No."] = range(1, len(leave_combined) + 1)
            st.dataframe(
                leave_combined[["No.", "Tanggal", "Klasifikasi", "Shift", "Jam Masuk", "Jam Keluar", "Status"]],
                width="stretch",
                height=min(60 + len(leave_combined) * 35, 380),
                hide_index=True,
                column_config={"Status": st.column_config.TextColumn("Attendance Results", width="large")},
            )

    # Expander 4: Rincian per Tipe Shift
    tipe_counts = detail_df["Tipe"].value_counts()
    s1_n = tipe_counts.get("S1", 0)
    s2_n = tipe_counts.get("S2", 0)
    h_n  = tipe_counts.get("H",  0)
    with st.expander(
        f"📊 Rincian per Tipe Shift  —  ☀️ S1: {s1_n}  |  🌙 S2: {s2_n}  |  🏖️ H: {h_n}",
        expanded=False,
    ):
        tab_s1, tab_s2, tab_h = st.tabs(["☀️ S1", "🌙 S2", "🏖️ H"])

        def _render_tipe_tab(tipe_key):
            df_tipe = detail_df[detail_df["Tipe"] == tipe_key].copy().reset_index(drop=True)
            if df_tipe.empty:
                st.info(f"ℹ️ Tidak ada data tipe {tipe_key} pada periode ini.")
                return
            df_tipe["No."]       = range(1, len(df_tipe) + 1)
            df_tipe["Jam Kerja"] = df_tipe["Jam Kerja"].apply(lambda x: f"{x:.1f} jam" if x > 0 else "-")
            total_jam = detail_df[detail_df["Tipe"] == tipe_key]["Jam Kerja"].sum()
            st.caption(f"📅 {len(df_tipe)} hari  |  ⏱ Total jam kerja: {total_jam:.1f} jam")
            st.dataframe(
                df_tipe[["No.", "Tanggal", "Shift", "Jam Masuk", "Jam Keluar", "Klasifikasi", "Jam Kerja"]],
                width="stretch",
                height=min(60 + len(df_tipe) * 35, 420),
                hide_index=True,
            )

        with tab_s1: _render_tipe_tab("S1")
        with tab_s2: _render_tipe_tab("S2")
        with tab_h:  _render_tipe_tab("H")

    # Expander 5: Detail Lengkap
    with st.expander(f"🗒️ Detail Lengkap per Hari  —  {len(detail_df)} hari tercatat", expanded=False):
        TIPE_LABEL = {"S1": "☀️ S1", "S2": "🌙 S2", "H": "🏖️ H"}
        dd = detail_df.copy()
        dd["Tipe"]      = dd["Tipe"].map(lambda x: TIPE_LABEL.get(x, x))
        dd["Jam Kerja"] = dd["Jam Kerja"].apply(lambda x: f"{x:.1f} jam" if x > 0 else "-")
        st.dataframe(
            dd[["No.", "Tanggal", "Tipe", "Shift", "Jam Masuk", "Jam Keluar", "Status", "Klasifikasi", "Jam Kerja"]],
            width="stretch",
            height=420,
            hide_index=True,
            column_config={"Status": st.column_config.TextColumn("Status Absensi", width="large")},
        )

    st.caption("💡 Klik di luar kotak ini untuk menutup")


# ──────────────────────────────────────────────────────────────
# Proses File
# ──────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def process_file(file_bytes):
    buf = io.BytesIO(file_bytes)
    df = pd.read_excel(
        buf,
        sheet_name="General statistics and attendan",
        header=4,
        dtype={"Earliest": str, "Latest": str},
    )

    required = ["Name", "Account", "Rules", "Shift", "Earliest", "Latest",
                "Attendance results", "Leave & Overtime Application"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom tidak ditemukan: {missing}")

    df = df[required].dropna(subset=["Account", "Rules"])
    df = df[~df["Account"].astype(str).str.strip().isin(["", "--"])]

    df["_statuses"] = df.apply(
        lambda r: classify(
            r["Earliest"], r["Shift"], r["Attendance results"],
            latest_raw=r["Latest"],
            leave_app=r["Leave & Overtime Application"],
        ),
        axis=1,
    )

    all_employees = (
        df.groupby("Account")["Rules"]
        .agg(lambda x: x.mode()[0])
        .reset_index()
    )

    df_classified = df[df["_statuses"].notna()].copy()
    df_exploded   = df_classified.explode("_statuses").rename(columns={"_statuses": "Status"})

    pivot = df_exploded.pivot_table(
        index="Account",
        columns="Status",
        values="Shift",
        aggfunc="count",
        fill_value=0,
    ).reset_index()
    pivot.columns.name = None

    for col in ["Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA", "DW", "K", "Off"]:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot = all_employees.merge(pivot, on="Account", how="left").fillna(0)
    for col in ["Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA", "DW", "K", "Off"]:
        pivot[col] = pivot[col].astype(int)

    name_map = df.groupby("Account")["Name"].first()
    pivot["Nama"] = pivot["Account"].map(name_map)

    pivot = pivot.sort_values(["Rules", "Nama"]).reset_index(drop=True)
    pivot.insert(0, "No.", range(1, len(pivot) + 1))
    result = pivot[["No.", "Nama", "Account", "Rules",
                    "Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA",
                    "DW", "K", "Off"]].copy()

    stats = {
        "total_rows": len(df),
        "classified": len(df_classified),
        "skipped"   : len(df) - len(df_classified),
        "employees" : len(result),
        "dist"      : df_exploded["Status"].value_counts(dropna=False).to_dict(),
    }
    return result, stats


# ──────────────────────────────────────────────────────────────
# Export Excel
# ──────────────────────────────────────────────────────────────

def to_excel_bytes(df, time_range=""):
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
    PURPLE = PatternFill("solid", fgColor="FDF4FF")
    PINK   = PatternFill("solid", fgColor="FFF1F2")
    CYAN   = PatternFill("solid", fgColor="F0F9FF")
    ORANGE = PatternFill("solid", fgColor="FFF7ED")
    ROSE   = PatternFill("solid", fgColor="FDF2F8")
    SLATE  = PatternFill("solid", fgColor="F8FAFC")

    thin   = Side(style="thin", color="CBD5E1")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center")
    LEFT   = Alignment(horizontal="left",   vertical="center")

    ws.merge_cells("A1:M1")
    ws["A1"] = "Summary Attendance"
    ws["A1"].font      = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    ws["A1"].fill      = DARK
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:M2")
    ws["A2"] = f"Time Range: {time_range}" if time_range else "Rekap Absensi"
    ws["A2"].font      = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    ws["A2"].fill      = BLUE
    ws["A2"].alignment = CENTER
    ws.row_dimensions[2].height = 16

    headers = ["No.", "Nama", "Account", "Rules",
               "Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA",
               "DW", "K", "Off"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        c.fill      = BLUE
        c.alignment = CENTER
        c.border    = BORDER
    ws.row_dimensions[3].height = 20

    col_style = {
        5:  (GREEN,  "166534"),
        6:  (YELLOW, "92400E"),
        7:  (RED,    "991B1B"),
        8:  (PURPLE, "7E22CE"),
        9:  (PINK,   "BE123C"),
        10: (CYAN,   "0369A1"),
        11: (ORANGE, "EA580C"),
        12: (ROSE,   "DB2777"),
        13: (SLATE,  "475569"),
    }

    ALL_STATUS_COLS = ["Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA", "DW", "K", "Off"]

    for ri, (_, row) in enumerate(df.iterrows()):
        er = ri + 4
        base_fill = ALT if ri % 2 == 0 else WHITE
        vals = [row["No."], row["Nama"], row["Account"], row["Rules"]] + \
               [row.get(c, 0) for c in ALL_STATUS_COLS]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=er, column=ci, value=v)
            c.border = BORDER
            c.font   = Font(name="Calibri", size=10)
            if ci <= 4:
                c.fill      = base_fill
                c.alignment = LEFT if ci > 1 else CENTER
            else:
                fill_pos, fc_pos = col_style[ci]
                c.fill      = fill_pos if v > 0 else base_fill
                c.font      = Font(name="Calibri", size=10,
                                   bold=(v > 0),
                                   color=fc_pos if v > 0 else "000000")
                c.alignment = CENTER
        ws.row_dimensions[er].height = 17

    tr = len(df) + 4
    ws.merge_cells(f"A{tr}:D{tr}")
    ws[f"A{tr}"]           = "TOTAL"
    ws[f"A{tr}"].font      = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    ws[f"A{tr}"].fill      = DARK
    ws[f"A{tr}"].alignment = CENTER
    ws[f"A{tr}"].border    = BORDER
    for ci in range(5, 14):
        c       = ws.cell(row=tr, column=ci)
        c.value = f"=SUM({get_column_letter(ci)}4:{get_column_letter(ci)}{tr-1})"
        c.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        c.fill  = DARK
        c.alignment = CENTER
        c.border    = BORDER

    for ci, w in enumerate([6, 32, 24, 28, 10, 10, 8, 8, 8, 8, 8, 8, 8], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes    = "A4"
    ws.auto_filter.ref = f"A3:M{tr-1}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# Definisi kolom
# ──────────────────────────────────────────────────────────────

CORE_COLS = ["No.", "Nama", "Account", "Rules", "Normal", "Late", "1/2 UL", "DW"]

OPTIONAL_COLS_DEF = [
    ("K",     "🩺 K (Sakit)",  "Sakit dgn Surat"),
    ("AL",    "🌴 AL",         "Annual Leave penuh"),
    ("1/2 AL","🌅 1/2 AL",     "Annual Leave setengah hari"),
    ("WFA",   "🏠 WFA",        "Work From Anywhere"),
    ("Off",   "🔕 Off",        "Rest / Not scheduled"),
]
OPTIONAL_KEYS   = [c[0] for c in OPTIONAL_COLS_DEF]
OPTIONAL_LABELS = {c[0]: c[1] for c in OPTIONAL_COLS_DEF}
OPTIONAL_DESCS  = {c[0]: c[2] for c in OPTIONAL_COLS_DEF}

COL_CONFIG_ALL = {
    "No."    : st.column_config.NumberColumn("No.", width="small"),
    "Nama"   : st.column_config.TextColumn("Nama", width="large"),
    "Account": st.column_config.TextColumn("Account", width="medium"),
    "Rules"  : st.column_config.TextColumn("Rules", width="medium"),
    "Normal" : st.column_config.NumberColumn("✅ Normal", format="%d", width="small"),
    "Late"   : st.column_config.NumberColumn("🕐 Late",   format="%d", width="small"),
    "1/2 UL" : st.column_config.NumberColumn("🚫 1/2 UL", format="%d", width="small"),
    "DW"     : st.column_config.NumberColumn("🚷 DW",     format="%d", width="small"),
    "K"      : st.column_config.NumberColumn("🩺 K",      format="%d", width="small"),
    "AL"     : st.column_config.NumberColumn("🌴 AL",     format="%d", width="small"),
    "1/2 AL" : st.column_config.NumberColumn("🌅 1/2 AL", format="%d", width="small"),
    "WFA"    : st.column_config.NumberColumn("🏠 WFA",    format="%d", width="small"),
    "Off"    : st.column_config.NumberColumn("🔕 Off",    format="%d", width="small"),
}


# ──────────────────────────────────────────────────────────────
# Dialog: Logic Klasifikasi  (DIPERBARUI — detail lengkap)
# ──────────────────────────────────────────────────────────────

_LOGIC_HTML = (
    '<div style="font-size:0.85rem;color:#334155;line-height:1.9;">'

    # ── Intro singkat ──
    '<div style="background:#f0f9ff;border-left:4px solid #0ea5e9;border-radius:8px;'
    'padding:0.7rem 1rem;margin-bottom:1.2rem;font-size:0.82rem;color:#0369a1;">'
    '<b>ℹ️ Tentang Logic ini:</b> Setiap baris absensi diproses melalui <b>8 langkah prioritas</b> '
    'secara berurutan. Begitu satu kondisi cocok, proses berhenti dan status ditetapkan. '
    'Satu hari dapat menghasilkan <b>dua status sekaligus</b> (dual-count) khusus untuk '
    'K-Sick, AL, 1/2 AL, dan WFA.'
    '</div>'

    # ── Tabel Status ──
    '<div style="font-weight:700;color:#0f172a;margin-bottom:0.4rem;font-size:0.82rem;'
    'text-transform:uppercase;letter-spacing:0.06em;">📋 Status &amp; Kondisi Pemicu</div>'

    '<table style="width:100%;border-collapse:collapse;margin-bottom:1.2rem;">'

    '<tr style="background:#f1f5f9;">'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;width:110px;">✅ Normal</td>'
    '<td style="padding:0.4rem 0.7rem;">Att Results mengandung <code>"normal"</code>, '
    'Punch In <b>tepat waktu atau lebih awal</b> dari shift start, '
    'dan tidak ada kondisi leave / keterlambatan lain.</td>'
    '</tr>'

    '<tr>'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;">🕐 Late</td>'
    '<td style="padding:0.4rem 0.7rem;">'
    '<b>Punch In terlambat 1–120 menit</b> dari jam mulai shift.<br>'
    'Berlaku <b>tanpa memandang Att. Results</b> — hanya selisih waktu Punch In vs shift start.<br>'
    'Output: <code>["Late"]</code> — bersifat standalone, tidak dual-count dengan Normal.</td>'
    '</tr>'

    '<tr style="background:#f1f5f9;">'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;">🚫 1/2 UL</td>'
    '<td style="padding:0.4rem 0.7rem;">'
    '<b>Punch In terlambat lebih dari 120 menit</b> dari jam mulai shift.<br>'
    'Berlaku <b>tanpa memandang Att. Results</b> — hanya selisih waktu Punch In vs shift start.<br>'
    'Output: <code>["1/2 UL"]</code> — bersifat standalone, tidak dual-count dengan Normal.</td>'
    '</tr>'

    '<tr>'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;">🚷 DW</td>'
    '<td style="padding:0.4rem 0.7rem;">Att Results mengandung kata <code>"Absence"</code> '
    '— karyawan tidak hadir sama sekali tanpa keterangan.</td>'
    '</tr>'

    '<tr style="background:#f1f5f9;">'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;">🩺 K</td>'
    '<td style="padding:0.4rem 0.7rem;">Leave mengandung <code>"K-Sick W Letter"</code>. '
    'Dicek <em>sebelum</em> logika Late / Normal.<br>'
    'Jika att juga mengandung "Normal": <b>Normal + K</b>. Jika tidak: <b>K saja</b>.</td>'
    '</tr>'

    '<tr>'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;">🔕 Off</td>'
    '<td style="padding:0.4rem 0.7rem;">Att Results bernilai tepat '
    '<code>"Normal (rest)"</code> atau <code>"Normal (not scheduled)"</code> '
    '— dicek <b>paling awal</b>, sebelum semua logika lainnya. '
    'Output: <code>["Normal", "Off"]</code>.</td>'
    '</tr>'

    '<tr style="background:#f1f5f9;">'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;">🌴 AL</td>'
    '<td style="padding:0.4rem 0.7rem;">Leave = <code>AnnualLeave</code> + Att normal '
    '+ <b>tidak ada punch sama sekali</b> → Normal + AL</td>'
    '</tr>'

    '<tr>'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;">🌅 1/2 AL</td>'
    '<td style="padding:0.4rem 0.7rem;">Leave = <code>AnnualLeave</code> + Att normal '
    '+ <b>ada Punch In dan Punch Out</b> → Normal + 1/2 AL</td>'
    '</tr>'

    '<tr style="background:#f1f5f9;">'
    '<td style="padding:0.4rem 0.7rem;font-weight:600;white-space:nowrap;">🏠 WFA</td>'
    '<td style="padding:0.4rem 0.7rem;">Leave mengandung <code>"WorkFromHome"</code> / '
    '<code>"WFH"</code> + Att normal → selalu <b>Normal + WFA</b></td>'
    '</tr>'

    '</table>'

    # ── Alur Keputusan ──
    '<div style="font-weight:700;color:#0f172a;margin-bottom:0.4rem;font-size:0.82rem;'
    'text-transform:uppercase;letter-spacing:0.06em;">🔀 Alur Keputusan (Urutan Prioritas)</div>'

    '<div style="background:#fff;border:1px solid #e2e8f0;border-radius:8px;'
    'padding:0.8rem 1rem;font-family:monospace;font-size:0.78rem;line-height:2.1;'
    'margin-bottom:1.2rem;color:#475569;">'
    '1. Att = "Normal (rest)" / "Normal (not scheduled)" &rarr; <b>🔕 Off</b> &mdash; selesai<br>'
    '2. Shift = Rest / Not scheduled / kosong / "--" &rarr; <b>dilewati</b> (None)<br>'
    '3. Att mengandung "Absence" &rarr; <b>🚷 DW</b> &mdash; selesai<br>'
    '4. Leave mengandung "K-Sick W Letter"?<br>'
    '&nbsp;&nbsp;&nbsp;+-- Att juga mengandung "normal" &rarr; <b>✅ Normal + 🩺 K</b><br>'
    '&nbsp;&nbsp;&nbsp;+-- Att tidak mengandung "normal" &rarr; <b>🩺 K</b><br>'
    '5. Att mengandung "normal" + "leave"?<br>'
    '&nbsp;&nbsp;&nbsp;+-- Leave = AnnualLeave + tidak ada punch &rarr; <b>✅ Normal + 🌴 AL</b><br>'
    '&nbsp;&nbsp;&nbsp;+-- Leave = AnnualLeave + ada kedua punch &rarr; <b>✅ Normal + 🌅 1/2 AL</b><br>'
    '&nbsp;&nbsp;&nbsp;+-- Leave = WFH / WorkFromHome &rarr; <b>✅ Normal + 🏠 WFA</b><br>'
    '6. Cek <b>Punch In vs Shift Start</b> (tanpa memandang Att. Results):<br>'
    '&nbsp;&nbsp;&nbsp;+-- Terlambat 1–120 mnt &rarr; <b>🕐 Late</b> &mdash; selesai<br>'
    '&nbsp;&nbsp;&nbsp;+-- Terlambat &gt; 120 mnt &rarr; <b>🚫 1/2 UL</b> &mdash; selesai<br>'
    '7. Punch In tepat / lebih awal + Att mengandung "normal" &rarr; <b>✅ Normal</b><br>'
    '8. Selain itu &rarr; <b>tidak diklasifikasi</b> (None / dilewati)'
    '</div>'

    # ── Aturan Late & 1/2 UL ──
    '<div style="font-weight:700;color:#0f172a;margin-bottom:0.4rem;font-size:0.82rem;'
    'text-transform:uppercase;letter-spacing:0.06em;">⏱ Aturan Late &amp; 1/2 UL — Berbasis Punch In</div>'

    '<div style="background:#fefce8;border-radius:8px;padding:0.7rem 1rem;margin-bottom:1.2rem;'
    'font-size:0.82rem;border-left:3px solid #eab308;">'
    '<b>Satu-satunya penentu Late / 1/2 UL adalah selisih Punch In vs jam mulai shift:</b><br>'
    '— <b>🕐 Late</b>: Punch In tercatat <b>1–120 menit</b> setelah jam mulai shift<br>'
    '— <b>🚫 1/2 UL</b>: Punch In tercatat <b>lebih dari 120 menit</b> setelah jam mulai shift<br>'
    '— Att. Results <b>tidak diperiksa</b> untuk menentukan Late / 1/2 UL<br>'
    '— Output bersifat <em>standalone</em> — tidak ada dual-count dengan Normal<br><br>'
    '<b>Contoh:</b> Shift 09:30–18:30, Punch In 17:21 &rarr; '
    'selisih 471 menit &gt; 120 menit &rarr; <b>🚫 1/2 UL</b>'
    '</div>'

    # ── Kepulangan Lebih Awal (Late Out) ──
    '<div style="font-weight:700;color:#0f172a;margin-bottom:0.4rem;font-size:0.82rem;'
    'text-transform:uppercase;letter-spacing:0.06em;">🏃 Kepulangan Lebih Awal (Late Out)</div>'

    '<div style="background:#fef9ec;border-radius:8px;padding:0.7rem 1rem;margin-bottom:1.2rem;'
    'font-size:0.82rem;border-left:3px solid #f59e0b;">'
    'Jika Punch In <b>tepat waktu / lebih awal</b> (langkah 6a = None), '
    'sistem memeriksa Punch Out vs jam selesai shift (langkah 6b):<br>'
    '— Pulang lebih awal <b>1–120 menit</b> dari shift end &rarr; <b>🕐 Late</b><br>'
    '— Pulang lebih awal <b>&gt; 120 menit</b> dari shift end &rarr; <b>🚫 1/2 UL</b><br>'
    '— Mendukung shift overnight (mis. 19:00–05:00 hari berikutnya)'
    '</div>'

    # ── Dual-Count ──
    '<div style="font-weight:700;color:#0f172a;margin-bottom:0.4rem;font-size:0.82rem;'
    'text-transform:uppercase;letter-spacing:0.06em;">2️⃣ Dual-Count (status ganda dalam satu hari)</div>'

    '<div style="background:#eff6ff;border-radius:8px;padding:0.6rem 1rem;margin-bottom:1.2rem;'
    'font-size:0.82rem;border-left:3px solid #3b82f6;">'
    'Satu hari bisa menghasilkan <b>2 status sekaligus</b> khusus untuk:<br>'
    '— K-Sick + att Normal &rarr; <b>✅ Normal + 🩺 K</b><br>'
    '— AnnualLeave + att Normal + ada punch &rarr; <b>✅ Normal + 🌅 1/2 AL</b><br>'
    '— AnnualLeave + att Normal + tanpa punch &rarr; <b>✅ Normal + 🌴 AL</b><br>'
    '— WFH + att Normal &rarr; <b>✅ Normal + 🏠 WFA</b><br>'
    '<span style="color:#ef4444;font-weight:600;">🕐 Late dan 🚫 1/2 UL tidak dual-count</span>'
    ' — output-nya adalah status tunggal berdasarkan keterlambatan Punch In.'
    '</div>'

    # ── Penyimpanan Database ──
    '<div style="font-weight:700;color:#0f172a;margin-bottom:0.4rem;font-size:0.82rem;'
    'text-transform:uppercase;letter-spacing:0.06em;">🗄️ Data yang Disimpan ke Database</div>'

    '<div style="background:#f8fafc;border-radius:8px;padding:0.7rem 1rem;margin-bottom:1.2rem;'
    'font-size:0.82rem;border-left:3px solid #94a3b8;">'
    '<b>Tabel <code>karyawan</code>:</b> account, nama, rules, department<br>'
    '<b>Tabel <code>absensi_harian</code></b> (per baris per hari):<br>'
    '&nbsp;&nbsp;tanggal, shift, tipe_shift (S1/S2/H), jam_masuk, jam_keluar, jam_kerja,<br>'
    '&nbsp;&nbsp;status_absensi (Attendance results asli), status_klasifikasi (misal <code>Normal/WFA</code>), periode (YYYY-MM)<br>'
    '<b>Tidak disimpan:</b> nama file, kolom Leave &amp; Overtime Application, hasil pivot/rekap '
    '(dihitung ulang saat dipanggil)'
    '</div>'

    # ── Pengecualian ──
    '<div style="font-weight:700;color:#0f172a;margin-bottom:0.4rem;font-size:0.82rem;'
    'text-transform:uppercase;letter-spacing:0.06em;">⚠️ Pengecualian &amp; Catatan Penting</div>'

    '<div style="background:#fef9ec;border-radius:8px;padding:0.6rem 1rem;'
    'font-size:0.82rem;border-left:3px solid #f59e0b;">'
    '— Shift <code>Rest</code> / <code>Not scheduled</code> / <code>--</code> / kosong '
    '→ dilewati, <em>kecuali</em> att = "Normal (rest)" yang masuk sebagai Off<br>'
    '— K-Sick diperiksa <em>sebelum</em> blok AL/WFA agar tidak tertukar<br>'
    '— Karyawan dengan AL / WFA / K-Sick <b>tidak dikenai</b> cek keterlambatan Punch In<br>'
    '— Jika tidak ada data Punch In (<code>not punched</code>) → tidak ada penalti Late / 1/2 UL'
    '</div>'

    '</div>'
)


@st.dialog("📋 Logic Klasifikasi Absensi", width="large")
def show_logic_dialog():
    st.markdown(_LOGIC_HTML, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# UI Utama
# ──────────────────────────────────────────────────────────────

st.markdown(
    '<div class="app-header">'
    '<div class="badge">🏢 HR TOOLS</div>'
    '<h1>🗓️ Absensi Rekap Generator</h1>'
    '<p>📤 Upload file Excel absensi &rarr; ⚙️ Hitung Normal / Late / K / AL / WFA / DW per karyawan &rarr; 📥 Download hasil</p>'
    '</div>',
    unsafe_allow_html=True,
)

# ── Upload + Tombol Logic ──
upload_col, btn_col = st.columns([5, 1], gap="medium")

with upload_col:
    st.markdown('<p class="section-title">📁 Upload File Excel</p>', unsafe_allow_html=True)
    periodes_tersedia = get_periodes()
    if periodes_tersedia:
        st.markdown("**📂 Atau pilih periode yang sudah tersimpan:**")
        periode_dipilih = st.selectbox(
            label="",
            options=["- Upload file baru -"] + periodes_tersedia,
            label_visibility="collapsed",
        )
    else:
        periode_dipilih = "- Upload file baru -"

    uploaded = st.file_uploader(
        label="",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
        help="File Excel dari sistem absensi. Sheet 'General statistics and attendan' harus ada.",
    )

with btn_col:
    st.markdown("<div style='margin-top:2.1rem'></div>", unsafe_allow_html=True)
    if st.button("📋 Logic\nKlasifikasi", use_container_width=True, type="secondary"):
        st.session_state.dialog_target = "logic"
        st.session_state.dialog_emp    = None
        st.rerun()

# ── Proses ──
_NEW_PERIODE_SENTINEL = "- Upload file baru -"

if uploaded is not None or periode_dipilih != _NEW_PERIODE_SENTINEL:
    if uploaded is not None:
        file_bytes = uploaded.read()

        with st.spinner("⚙️ Memproses data absensi..."):
            try:
                df_result, stats = process_file(file_bytes)
            except Exception as e:
                st.error(f"❌ Gagal memproses file: {e}")
                st.stop()

        try:
            import io as _io, re as _re, pandas as _pd
            _buf = _io.BytesIO(file_bytes)
            df_raw = _pd.read_excel(
                _buf,
                sheet_name="General statistics and attendan",
                header=4,
                dtype={"Earliest": str, "Latest": str},
            )
            _periode = None
            for _val in df_raw["Time"].astype(str):
                _m = _re.search(r'(\d{4})/(\d{2})/(\d{2})', _val)
                if _m:
                    _periode = f"{_m.group(1)}-{_m.group(2)}"
                    break
            if _periode is None:
                _periode = "unknown"

            df_raw = df_raw.dropna(subset=["Account", "Rules"])
            df_raw = df_raw[~df_raw["Account"].astype(str).str.strip().isin(["", "--"])]
            df_raw["_tipe_shift"] = df_raw["Shift"].apply(classify_shift_type)
            df_raw["_status_klasifikasi"] = df_raw.apply(
                lambda r: classify_str(
                    r["Earliest"], r["Shift"], r["Attendance results"],
                    latest_raw=r["Latest"],
                    leave_app=r.get("Leave & Overtime Application"),
                ), axis=1,
            )
            save_periode(df_raw, _periode)
        except Exception as e:
            st.warning(f"⚠️ Gagal simpan ke database: {e}")
    else:
        df_raw_db = get_rekap(periode_dipilih)
        df_result = df_raw_db.rename(columns={
            "nama": "Nama", "account": "Account", "rules": "Rules",
            "normal": "Normal", "late": "Late", "half_ul": "1/2 UL",
            "half_al": "1/2 AL", "al": "AL", "wfa": "WFA",
            "dw": "DW", "k_sick": "K", "off_count": "Off",
        })
        for col in ["Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA", "DW", "K", "Off"]:
            if col not in df_result.columns:
                df_result[col] = 0
        file_bytes = None
        stats = {
            "total_rows": len(df_result),
            "classified": int(df_result[["Normal", "Late", "1/2 UL"]].sum().sum()),
            "skipped": 0,
            "employees": len(df_result),
            "dist": {},
        }
        df_result.insert(0, "No.", range(1, len(df_result) + 1))

    # ── Metric Cards ──
    total_n   = int(df_result["Normal"].sum())
    total_l   = int(df_result["Late"].sum())
    total_k   = int(df_result["1/2 UL"].sum())
    total_al  = int(df_result["AL"].sum())
    total_hal = int(df_result["1/2 AL"].sum())
    total_wfa = int(df_result["WFA"].sum())
    total_dw  = int(df_result["DW"].sum())
    total_ks  = int(df_result["K"].sum())
    total_off = int(df_result["Off"].sum())
    total_e   = stats["employees"]

    st.markdown(f"""
<div class="metric-row">
  <div class="metric-card metric-normal">
    <span class="icon">✅</span>
    <div class="label">Normal</div>
    <div class="value">{total_n:,}</div>
    <div class="sub">Hadir tepat / lebih awal</div>
  </div>
  <div class="metric-card metric-late">
    <span class="icon">🕐</span>
    <div class="label">Late</div>
    <div class="value">{total_l:,}</div>
    <div class="sub">Terlambat masuk 1–120 mnt</div>
  </div>
  <div class="metric-card metric-k">
    <span class="icon">🚫</span>
    <div class="label">1/2 UL</div>
    <div class="value">{total_k:,}</div>
    <div class="sub">Terlambat masuk &gt;120 mnt</div>
  </div>
  <div class="metric-card metric-dw">
    <span class="icon">🚷</span>
    <div class="label">DW</div>
    <div class="value">{total_dw:,}</div>
    <div class="sub">Absence / Tidak hadir</div>
  </div>
  <div class="metric-card metric-total">
    <span class="icon">👥</span>
    <div class="label">Karyawan</div>
    <div class="value">{total_e:,}</div>
    <div class="sub">Total dalam periode</div>
  </div>
</div>
<div class="metric-row" style="margin-top:-1rem;">
  <div class="metric-card metric-ksick">
    <span class="icon">🩺</span>
    <div class="label">K-Sick</div>
    <div class="value">{total_ks:,}</div>
    <div class="sub">Sakit dengan surat</div>
  </div>
  <div class="metric-card metric-al">
    <span class="icon">🌴</span>
    <div class="label">AL</div>
    <div class="value">{total_al:,}</div>
    <div class="sub">Annual Leave penuh</div>
  </div>
  <div class="metric-card metric-half-al">
    <span class="icon">🌅</span>
    <div class="label">1/2 AL</div>
    <div class="value">{total_hal:,}</div>
    <div class="sub">Annual Leave setengah hari</div>
  </div>
  <div class="metric-card metric-wfa">
    <span class="icon">🏠</span>
    <div class="label">WFA</div>
    <div class="value">{total_wfa:,}</div>
    <div class="sub">Work From Anywhere</div>
  </div>
  <div class="metric-card metric-off">
    <span class="icon">🔕</span>
    <div class="label">Off</div>
    <div class="value">{total_off:,}</div>
    <div class="sub">Rest / Not scheduled</div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Filter & Tabel ──
    st.markdown('<p class="section-title">👥 Hasil Summary per Karyawan</p>', unsafe_allow_html=True)

    fcol1, fcol2, fcol3 = st.columns([2, 2, 1])
    with fcol1:
        all_rules = sorted(df_result["Rules"].unique().tolist())
        sel_rules = st.multiselect("🔖 Filter Rules", options=all_rules, placeholder="Semua Rules")
    with fcol2:
        search = st.text_input("🔍 Cari Nama / Account", placeholder="Ketik nama atau account...")
    with fcol3:
        show_late_only = st.checkbox("⚠️ Hanya Late/K/DW", value=False)

    with st.expander("👁️ Tampilkan / Sembunyikan Kolom Kategori", expanded=False):
        st.markdown(
            '<div style="font-size:0.83rem;color:#64748b;margin-bottom:0.6rem;">'
            'Kolom <b>No., Nama, Account, Rules, Normal, Late, 1/2 UL, DW</b> selalu tampil. '
            'Pilih kategori tambahan di bawah:</div>',
            unsafe_allow_html=True,
        )
        opt_cols_selected = []
        oc1, oc2, oc3, oc4, oc5 = st.columns(5)
        opt_col_ui = [oc1, oc2, oc3, oc4, oc5]
        for i, (key, label, desc) in enumerate(OPTIONAL_COLS_DEF):
            with opt_col_ui[i]:
                checked = st.checkbox(label, value=False, help=desc, key=f"col_{key}")
                if checked:
                    opt_cols_selected.append(key)

    visible_cols = CORE_COLS + opt_cols_selected

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
        df_show = df_show[
            (df_show["Late"] > 0) | (df_show["1/2 UL"] > 0) | (df_show["DW"] > 0)
        ]

    df_show = df_show.copy()
    df_show["No."] = range(1, len(df_show) + 1)

    hidden = [OPTIONAL_LABELS[k] for k in OPTIONAL_KEYS if k not in opt_cols_selected]
    st.caption(
        "💡 Klik baris untuk melihat rincian harian" +
        (f"  |  Kolom tersembunyi: {', '.join(hidden)}" if hidden else "")
    )

    sel_event = st.dataframe(
        df_show[visible_cols],
        width="stretch",
        height=520,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        column_config={k: v for k, v in COL_CONFIG_ALL.items() if k in visible_cols},
    )

    sel_rows = sel_event.selection.rows if sel_event and sel_event.selection else []
    if sel_rows and file_bytes is not None:
        idx = sel_rows[0]
        if idx < len(df_show):
            emp = df_show.iloc[idx]
            new_emp = {"account": emp["Account"], "nama": emp["Nama"], "rules": emp["Rules"]}
            if st.session_state.dialog_emp != new_emp or st.session_state.dialog_target != "detail":
                st.session_state.dialog_target = "detail"
                st.session_state.dialog_emp    = new_emp
                st.rerun()
    elif sel_rows and file_bytes is None:
        st.info("ℹ️ Rincian harian hanya tersedia saat file Excel diupload langsung.")

    # ── Dialog Dispatch ──
    if st.session_state.dialog_target == "logic":
        st.session_state.dialog_target = None
        show_logic_dialog()
    elif st.session_state.dialog_target == "detail" and st.session_state.dialog_emp and file_bytes is not None:
        emp_s = st.session_state.dialog_emp
        st.session_state.dialog_target = None
        show_daily_detail(
            account=emp_s["account"],
            nama=emp_s["nama"],
            rules=emp_s["rules"],
            file_bytes=file_bytes,
        )

    st.caption(
        f"📊 Menampilkan {len(df_show):,} dari {len(df_result):,} karyawan  |  "
        f"📋 Total baris diproses: {stats['total_rows']:,}  |  "
        f"✅ Diklasifikasikan: {stats['classified']:,}  |  "
        f"⏭️ Dilewati (Rest/dll): {stats['skipped']:,}"
    )

    # ── Download ──
    st.markdown("---")
    dcol1, dcol2, dcol3 = st.columns([1, 1, 2])

    try:
        buf = io.BytesIO(file_bytes)
        raw = pd.read_excel(buf, sheet_name="General statistics and attendan",
                            header=None, nrows=2)
        tr_text = str(raw.iloc[1, 0])
        m = re.search(r'Time Range[:\s]*([\d/\u2013\-\s]+)', tr_text)
        time_range = m.group(1).strip() if m else ""
    except Exception:
        time_range = ""

    fname = (
        "Rekap_Absensi_" + time_range.replace(" ", "_").replace("\u2013", "sd") + ".xlsx"
        if time_range else "Rekap_Absensi.xlsx"
    )

    with dcol1:
        xlsx_bytes = to_excel_bytes(df_result, time_range)
        st.download_button(
            label="📥 Download Rekap Lengkap (.xlsx)",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    with dcol2:
        if len(df_show) < len(df_result):
            xlsx_filtered = to_excel_bytes(df_show, time_range)
            st.download_button(
                label="📥 Download Hasil Filter (.xlsx)",
                data=xlsx_filtered,
                file_name=f"Filter_{fname}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )

    # ── Ringkasan per Rules ──
    with st.expander("📊 Ringkasan per Rules"):
        grp = df_result.groupby("Rules").agg(
            Karyawan=("Account", "count"),
            Normal=("Normal", "sum"),
            Late=("Late", "sum"),
            **{"1/2 UL": ("1/2 UL", "sum")},
            DW=("DW", "sum"),
            K=("K", "sum"),
            AL=("AL", "sum"),
            WFA=("WFA", "sum"),
            Off=("Off", "sum"),
        ).reset_index().sort_values("1/2 UL", ascending=False)
        total_absen = grp["Normal"] + grp["Late"] + grp["1/2 UL"]
        grp["Late Rate"] = (grp["Late"]   / total_absen.replace(0, 1) * 100).round(1)
        grp["1/2 UL Rate"] = (grp["1/2 UL"] / total_absen.replace(0, 1) * 100).round(1)
        st.dataframe(
            grp,
            width="stretch",
            hide_index=True,
            column_config={
                "Rules"      : st.column_config.TextColumn("Rules"),
                "Karyawan"   : st.column_config.NumberColumn("👥 Karyawan", format="%d"),
                "Normal"     : st.column_config.NumberColumn("✅ Normal", format="%d"),
                "Late"       : st.column_config.NumberColumn("🕐 Late", format="%d"),
                "1/2 UL"     : st.column_config.NumberColumn("🚫 1/2 UL", format="%d"),
                "DW"         : st.column_config.NumberColumn("🚷 DW", format="%d"),
                "K"          : st.column_config.NumberColumn("🩺 K", format="%d"),
                "AL"         : st.column_config.NumberColumn("🌴 AL", format="%d"),
                "WFA"        : st.column_config.NumberColumn("🏠 WFA", format="%d"),
                "Off"        : st.column_config.NumberColumn("🔕 Off", format="%d"),
                "Late Rate"  : st.column_config.NumberColumn("% Late", format="%.1f%%"),
                "1/2 UL Rate": st.column_config.NumberColumn("% 1/2 UL", format="%.1f%%"),
            },
        )

else:
    st.markdown(
        '<div style="text-align:center;padding:3rem 2rem;color:#94a3b8;">'
        '<div style="font-size:3.5rem;margin-bottom:1rem;">📁</div>'
        '<div style="font-size:1.1rem;font-weight:600;color:#64748b;margin-bottom:0.5rem;">'
        'Belum ada file yang diupload</div>'
        '<div style="font-size:0.9rem;">Upload file <code>.xlsx</code> absensi di atas untuk memulai</div>'
        '</div>',
        unsafe_allow_html=True,
    )