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
# ──────────────────────────────────────────────────────────────
# Custom CSS
# ──────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.main .block-container {
    padding: 2rem 3rem 4rem;
    max-width: 1300px;
}

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

.metric-row {
    display: flex;
    gap: 1rem;
    margin: 1.5rem 0;
    flex-wrap: wrap;
}
.metric-card {
    flex: 1;
    min-width: 120px;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    font-weight: 600;
}
.metric-normal  { background: #f0fdf4; border-left: 4px solid #22c55e; }
.metric-late    { background: #fffbeb; border-left: 4px solid #f59e0b; }
.metric-k       { background: #fef2f2; border-left: 4px solid #ef4444; }
.metric-total   { background: #eff6ff; border-left: 4px solid #3b82f6; }
.metric-al      { background: #fdf4ff; border-left: 4px solid #a855f7; }
.metric-half-al { background: #fff1f2; border-left: 4px solid #fb7185; }
.metric-wfa     { background: #f0f9ff; border-left: 4px solid #0ea5e9; }

.metric-card .label { font-size: 0.78rem; color: #64748b; font-weight: 500; text-transform: uppercase; letter-spacing: 0.07em; }
.metric-card .value { font-size: 2rem; font-weight: 700; margin-top: 0.2rem; font-family: 'DM Mono', monospace; }
.metric-normal  .value { color: #16a34a; }
.metric-late    .value { color: #d97706; }
.metric-k       .value { color: #dc2626; }
.metric-total   .value { color: #2563eb; }
.metric-al      .value { color: #9333ea; }
.metric-half-al .value { color: #e11d48; }
.metric-wfa     .value { color: #0284c7; }

.rules-pill {
    display: inline-block;
    background: #e0f2fe;
    color: #0369a1;
    padding: 0.15rem 0.6rem;
    border-radius: 999px;
    font-size: 0.78rem;
    font-family: 'DM Mono', monospace;
}

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

.info-box {
    background: #eff6ff;
    border-radius: 10px;
    padding: 1rem 1.3rem;
    font-size: 0.88rem;
    color: #1e40af;
    border: 1px solid #bfdbfe;
    line-height: 1.7;
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

        # classify() sekarang return list atau None
        _klas_raw = classify(
            r["Earliest"], r["Shift"], r["Attendance results"],
            latest_raw=r["Latest"],
            leave_app=r.get("Leave & Overtime Application"),
        )
        # Display: gabung dengan " / " jika lebih dari satu status
        _klas_display = " / ".join(_klas_raw) if _klas_raw else None

        rows.append({
            "Tanggal"        : r["Tanggal"],
            "Shift"          : str(r["Shift"]).strip(),
            "Tipe"           : tipe,
            "Jam Masuk"      : str(r["Earliest"]).strip() if pd.notna(r["Earliest"]) else "--",
            "Jam Keluar"     : str(r["Latest"]).strip()   if pd.notna(r["Latest"])   else "--",
            "Status"         : str(r["Attendance results"]).strip() if pd.notna(r["Attendance results"]) else "--",
            "Jam Kerja"      : r["Jam Kerja"],
            "Klasifikasi"    : _klas_display,   # string untuk tampilan
            "Klasifikasi_raw": _klas_raw,        # list untuk filter
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
    with st.spinner("Memuat rincian harian…"):
        detail_df, summary_df = get_employee_daily(file_bytes, account)

    st.markdown(f"""
<div style="background:#f8fafc;border-radius:10px;padding:1rem 1.4rem;
            border-left:4px solid #3b82f6;margin-bottom:1.2rem;">
  <div style="font-size:1.05rem;font-weight:700;color:#1e293b;">{nama}</div>
  <div style="font-size:0.82rem;color:#64748b;margin-top:0.2rem;">
    <code>{account}</code> &nbsp;·&nbsp; {rules}
  </div>
</div>
""", unsafe_allow_html=True)

    tipe_cfg = {
        "S1": ("🌅", "S1",  "#f0fdf4", "#22c55e", "#166534"),
        "S2": ("🌙", "S2",  "#faf5ff", "#a855f7", "#6b21a8"),
        "H" : ("🏖️", "H",   "#fff7ed", "#fb923c", "#9a3412"),
    }
    cols = st.columns(3)
    for i, tipe in enumerate(["S1", "S2", "H"]):
        row = summary_df[summary_df["Tipe"] == tipe]
        hari = int(row["Hari"].values[0])        if len(row) else 0
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
  <div style="font-size:0.8rem;color:#64748b;margin-top:0.1rem;">{jam:.1f} jam kerja</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:1.2rem'></div>", unsafe_allow_html=True)

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

    # Emoji map untuk semua status
    KLAS_EMOJI = {
        "Normal"   : "✅ Normal",
        "Late"     : "🟡 Late",
        "1/2 UL"   : "🔴 1/2 UL",
        "AL"       : "🟣 AL",
        "1/2 AL"   : "🩷 ½AL",
        "WFA"      : "🔵 WFA",
    }

    # ── Filter pakai Klasifikasi_raw (list) ──────────────────
    late_df = detail_df[detail_df["Klasifikasi_raw"].apply(
        lambda x: has_status(x, "Late")
    )].copy()
    k_df    = detail_df[detail_df["Klasifikasi_raw"].apply(
        lambda x: has_status(x, "1/2 UL")
    )].copy()
    al_df   = detail_df[detail_df["Klasifikasi_raw"].apply(
        lambda x: has_status(x, "AL") or has_status(x, "1/2 AL")
    )].copy()
    wfa_df  = detail_df[detail_df["Klasifikasi_raw"].apply(
        lambda x: has_status(x, "WFA")
    )].copy()

    n_late = len(late_df)
    n_k    = len(k_df)
    n_al   = len(al_df)
    n_wfa  = len(wfa_df)

    # ── Expander 1: Keterlambatan ──
    late_label = f"⚠️ Rincian Keterlambatan  —  🟡 {n_late} Late  ·  🔴 {n_k} ½UL"
    with st.expander(late_label, expanded=False):
        if n_late == 0 and n_k == 0:
            st.success("✅ Tidak ada keterlambatan pada periode ini.")
        else:
            mc1, mc2 = st.columns(2)
            with mc1:
                st.markdown(f"""
<div style="background:#fffbeb;border-left:4px solid #f59e0b;border-radius:10px;
            padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">
  <div style="font-size:0.72rem;color:#92400e;font-weight:600;text-transform:uppercase;">🟡 Late</div>
  <div style="font-size:1.7rem;font-weight:700;color:#d97706;font-family:'DM Mono',monospace;">
    {n_late}<span style="font-size:.9rem"> hari</span></div>
</div>""", unsafe_allow_html=True)
            with mc2:
                st.markdown(f"""
<div style="background:#fef2f2;border-left:4px solid #ef4444;border-radius:10px;
            padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">
  <div style="font-size:0.72rem;color:#991b1b;font-weight:600;text-transform:uppercase;">🔴 ½UL (&gt;2 jam)</div>
  <div style="font-size:1.7rem;font-weight:700;color:#dc2626;font-family:'DM Mono',monospace;">
    {n_k}<span style="font-size:.9rem"> hari</span></div>
</div>""", unsafe_allow_html=True)

            combined = pd.concat([late_df, k_df]).sort_values("Tanggal").reset_index(drop=True)
            combined["No."]       = range(1, len(combined) + 1)
            combined["Terlambat"] = combined.apply(_menit_terlambat, axis=1)
            combined["Klasifikasi"] = combined["Klasifikasi"].fillna("—")
            st.dataframe(
                combined[["No.", "Tanggal", "Klasifikasi", "Shift", "Jam Masuk", "Jam Keluar", "Terlambat"]],
                use_container_width=True,
                height=min(60 + len(combined) * 35, 380),
                hide_index=True,
            )

    # ── Expander 2: Leave (AL / ½AL / WFA) ──
    leave_label = f"🏖️ Rincian Leave  —  🟣 AL: {n_al}  ·  🔵 WFA: {n_wfa}"
    with st.expander(leave_label, expanded=False):
        if n_al == 0 and n_wfa == 0:
            st.info("Tidak ada data AL / ½AL / WFA pada periode ini.")
        else:
            lc1, lc2, lc3 = st.columns(3)
            n_full_al = len(detail_df[detail_df["Klasifikasi_raw"].apply(
                lambda x: has_status(x, "AL")
            )])
            n_half_al = len(detail_df[detail_df["Klasifikasi_raw"].apply(
                lambda x: has_status(x, "1/2 AL")
            )])
            for (col, label, val, bg, bc, tc) in [
                (lc1, "🟣 AL (Full)",      n_full_al, "#fdf4ff", "#a855f7", "#7e22ce"),
                (lc2, "🩷 ½AL (Setengah)", n_half_al, "#fff1f2", "#fb7185", "#be123c"),
                (lc3, "🔵 WFA",            n_wfa,     "#f0f9ff", "#0ea5e9", "#0369a1"),
            ]:
                with col:
                    st.markdown(f"""
<div style="background:{bg};border-left:4px solid {bc};border-radius:10px;
            padding:.8rem 1.1rem;text-align:center;margin-bottom:.8rem;">
  <div style="font-size:0.72rem;color:{tc};font-weight:600;text-transform:uppercase;">{label}</div>
  <div style="font-size:1.7rem;font-weight:700;color:{tc};font-family:'DM Mono',monospace;">
    {val}<span style="font-size:.9rem"> hari</span></div>
</div>""", unsafe_allow_html=True)

            leave_combined = pd.concat([al_df, wfa_df]).sort_values("Tanggal").reset_index(drop=True)
            # Deduplikasi: satu hari bisa muncul di al_df dan wfa_df sekaligus (jika dual-count)
            leave_combined = leave_combined.drop_duplicates(subset=["Tanggal"]).reset_index(drop=True)
            leave_combined["No."] = range(1, len(leave_combined) + 1)
            st.dataframe(
                leave_combined[["No.", "Tanggal", "Klasifikasi", "Shift", "Jam Masuk", "Jam Keluar", "Status"]],
                use_container_width=True,
                height=min(60 + len(leave_combined) * 35, 380),
                hide_index=True,
                column_config={
                    "Status": st.column_config.TextColumn("Attendance Results", width="large"),
                },
            )

    # ── Expander 3: Rincian per Tipe Shift ──
    tipe_counts = detail_df["Tipe"].value_counts()
    s1_n = tipe_counts.get("S1", 0)
    s2_n = tipe_counts.get("S2", 0)
    h_n  = tipe_counts.get("H",  0)
    with st.expander(
        f"📅 Rincian per Tipe Shift  —  🌅 S1: {s1_n}  ·  🌙 S2: {s2_n}  ·  🏖️ H: {h_n}",
        expanded=False
    ):
        tab_s1, tab_s2, tab_h = st.tabs(["🌅 S1", "🌙 S2", "🏖️ H"])

        def _render_tipe_tab(tipe_key):
            df_tipe = detail_df[detail_df["Tipe"] == tipe_key].copy().reset_index(drop=True)
            if df_tipe.empty:
                st.info(f"Tidak ada data tipe {tipe_key} pada periode ini.")
                return
            df_tipe["No."]       = range(1, len(df_tipe) + 1)
            df_tipe["Jam Kerja"] = df_tipe["Jam Kerja"].apply(lambda x: f"{x:.1f} jam" if x > 0 else "—")
            total_jam = detail_df[detail_df["Tipe"] == tipe_key]["Jam Kerja"].sum()
            st.caption(f"{len(df_tipe)} hari  ·  Total jam kerja: {total_jam:.1f} jam")
            st.dataframe(
                df_tipe[["No.", "Tanggal", "Shift", "Jam Masuk", "Jam Keluar", "Klasifikasi", "Jam Kerja"]],
                use_container_width=True,
                height=min(60 + len(df_tipe) * 35, 420),
                hide_index=True,
            )

        with tab_s1: _render_tipe_tab("S1")
        with tab_s2: _render_tipe_tab("S2")
        with tab_h:  _render_tipe_tab("H")

    # ── Expander 4: Detail Lengkap ──
    with st.expander(f"🗂️ Detail Lengkap per Hari  —  {len(detail_df)} hari tercatat", expanded=False):
        TIPE_EMOJI = {"S1": "🌅 S1", "S2": "🌙 S2", "H": "🏖️ H"}
        dd = detail_df.copy()
        dd["Tipe"]      = dd["Tipe"].map(lambda x: TIPE_EMOJI.get(x, x))
        dd["Jam Kerja"] = dd["Jam Kerja"].apply(lambda x: f"{x:.1f} jam" if x > 0 else "—")
        st.dataframe(
            dd[["No.", "Tanggal", "Tipe", "Shift", "Jam Masuk", "Jam Keluar", "Status", "Klasifikasi", "Jam Kerja"]],
            use_container_width=True,
            height=420,
            hide_index=True,
            column_config={
                "Status": st.column_config.TextColumn("Status Absensi", width="large"),
            },
        )

    st.caption("Klik di luar kotak ini untuk menutup")


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
        dtype={"Earliest": str, "Latest": str}
    )

    required = ["Name", "Account", "Rules", "Shift", "Earliest", "Latest",
                "Attendance results", "Leave & Overtime Application"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom tidak ditemukan: {missing}")

    df = df[required].dropna(subset=["Account", "Rules"])
    df = df[~df["Account"].astype(str).str.strip().isin(["", "--"])]

    # classify() sekarang return list — simpan ke _statuses
    df["_statuses"] = df.apply(
        lambda r: classify(
            r["Earliest"], r["Shift"], r["Attendance results"],
            latest_raw=r["Latest"],
            leave_app=r["Leave & Overtime Application"],
        ),
        axis=1,
    )

    # Rules: ambil nilai terbanyak per Account
    all_employees = (
        df.groupby("Account")["Rules"]
        .agg(lambda x: x.mode()[0])
        .reset_index()
    )

    # Baris yang punya klasifikasi (sebelum explode, untuk stats)
    df_classified = df[df["_statuses"].notna()].copy()

    # Explode list → setiap status jadi baris sendiri, lalu pivot
    df_exploded = df_classified.explode("_statuses").rename(columns={"_statuses": "Status"})

    pivot = df_exploded.pivot_table(
        index="Account",
        columns="Status",
        values="Shift",
        aggfunc="count",
        fill_value=0,
    ).reset_index()
    pivot.columns.name = None

    # Pastikan semua kolom status ada
    for col in ["Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA"]:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot = all_employees.merge(pivot, on="Account", how="left").fillna(0)
    for col in ["Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA"]:
        pivot[col] = pivot[col].astype(int)

    name_map = df.groupby("Account")["Name"].first()
    pivot["Nama"] = pivot["Account"].map(name_map)

    pivot = pivot.sort_values(["Rules", "Nama"]).reset_index(drop=True)
    pivot.insert(0, "No.", range(1, len(pivot) + 1))
    result = pivot[["No.", "Nama", "Account", "Rules",
                    "Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA"]].copy()

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

    thin   = Side(style="thin", color="CBD5E1")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center")
    LEFT   = Alignment(horizontal="left",   vertical="center")

    # Title
    ws.merge_cells("A1:J1")
    ws["A1"] = "Summary Attendance"
    ws["A1"].font      = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    ws["A1"].fill      = DARK
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 30

    # Subtitle
    ws.merge_cells("A2:J2")
    ws["A2"] = f"Time Range: {time_range}" if time_range else "Rekap Absensi"
    ws["A2"].font      = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    ws["A2"].fill      = BLUE
    ws["A2"].alignment = CENTER
    ws.row_dimensions[2].height = 16

    # Header
    headers = ["No.", "Nama", "Account", "Rules",
               "Normal", "Late", "1/2 UL", "AL", "1/2 AL", "WFA"]
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
    }

    for ri, (_, row) in enumerate(df.iterrows()):
        er = ri + 4
        base_fill = ALT if ri % 2 == 0 else WHITE
        vals = [row["No."], row["Nama"], row["Account"], row["Rules"],
                row["Normal"], row["Late"], row["1/2 UL"],
                row["AL"], row["1/2 AL"], row["WFA"]]
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

    # Total row
    tr = len(df) + 4
    ws.merge_cells(f"A{tr}:D{tr}")
    ws[f"A{tr}"]           = "TOTAL"
    ws[f"A{tr}"].font      = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    ws[f"A{tr}"].fill      = DARK
    ws[f"A{tr}"].alignment = CENTER
    ws[f"A{tr}"].border    = BORDER
    for ci in range(5, 11):
        c       = ws.cell(row=tr, column=ci)
        c.value = f"=SUM({get_column_letter(ci)}4:{get_column_letter(ci)}{tr-1})"
        c.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        c.fill  = DARK
        c.alignment = CENTER
        c.border    = BORDER

    for ci, w in enumerate([6, 32, 24, 28, 10, 10, 8, 8, 8, 8], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes    = "A4"
    ws.auto_filter.ref = f"A3:J{tr-1}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────

st.markdown("""
<div class="app-header">
  <div class="badge">HR TOOLS</div>
  <h1>🗓️ Absensi Rekap Generator</h1>
  <p>Upload file Excel absensi → Hitung Normal / Late / K / AL / WFA per karyawan → Download hasil</p>
</div>
""", unsafe_allow_html=True)

col_up, col_info = st.columns([2, 1], gap="large")

with col_up:
    st.markdown('<p class="section-title">📂 Upload File Excel</p>', unsafe_allow_html=True)
    periodes_tersedia = get_periodes()
    if periodes_tersedia:
        st.markdown("**📅 Atau pilih periode yang sudah tersimpan:**")
        periode_dipilih = st.selectbox(
            label="", options=["— Upload file baru —"] + periodes_tersedia,
            label_visibility="collapsed"
        )
    else:
        periode_dipilih = "— Upload file baru —"

    uploaded = st.file_uploader(
        label="", type=["xlsx", "xls"], label_visibility="collapsed",
        help="File Excel dari sistem absensi. Sheet 'General statistics and attendan' harus ada."
    )

with col_info:
    st.markdown("""
<div class="info-box">
<strong>Logic Klasifikasi:</strong><br>
✅ <strong>Normal</strong> — tepat waktu / att results mengandung "Normal"<br>
🟡 <strong>Late</strong> — terlambat ≤ 2 jam<br>
🔴 <strong>½UL</strong> — terlambat &gt; 2 jam<br>
🟣 <strong>AL</strong> — Annual Leave (tanpa punch)<br>
🩷 <strong>½AL</strong> — Annual Leave (ada punch in &amp; out)<br>
🔵 <strong>WFA</strong> — Work From Home (ada punch in &amp; out)<br><br>
<strong>Dual-count:</strong><br>
• Normal (Leave) + WFH + punch → Normal <em>dan</em> WFA<br>
• Normal (Leave) + AnnualLeave → Normal <em>dan</em> AL/½AL<br><br>
<strong>Catatan:</strong><br>
• WFH tanpa punch → Normal saja<br>
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
            # classify_str() → join list jadi string untuk DB
            df_raw["_status_klasifikasi"] = df_raw.apply(
                lambda r: classify_str(
                    r["Earliest"], r["Shift"], r["Attendance results"],
                    latest_raw=r["Latest"],
                    leave_app=r.get("Leave & Overtime Application"),
                ), axis=1
            )
            save_periode(df_raw, _periode)
        except Exception as e:
            st.warning(f"⚠️ Gagal simpan ke database: {e}")
    else:
        df_result = get_rekap(periode_dipilih)
        df_result = df_result.rename(columns={
            "nama": "Nama", "account": "Account", "rules": "Rules",
            "normal": "Normal", "late": "Late", "k": "1/2 UL",
        })
        for col in ["AL", "1/2 AL", "WFA", "1/2 UL"]:
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
    total_e   = stats["employees"]

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
    <div class="label">½ UL (&gt;2 jam)</div>
    <div class="value">{total_k:,}</div>
  </div>
  <div class="metric-card metric-al">
    <div class="label">Annual Leave (AL)</div>
    <div class="value">{total_al:,}</div>
  </div>
  <div class="metric-card metric-half-al">
    <div class="label">½ Annual Leave</div>
    <div class="value">{total_hal:,}</div>
  </div>
  <div class="metric-card metric-wfa">
    <div class="label">WFA (Work from Home)</div>
    <div class="value">{total_wfa:,}</div>
  </div>
  <div class="metric-card metric-total">
    <div class="label">Jumlah Karyawan</div>
    <div class="value">{total_e:,}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Filter & Tabel ──
    st.markdown('<p class="section-title">📊 Hasil Summary per Karyawan</p>', unsafe_allow_html=True)

    fcol1, fcol2, fcol3 = st.columns([2, 2, 1])
    with fcol1:
        all_rules = sorted(df_result["Rules"].unique().tolist())
        sel_rules = st.multiselect("Filter Rules", options=all_rules, placeholder="Semua Rules")
    with fcol2:
        search = st.text_input("Cari Nama / Account", placeholder="Ketik nama atau account…")
    with fcol3:
        show_late_only = st.checkbox("Hanya Late/K", value=False)

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
        df_show = df_show[(df_show["Late"] > 0) | (df_show["1/2 UL"] > 0)]

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
            "No."    : st.column_config.NumberColumn("No.", width="small"),
            "Nama"   : st.column_config.TextColumn("Nama", width="large"),
            "Account": st.column_config.TextColumn("Account", width="medium"),
            "Rules"  : st.column_config.TextColumn("Rules", width="medium"),
            "Normal" : st.column_config.NumberColumn("Normal ✅", format="%d", width="small"),
            "Late"   : st.column_config.NumberColumn("Late 🟡", format="%d", width="small"),
            "1/2 UL" : st.column_config.NumberColumn("½UL 🔴", format="%d", width="small"),
            "AL"     : st.column_config.NumberColumn("AL 🟣", format="%d", width="small"),
            "1/2 AL" : st.column_config.NumberColumn("½AL 🩷", format="%d", width="small"),
            "WFA"    : st.column_config.NumberColumn("WFA 🔵", format="%d", width="small"),
        },
    )

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
               f"Dilewati (Rest/dll): {stats['skipped']:,}")

    # ── Download ──
    st.markdown("---")
    dcol1, dcol2, dcol3 = st.columns([1, 1, 2])

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
        fname = (f"Rekap_Absensi_{time_range.replace(' ','_').replace('–','s/d')}.xlsx"
                 if time_range else "Rekap_Absensi.xlsx")
        st.download_button(
            label="⬇️  Download Rekap Lengkap (.xlsx)",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dcol2:
        if len(df_show) < len(df_result):
            xlsx_filtered = to_excel_bytes(df_show, time_range)
            st.download_button(
                label="⬇️  Download Hasil Filter (.xlsx)",
                data=xlsx_filtered,
                file_name=f"Rekap_Filter_{fname}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # ── Ringkasan per Rules ──
    with st.expander("📈 Ringkasan per Rules"):
        grp = df_result.groupby("Rules").agg(
            Karyawan=("Account", "count"),
            Normal=("Normal", "sum"),
            Late=("Late", "sum"),
            **{"1/2 UL": ("1/2 UL", "sum")},
            AL=("AL", "sum"),
            WFA=("WFA", "sum"),
        ).reset_index().sort_values("1/2 UL", ascending=False)
        total_absen = grp["Normal"] + grp["Late"] + grp["1/2 UL"]
        grp["Late Rate"] = (grp["Late"]   / total_absen * 100).round(1)
        grp["½UL Rate"]  = (grp["1/2 UL"] / total_absen * 100).round(1)
        st.dataframe(
            grp,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Rules"    : st.column_config.TextColumn("Rules"),
                "Karyawan" : st.column_config.NumberColumn("Karyawan", format="%d"),
                "Normal"   : st.column_config.NumberColumn("Normal ✅", format="%d"),
                "Late"     : st.column_config.NumberColumn("Late 🟡", format="%d"),
                "1/2 UL"   : st.column_config.NumberColumn("½UL 🔴", format="%d"),
                "AL"       : st.column_config.NumberColumn("AL 🟣", format="%d"),
                "WFA"      : st.column_config.NumberColumn("WFA 🔵", format="%d"),
                "Late Rate": st.column_config.NumberColumn("% Late", format="%.1f%%"),
                "½UL Rate" : st.column_config.NumberColumn("% ½UL", format="%.1f%%"),
            },
        )

else:
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