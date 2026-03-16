import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.gridspec import GridSpec
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
import io
from datetime import datetime, date
from supabase import create_client, Client

# ── Supabase Config ───────────────────────────────────────────────────────────
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]

@st.cache_resource
def get_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

# ── Lists ─────────────────────────────────────────────────────────────────────
PROCESSES = ["Pretest", "Adjustment", "Minirel", "Verification"]

FAILURE_TYPES = [
    "DigitalInitializeCheck", "ABUS", "KDMI",
    "Adjust_ReferenceFrequency", "Adjust_IFRangingGain",
    "Adjust_SourceInsertionLossRatio", "Adjust_Compression",
    "Adjust_WaveshaperIDAC", "Adjust_PreVirtualBridge",
    "Adjust_StepAttenuatorPort1", "Adjust_RchGain",
    "Adjust_StepAttenuatorOtherPorts", "Adjust_SourceLevelDAC",
    "Adjust_VirtualBridge", "Adjust_SpectrumAnalyzer_Gain",
    "Adjust_IFFlatness", "IFFlatness", "PowerLinearityWithPowerSweep",
    "PowerLevelTransient", "TraceNoiseThrough", "Harmonics",
    "UnratioedPowerMeasurement", "SourceOnOffRatio", "UncorrectedPortChar",
    "CompressionRch", "PulseProfile_PreCheck", "OutputLevelRange",
    "RampSweep", "NoiseFloor", "Crosstalk", "IFRanging",
    "TraceNoiseOpen", "CompressionOpen", "OutputAccuracyOddPorts",
    "LOLeakageOddPorts", "OutputAccuracyEvenPorts", "LOLeakageEvenPorts",
    "Compression12", "FinalCheckProductID", "DDR3_Reset", "DDR3_Function",
    "Adjust_WriteNominalFiles", "WritePCASerialNumber", "LED",
    "LoopCountIncrement", "UnratioedPowerMeasurementOpen",
    "NoiseFloorOpen", "DDR3", "SweepWithDDR3Tests", "LoopCountClose",
]

COLUMNS = ["Date", "Model", "Serial Number", "Process",
           "Failure Type", "Description", "Remark", "Photo Path"]

# ── Supabase helpers ──────────────────────────────────────────────────────────

def load_records() -> pd.DataFrame:
    """Load all records from Supabase."""
    try:
        supabase = get_supabase()
        res = supabase.table("failure_log") \
                      .select("*") \
                      .order("created_at", desc=False) \
                      .execute()
        if not res.data:
            return pd.DataFrame(columns=COLUMNS)
        df = pd.DataFrame(res.data)
        df = df.rename(columns={
            "date":          "Date",
            "model":         "Model",
            "serial_number": "Serial Number",
            "process":       "Process",
            "failure_type":  "Failure Type",
            "description":   "Description",
            "remark":        "Remark",
            "photo_path":    "Photo Path",
        })
        keep = ["id"] + COLUMNS + ["created_at"]
        df = df[[c for c in keep if c in df.columns]]
        df.fillna("", inplace=True)
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
        return df
    except Exception as e:
        st.error(f"Error loading records: {e}")
        return pd.DataFrame(columns=COLUMNS)


def insert_record(row: dict):
    """Insert one record into Supabase."""
    supabase = get_supabase()
    supabase.table("failure_log").insert({
        "date":          row.get("Date", ""),
        "model":         row.get("Model", ""),
        "serial_number": row.get("Serial Number", ""),
        "process":       row.get("Process", ""),
        "failure_type":  row.get("Failure Type", ""),
        "description":   row.get("Description", ""),
        "remark":        row.get("Remark", ""),
        "photo_path":    row.get("Photo Path", ""),
    }).execute()


def delete_record(record_id: int):
    """Delete a record by its Supabase ID."""
    supabase = get_supabase()
    supabase.table("failure_log").delete().eq("id", record_id).execute()


# ── Pareto helpers ────────────────────────────────────────────────────────────

def compute_pareto(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "Failure Type" not in df.columns:
        return pd.DataFrame()
    counts = df["Failure Type"].value_counts().reset_index()
    counts.columns = ["Failure Type", "Count"]
    total = counts["Count"].sum()
    counts["Percentage (%)"] = (counts["Count"] / total * 100).round(1)
    counts["Cumulative (%)"] = counts["Percentage (%)"].cumsum().round(1)
    return counts


def plot_pareto(pareto_df: pd.DataFrame, title: str = "Pareto Chart") -> plt.Figure:
    fig = plt.figure(figsize=(10, 5))
    ax1 = fig.add_subplot(111)
    ax2 = ax1.twinx()

    labels   = pareto_df["Failure Type"].tolist()
    counts   = pareto_df["Count"].tolist()
    cum_pct  = pareto_df["Cumulative (%)"].tolist()
    x = range(len(labels))

    bars = ax1.bar(x, counts, color="#185fa5", width=0.55, zorder=3, label="Count")
    ax1.set_xticks(list(x))
    ax1.set_xticklabels(labels, rotation=25, ha="right", fontsize=8)
    ax1.set_ylabel("Failure Count", fontsize=10)
    ax1.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax1.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
    ax1.set_axisbelow(True)

    for bar, v in zip(bars, counts):
        ax1.text(bar.get_x() + bar.get_width() / 2,
                 bar.get_height() + 0.2, str(v),
                 ha="center", va="bottom", fontsize=8, color="#185fa5")

    ax2.plot(list(x), cum_pct, color="#e24b4a", marker="o",
             markersize=5, linewidth=2, label="Cumulative %", zorder=4)
    ax2.axhline(80, color="#e24b4a", linestyle=":", linewidth=1, alpha=0.5)
    ax2.set_ylabel("Cumulative %", fontsize=10)
    ax2.set_ylim(0, 110)
    ax2.yaxis.set_major_formatter(mticker.PercentFormatter())

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2,
               loc="upper left", fontsize=9, framealpha=0.8)
    ax1.set_title(title, fontsize=12, fontweight="bold", pad=12)
    fig.tight_layout()
    return fig


def export_pareto_to_excel(pareto_df, start_date, end_date) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pareto_Report"

    ws.merge_cells("A1:F1")
    ws["A1"].value = "Failure Pareto Report"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:F2")
    s = str(start_date) if start_date else "All"
    e = str(end_date)   if end_date   else "All"
    ws["A2"].value = f"Date Range: {s}  →  {e}    |    Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A2"].font = Font(size=10, color="444444")
    ws["A2"].alignment = Alignment(horizontal="center")

    headers = ["#", "Failure Type", "Count", "Percentage (%)", "Cumulative (%)"]
    widths  = [6, 30, 10, 18, 18]
    hfill   = PatternFill("solid", fgColor="1F4E79")
    hfont   = Font(bold=True, color="FFFFFF", size=11)
    thin    = Side(style="thin", color="CCCCCC")
    border  = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=4, column=ci, value=h)
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w

    for i, row in pareto_df.iterrows():
        er = i + 5
        alt = PatternFill("solid", fgColor="EBF3FB") if er % 2 == 0 else None
        for ci, v in enumerate([i+1, row["Failure Type"], row["Count"],
                                  row["Percentage (%)"], row["Cumulative (%)"]], 1):
            cell = ws.cell(row=er, column=ci, value=v)
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
            if alt: cell.fill = alt
        if i == 0:
            for ci in range(1, 6):
                ws.cell(row=er, column=ci).fill = PatternFill("solid", fgColor="D6E8F5")
                ws.cell(row=er, column=ci).font = Font(bold=True)

    last_row = len(pareto_df) + 4
    bar = BarChart(); bar.type = "col"
    bar.title = "Failure Count by Type"
    bar.y_axis.title = "Count"; bar.x_axis.title = "Failure Type"
    bar.style = 10; bar.width = 18; bar.height = 12
    data = Reference(ws, min_col=3, min_row=4, max_row=last_row)
    cats = Reference(ws, min_col=2, min_row=5, max_row=last_row)
    bar.add_data(data, titles_from_data=True); bar.set_categories(cats)
    ws.add_chart(bar, "H4")

    line = LineChart(); line.title = "Cumulative %"
    line.y_axis.title = "Cumulative %"; line.style = 10
    line.width = 18; line.height = 12
    cum_data = Reference(ws, min_col=5, min_row=4, max_row=last_row)
    line.add_data(cum_data, titles_from_data=True); line.set_categories(cats)
    ws.add_chart(line, "H22")

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def export_full_log_to_excel(df: pd.DataFrame) -> bytes:
    """Export complete failure log to Excel."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Failure_Log"

    hfill  = PatternFill("solid", fgColor="1F4E79")
    hfont  = Font(bold=True, color="FFFFFF", size=11)
    thin   = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    col_widths = [14, 14, 16, 14, 28, 35, 20, 30]

    for ci, (col, w) in enumerate(zip(COLUMNS, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"

    sn_counts = df["Serial Number"].value_counts() if not df.empty else {}
    for ri, row_data in df.iterrows():
        er = ri + 2
        alt = PatternFill("solid", fgColor="EBF3FB") if er % 2 == 0 else None
        sn  = row_data.get("Serial Number", "")
        for ci, col in enumerate(COLUMNS, 1):
            cell = ws.cell(row=er, column=ci, value=str(row_data.get(col, "")))
            cell.alignment = Alignment(vertical="center")
            cell.border = border
            if alt: cell.fill = alt
        if sn and sn_counts.get(sn, 1) > 1:
            ws.cell(row=er, column=3).fill = PatternFill("solid", fgColor="FFD966")

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Failure Tracking System",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
  [data-testid="stSidebar"] { background: #1e2a3a; }
  [data-testid="stSidebar"] * { color: #e0e8f0 !important; }
  .block-container { padding-top: 1.5rem; }
  .metric-card { background:#EBF3FB;border-radius:10px;padding:14px 18px;text-align:center; }
  .metric-card .val { font-size:2rem;font-weight:600;color:#185fa5; }
  .metric-card .lbl { font-size:.78rem;color:#555;margin-top:2px; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔬 Failure Tracker")
    st.markdown("*Manufacturing QC System*")
    st.divider()
    page = st.radio("Navigation", ["📋 Data Entry", "📊 Pareto Analysis", "📁 Failure Log"],
                    label_visibility="collapsed")
    st.divider()
    df_sidebar = load_records()
    st.markdown(f"**Total Records:** {len(df_sidebar)}")
    if not df_sidebar.empty and "Failure Type" in df_sidebar.columns:
        top = df_sidebar["Failure Type"].value_counts().idxmax()
        st.markdown(f"**Top Failure:** {top}")
    st.divider()
    if not df_sidebar.empty:
        excel_bytes = export_full_log_to_excel(df_sidebar[COLUMNS] if all(c in df_sidebar.columns for c in COLUMNS) else df_sidebar)
        st.download_button(
            "⬇️ Download Full Excel Log",
            data=excel_bytes,
            file_name=f"failure_log_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ════════════════════════════════════════════════════════════════════════════
# PAGE 1 — DATA ENTRY
# ════════════════════════════════════════════════════════════════════════════
if page == "📋 Data Entry":
    st.title("📋 Record New Failure")

    with st.form("entry_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            f_date    = st.date_input("Date *", value=date.today())
            f_model   = st.text_input("Model *", placeholder="e.g. ABC123")
            f_sn      = st.text_input("Serial Number *", placeholder="e.g. SN00981")
            f_process = st.selectbox("Process *", [""] + PROCESSES)
        with c2:
            f_type   = st.selectbox("Failure Type *", [""] + FAILURE_TYPES)
            f_remark = st.text_input("Remark", placeholder="e.g. Reworked, Scrapped…")
            f_photo  = st.file_uploader("Photo (optional)", type=["jpg","jpeg","png","bmp","tiff"])

        f_desc = st.text_area("Description", placeholder="Describe the failure in detail…", height=80)
        submitted = st.form_submit_button("💾 Save Record", use_container_width=True, type="primary")

    if submitted:
        errors = []
        if not f_model.strip():  errors.append("Model")
        if not f_sn.strip():     errors.append("Serial Number")
        if not f_process:        errors.append("Process")
        if not f_type:           errors.append("Failure Type")

        if errors:
            st.error(f"Required fields missing: {', '.join(errors)}")
        else:
            photo_path = f_photo.name if f_photo else ""
            try:
                insert_record({
                    "Date":          str(f_date),
                    "Model":         f_model.strip(),
                    "Serial Number": f_sn.strip(),
                    "Process":       f_process,
                    "Failure Type":  f_type,
                    "Description":   f_desc.strip(),
                    "Remark":        f_remark.strip(),
                    "Photo Path":    photo_path,
                })
                st.success(f"✅ Record saved! Serial **{f_sn.strip()}** logged for {f_type}.")
                st.cache_resource.clear()
                st.rerun()
            except Exception as e:
                st.error(f"Failed to save: {e}")

    df_all = load_records()
    if not df_all.empty:
        st.divider()
        st.subheader("Recent Entries")
        show_cols = [c for c in COLUMNS if c in df_all.columns]
        st.dataframe(df_all[show_cols].tail(10).iloc[::-1].reset_index(drop=True),
                     use_container_width=True, height=260)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 2 — PARETO ANALYSIS
# ════════════════════════════════════════════════════════════════════════════
elif page == "📊 Pareto Analysis":
    st.title("📊 Pareto Analysis")
    df_all = load_records()

    with st.container(border=True):
        fc1, fc2, fc3 = st.columns([2, 2, 1])
        with fc1: start_date = st.date_input("Start Date", value=None, key="d_start")
        with fc2: end_date   = st.date_input("End Date",   value=None, key="d_end")
        with fc3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Clear Filter", use_container_width=True):
                st.session_state.d_start = None
                st.session_state.d_end   = None
                st.rerun()

    df = df_all.copy()
    if not df.empty and "Date" in df.columns:
        if start_date: df = df[df["Date"] >= start_date]
        if end_date:   df = df[df["Date"] <= end_date]

    pareto_df = compute_pareto(df)
    total   = len(df)
    n_types = len(pareto_df)
    top_type = pareto_df.iloc[0]["Failure Type"] if not pareto_df.empty else "—"
    top_pct  = pareto_df.iloc[0]["Percentage (%)"] if not pareto_df.empty else 0
    pareto80 = 0
    if not pareto_df.empty:
        for _, row in pareto_df.iterrows():
            pareto80 += 1
            if row["Cumulative (%)"] >= 80: break

    m1, m2, m3, m4 = st.columns(4)
    for col, val, lbl, sub in [
        (m1, total,         "Total Failures",   "in selected range"),
        (m2, n_types,       "Failure Types",    "distinct types"),
        (m3, f"{top_pct}%", f"Top: {top_type}", "of total failures"),
        (m4, pareto80,      "Types → 80% rule", "drive 80% of failures"),
    ]:
        col.markdown(
            f'<div class="metric-card"><div class="val">{val}</div>'
            f'<div class="lbl">{lbl}<br><span style="font-size:.7rem">{sub}</span></div></div>',
            unsafe_allow_html=True)

    st.markdown("")
    if pareto_df.empty:
        st.info("No data found for the selected date range.")
    else:
        with st.container(border=True):
            st.subheader("Pareto Chart")
            title_str = "Pareto Chart"
            if start_date or end_date:
                s = str(start_date) if start_date else "All"
                e = str(end_date)   if end_date   else "All"
                title_str = f"Pareto Chart  ({s} → {e})"
            fig = plot_pareto(pareto_df, title=title_str)
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)

        with st.container(border=True):
            st.subheader("Summary Table")
            display_df = pareto_df.copy()
            display_df.insert(0, "#", range(1, len(display_df)+1))
            st.dataframe(display_df, use_container_width=True, hide_index=True)

        excel_bytes = export_pareto_to_excel(pareto_df, start_date, end_date)
        fname = f"pareto_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button("⬇️ Export Pareto Report to Excel",
                           data=excel_bytes, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 3 — FAILURE LOG
# ════════════════════════════════════════════════════════════════════════════
elif page == "📁 Failure Log":
    st.title("📁 Failure Log")
    df_all = load_records()

    if df_all.empty:
        st.info("No records yet. Use Data Entry to log failures.")
    else:
        with st.container(border=True):
            sc1, sc2, sc3 = st.columns(3)
            search_sn   = sc1.text_input("Search Serial Number", placeholder="e.g. SN001")
            filter_type = sc2.selectbox("Filter by Failure Type", ["All"] + FAILURE_TYPES)
            filter_proc = sc3.selectbox("Filter by Process",      ["All"] + PROCESSES)

        df_view = df_all.copy()
        if search_sn:
            df_view = df_view[df_view["Serial Number"].str.contains(search_sn, case=False, na=False)]
        if filter_type != "All":
            df_view = df_view[df_view["Failure Type"] == filter_type]
        if filter_proc != "All":
            df_view = df_view[df_view["Process"] == filter_proc]

        st.markdown(f"**{len(df_view)}** records shown (of {len(df_all)} total)")

        sn_counts = df_all["Serial Number"].value_counts()
        df_disp = df_view.copy().reset_index(drop=True)
        df_disp["Repeat Unit?"] = df_disp["Serial Number"].apply(
            lambda sn: f"⚠️ ×{sn_counts[sn]}" if sn_counts.get(sn, 1) > 1 else "✅ New"
        )

        show_cols = [c for c in COLUMNS if c in df_disp.columns] + ["Repeat Unit?"]
        st.dataframe(df_disp[show_cols], use_container_width=True, height=420,
                     column_config={
                         "Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
                         "Repeat Unit?": st.column_config.TextColumn(width="small"),
                     })

        st.divider()
        with st.expander("🗑️ Delete a Record"):
            st.warning("Deleting is permanent and cannot be undone.")
            if "id" in df_view.columns:
                del_options = {
                    f"Row {i+1} | {row['Date']} | {row.get('Serial Number','')} | {row.get('Failure Type','')}": row["id"]
                    for i, (_, row) in enumerate(df_view.iterrows())
                }
                selected = st.selectbox("Select record to delete", list(del_options.keys()))
                if st.button("Delete Record", type="primary"):
                    delete_record(del_options[selected])
                    st.success("Record deleted.")
                    st.rerun()
