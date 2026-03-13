"""
AMS (Abbott Application Management Support) Operational Dashboard
Run with: streamlit run ams_dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime
import io

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="AMS Operations Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  GLOBAL STYLES
# ─────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

  html, body, [class*="css"] {
      font-family: 'IBM Plex Sans', sans-serif;
  }

  /* ── Sidebar background & border ── */
  [data-testid="stSidebar"] {
      background: #0f1923;
      border-right: 2px solid #009fe3;
  }

  /* Sidebar headings and plain text */
  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3,
  [data-testid="stSidebar"] p,
  [data-testid="stSidebar"] .stMarkdown,
  [data-testid="stSidebar"] .stCaption { color: #e8edf2 !important; }

  /* Widget labels */
  [data-testid="stSidebar"] label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stDateInput label {
      color: #7db9d4 !important;
      font-size: 0.78rem; font-weight: 600;
      text-transform: uppercase; letter-spacing: 0.08em;
  }

  /* File uploader box */
  [data-testid="stSidebar"] [data-testid="stFileUploader"] {
      background: #1c2a38;
      border: 1px dashed #3a6a8a;
      border-radius: 8px;
      padding: 0.4rem;
  }
  [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {
      color: #a8c8e0 !important;
  }
  [data-testid="stSidebar"] [data-testid="stFileUploader"] button {
      background: #005587 !important;
      color: #ffffff !important;
      border: none !important;
      border-radius: 6px !important;
  }

  /* Uploaded file name chip */
  [data-testid="stSidebar"] [data-testid="stFileUploaderFile"] {
      background: #1c2a38;
      border-radius: 6px;
      padding: 0.3rem 0.6rem;
  }
  [data-testid="stSidebar"] [data-testid="stFileUploaderFileName"],
  [data-testid="stSidebar"] [data-testid="stFileUploaderFileData"] {
      color: #e8edf2 !important;
  }

  /* Multiselect pill + input box */
  [data-testid="stSidebar"] [data-baseweb="select"] > div {
      background: #1c2a38 !important;
      border-color: #3a6a8a !important;
  }
  [data-testid="stSidebar"] [data-baseweb="select"] * { color: #e8edf2 !important; }
  [data-testid="stSidebar"] [data-baseweb="tag"] {
      background: #005587 !important;
  }
  [data-testid="stSidebar"] [data-baseweb="tag"] span { color: #ffffff !important; }

  /* Date input */
  [data-testid="stSidebar"] [data-testid="stDateInput"] input {
      background: #1c2a38 !important;
      color: #e8edf2 !important;
      border-color: #3a6a8a !important;
  }

  /* Divider */
  [data-testid="stSidebar"] hr { border-color: #2a3f55; }

  /* ── Top header banner ── */
  .ams-header {
      background: linear-gradient(135deg, #005587 0%, #009fe3 60%, #00c4b3 100%);
      padding: 1.4rem 2rem 1.2rem 2rem;
      border-radius: 10px;
      margin-bottom: 1.4rem;
      display: flex; align-items: center; justify-content: space-between;
  }
  .ams-header h1 { margin: 0; color: #fff; font-size: 1.6rem; font-weight: 700; letter-spacing: -0.02em; }
  .ams-header p  { margin: 0; color: rgba(255,255,255,0.75); font-size: 0.82rem; font-family: 'IBM Plex Mono', monospace; }
  .ams-badge { background: rgba(255,255,255,0.18); color: #fff; padding: 0.25rem 0.75rem; border-radius: 20px; font-size: 0.75rem; font-weight: 600; letter-spacing: 0.06em; }

  /* ── KPI cards ── */
  .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin-bottom: 1.4rem; }
  .kpi-card {
      background: #ffffff; border: 1px solid #e2e8ef;
      border-radius: 10px; padding: 1.1rem 1.3rem;
      border-left: 4px solid #009fe3;
      box-shadow: 0 1px 6px rgba(0,0,0,0.06);
  }
  .kpi-card.red   { border-left-color: #e63946; }
  .kpi-card.green { border-left-color: #2dc653; }
  .kpi-card.amber { border-left-color: #f4a200; }
  .kpi-label { font-size: 0.72rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: #6b7a90; margin-bottom: 0.3rem; }
  .kpi-value { font-size: 2rem; font-weight: 700; color: #0f1923; line-height: 1; }
  .kpi-sub   { font-size: 0.76rem; color: #6b7a90; margin-top: 0.25rem; }

  /* ── Section headers ── */
  .section-title {
      font-size: 0.78rem; font-weight: 700; text-transform: uppercase;
      letter-spacing: 0.1em; color: #005587;
      border-bottom: 2px solid #009fe3; padding-bottom: 0.3rem;
      margin: 1rem 0 0.8rem 0;
  }

  /* ── Throughput badge ── */
  .throughput-badge { display: inline-block; padding: 0.2rem 0.7rem; border-radius: 4px; font-size: 0.82rem; font-weight: 700; font-family: 'IBM Plex Mono', monospace; }
  .tp-green  { background: #d4f7dc; color: #1a7a35; }
  .tp-yellow { background: #fff3cd; color: #7a5a00; }
  .tp-red    { background: #fde8e8; color: #9b1a1a; }

  /* ── Streamlit overrides ── */
  .stTabs [data-baseweb="tab-list"] { gap: 8px; border-bottom: 2px solid #e2e8ef; }
  .stTabs [data-baseweb="tab"] {
      background: transparent; border: none; border-radius: 6px 6px 0 0;
      color: #6b7a90; font-weight: 600; font-size: 0.82rem; letter-spacing: 0.04em;
      padding: 0.5rem 1.2rem;
  }
  .stTabs [aria-selected="true"] { background: #009fe3 !important; color: #fff !important; }
  div[data-testid="stMetric"] { background: transparent; }
  .stDownloadButton > button {
      background: #005587; color: #fff; border: none;
      border-radius: 6px; font-weight: 600; font-size: 0.8rem;
  }
  .stDownloadButton > button:hover { background: #009fe3; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
EXCEL_SERIAL_BASE = datetime.datetime(1899, 12, 30)

def serial_to_date(s):
    """Convert Excel serial integer to Python date."""
    try:
        return (EXCEL_SERIAL_BASE + datetime.timedelta(days=int(s))).date()
    except Exception:
        return None

def parse_date_col(series):
    """Try multiple formats to parse date column."""
    for fmt in ("%m-%d-%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return pd.to_datetime(series, format=fmt, errors="raise")
        except Exception:
            continue
    return pd.to_datetime(series, infer_datetime_format=True, errors="coerce")

def week_label(d):
    """Return 'DD-Mon' label for a date."""
    if pd.isnull(d):
        return "Unknown"
    return d.strftime("%-d-%b") if hasattr(d, "strftime") else str(d)

def extract_app(group_str):
    """
    Extract a short application / function name from the Assignment Group string.
    e.g. 'IBM-GLOBAL-Appspprt-ColTech-Non-Critical Sharepoint'
         -> 'Sharepoint'
    """
    if pd.isnull(group_str):
        return "Unknown"
    parts = str(group_str).split("-")
    # Take the last token after the final dash, then split on space
    tail = parts[-1].strip()
    # Remove leading 'Non-Critical ' prefix if present (after split tail may already be clean)
    return tail if tail else group_str

def age_bucket(days):
    if days <= 60:
        return "0–60 Days"
    elif days <= 90:
        return "61–90 Days"
    elif days <= 120:
        return "91–120 Days"
    else:
        return "120+ Days"

AGE_ORDER   = ["0–60 Days", "61–90 Days", "91–120 Days", "120+ Days"]
AGE_COLORS  = ["#009fe3", "#f4a200", "#e63946", "#6b0f1a"]
PLOTLY_TEAL = "#009fe3"
PLOTLY_NAVY = "#005587"

PRIORITY_MAP = {
    "Critical": "P1", "High": "P2", "Moderate": "P3",
    "Low": "P4", "P1": "P1", "P2": "P2", "P3": "P3", "P4": "P4",
}
PRIORITY_COLORS = {"P1": "#e63946", "P2": "#f4a200", "P3": "#009fe3", "P4": "#2dc653"}


# ─────────────────────────────────────────────
#  DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_and_preprocess(created_bytes, resolved_bytes):
    def _read(b):
        name = getattr(b, "name", "")
        if name.endswith(".csv"):
            return pd.read_csv(b)
        return pd.read_excel(b)

    cr = _read(created_bytes)
    re = _read(resolved_bytes)

    # Normalise column names
    cr.columns = [c.strip() for c in cr.columns]
    re.columns = [c.strip() for c in re.columns]

    # ── Created dataset ──
    cr["Created On"]  = parse_date_col(cr["Created On"])
    if "Resolved On" in cr.columns:
        cr["Resolved On"] = parse_date_col(cr["Resolved On"])

    # Week Ending: if numeric treat as Excel serial; else parse as date
    if "Created Week Ending" in cr.columns:
        if pd.api.types.is_numeric_dtype(cr["Created Week Ending"]):
            cr["Created Week Ending"] = cr["Created Week Ending"].apply(
                lambda x: EXCEL_SERIAL_BASE + datetime.timedelta(days=int(x))
                if pd.notna(x) else pd.NaT
            )
        else:
            cr["Created Week Ending"] = parse_date_col(cr["Created Week Ending"])

    # ── Resolved dataset ──
    re["Resolved On"] = parse_date_col(re["Resolved On"])
    if "Created On" in re.columns:
        re["Created On"] = parse_date_col(re["Created On"])

    if "Resolved Week Ending" in re.columns:
        if pd.api.types.is_numeric_dtype(re["Resolved Week Ending"]):
            re["Resolved Week Ending"] = re["Resolved Week Ending"].apply(
                lambda x: EXCEL_SERIAL_BASE + datetime.timedelta(days=int(x))
                if pd.notna(x) else pd.NaT
            )
        else:
            re["Resolved Week Ending"] = parse_date_col(re["Resolved Week Ending"])

    # Application shortname
    for df in (cr, re):
        if "Assignment Group" in df.columns:
            df["Application"] = df["Assignment Group"].apply(extract_app)
        else:
            df["Application"] = "Unknown"

    # Priority normalisation
    for df in (cr, re):
        if "Priority" in df.columns:
            df["Priority_Norm"] = df["Priority"].map(PRIORITY_MAP).fillna(df["Priority"])
        else:
            df["Priority_Norm"] = "Unknown"

    # Resolution time (days) on created file
    if "Resolved On" in cr.columns and "Created On" in cr.columns:
        cr["Resolution Days"] = (cr["Resolved On"] - cr["Created On"]).dt.days

    # Open incidents = in created but NOT in resolved
    resolved_numbers = set(re["Number"].dropna().unique())
    cr["Is Open"] = ~cr["Number"].isin(resolved_numbers)

    # Ageing for open
    today = pd.Timestamp.today().normalize()
    cr["Age Days"] = (today - cr["Created On"]).dt.days.clip(lower=0)
    cr["Age Bucket"] = cr["Age Days"].apply(age_bucket)
    cr["Age Bucket"] = pd.Categorical(cr["Age Bucket"], categories=AGE_ORDER, ordered=True)

    return cr, re


# ─────────────────────────────────────────────
#  CHART HELPERS
# ─────────────────────────────────────────────
def chart_layout(fig, title="", height=360):
    fig.update_layout(
        title=dict(text=title, font=dict(size=13, color="#0f1923", family="IBM Plex Sans"), x=0, pad=dict(l=0)),
        height=height,
        margin=dict(l=10, r=10, t=40 if title else 20, b=10),
        paper_bgcolor="white",
        plot_bgcolor="#f7fafd",
        font=dict(family="IBM Plex Sans", size=11, color="#3a4a5c"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(size=10)),
        xaxis=dict(gridcolor="#e8edf2", linecolor="#c8d4df"),
        yaxis=dict(gridcolor="#e8edf2", linecolor="#c8d4df"),
    )
    return fig


def weekly_throughput_fig(cr_w, re_w):
    """Bar+line chart matching the reference image."""
    # Aggregate
    cr_agg = (cr_w.groupby("Created Week Ending")
                   .size().reset_index(name="Created"))
    re_agg = (re_w.groupby("Resolved Week Ending")
                   .size().reset_index(name="Resolved"))

    cr_agg = cr_agg.rename(columns={"Created Week Ending": "Week"})
    re_agg = re_agg.rename(columns={"Resolved Week Ending": "Week"})

    weeks = sorted(set(cr_agg["Week"]).union(set(re_agg["Week"])))
    df = pd.DataFrame({"Week": weeks})
    df = df.merge(cr_agg, on="Week", how="left").merge(re_agg, on="Week", how="left").fillna(0)
    df["Throughput"] = (df["Resolved"] / df["Created"].replace(0, np.nan) * 100).round(1)
    df["Week Label"] = df["Week"].dt.strftime("%-d-%b")
    df = df.sort_values("Week")

    fig = go.Figure()
    fig.add_bar(x=df["Week Label"], y=df["Created"],  name="Created",  marker_color=PLOTLY_NAVY, opacity=0.85)
    fig.add_bar(x=df["Week Label"], y=df["Resolved"], name="Resolved", marker_color=PLOTLY_TEAL, opacity=0.85)
    fig.add_scatter(x=df["Week Label"], y=df["Throughput"],
                    name="Throughput %", yaxis="y2",
                    mode="lines+markers+text",
                    line=dict(color="#e63946", width=2.5),
                    marker=dict(size=7, color="#e63946"),
                    text=[f"{v:.0f}%" for v in df["Throughput"]],
                    textposition="top center",
                    textfont=dict(size=9, color="#e63946"))

    fig.update_layout(
        barmode="group",
        yaxis=dict(title="Count", gridcolor="#e8edf2"),
        yaxis2=dict(title="Throughput %", overlaying="y", side="right",
                    showgrid=False, range=[0, max(df["Throughput"].max() * 1.4, 120)]),
    )
    return chart_layout(fig, "Weekly Incident Throughput", height=400), df


def ageing_fig(open_df):
    """Stacked bar chart of open incidents by app and age bucket."""
    grp = (open_df.groupby(["Application", "Age Bucket"], observed=True)
                  .size().reset_index(name="Count"))
    pivot = (grp.pivot(index="Application", columns="Age Bucket", values="Count")
                .fillna(0).reset_index())
    # Sort by total desc
    pivot["Total"] = pivot[[c for c in AGE_ORDER if c in pivot.columns]].sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=True).tail(20)

    fig = go.Figure()
    for bucket, color in zip(AGE_ORDER, AGE_COLORS):
        if bucket in pivot.columns:
            fig.add_bar(
                y=pivot["Application"], x=pivot[bucket],
                name=bucket, orientation="h",
                marker_color=color,
            )
    fig.update_layout(barmode="stack",
                      xaxis_title="Open Incidents",
                      yaxis=dict(tickfont=dict(size=9)),
                      height=max(350, len(pivot) * 28 + 80))
    return chart_layout(fig, "Open Incident Ageing by Application")


def trend_fig(cr_w):
    daily = cr_w.groupby(cr_w["Created On"].dt.date).size().reset_index(name="Count")
    daily.columns = ["Date", "Count"]
    daily = daily.sort_values("Date")
    fig = px.area(daily, x="Date", y="Count",
                  color_discrete_sequence=[PLOTLY_TEAL])
    fig.update_traces(fill="tozeroy", fillcolor="rgba(0,159,227,0.12)", line_width=2)
    return chart_layout(fig, "Daily Incident Creation Trend", height=300)


def top_apps_fig(cr_w):
    top = (cr_w.groupby("Application").size()
                .reset_index(name="Count")
                .sort_values("Count", ascending=True).tail(15))
    fig = px.bar(top, x="Count", y="Application", orientation="h",
                 color="Count", color_continuous_scale=["#b3d9f0", PLOTLY_NAVY])
    fig.update_coloraxes(showscale=False)
    return chart_layout(fig, "Top Applications by Incident Volume", height=380)


def priority_fig(cr_w):
    p = cr_w["Priority_Norm"].value_counts().reset_index()
    p.columns = ["Priority", "Count"]
    colors = [PRIORITY_COLORS.get(x, "#aaa") for x in p["Priority"]]
    fig = px.pie(p, names="Priority", values="Count",
                 color="Priority",
                 color_discrete_map=PRIORITY_COLORS,
                 hole=0.52)
    fig.update_traces(textinfo="label+percent", textfont_size=11)
    return chart_layout(fig, "Priority Distribution", height=320)


def resolution_time_fig(cr_w):
    rt = (cr_w.dropna(subset=["Resolution Days"])
               .groupby("Application")["Resolution Days"]
               .mean().reset_index()
               .rename(columns={"Resolution Days": "Avg Days"})
               .sort_values("Avg Days", ascending=True).tail(15))
    fig = px.bar(rt, x="Avg Days", y="Application", orientation="h",
                 color="Avg Days", color_continuous_scale=["#d4f7dc", "#e63946"])
    fig.update_coloraxes(showscale=False)
    return chart_layout(fig, "Avg Resolution Time by Application (Days)", height=380)


def open_closed_fig(cr_w):
    counts = cr_w["Is Open"].value_counts()
    labels = ["Open", "Resolved"]
    values = [counts.get(True, 0), counts.get(False, 0)]
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        marker_colors=["#e63946", "#2dc653"],
        hole=0.5, textinfo="label+value+percent", textfont_size=11,
    ))
    return chart_layout(fig, "Open vs Resolved Incidents", height=300)


def sla_breach_fig(cr_w):
    """Simple proxy: incidents older than 5 days still open = breached."""
    SLA_THRESHOLDS = {"P1": 1, "P2": 3, "P3": 5, "P4": 7}
    open_df = cr_w[cr_w["Is Open"]].copy()
    open_df["SLA Days"] = open_df["Priority_Norm"].map(SLA_THRESHOLDS).fillna(5)
    open_df["Breached"] = open_df["Age Days"] > open_df["SLA Days"]
    grp = (open_df.groupby(["Application", "Breached"]).size()
                  .reset_index(name="Count"))
    breached = grp[grp["Breached"]].sort_values("Count", ascending=True).tail(12)
    if breached.empty:
        return None
    fig = px.bar(breached, x="Count", y="Application", orientation="h",
                 color_discrete_sequence=["#e63946"])
    return chart_layout(fig, "Top Applications – SLA Breached Open Incidents", height=340)


def category_fig(cr_w):
    """Use Assignment Group as category proxy."""
    top = (cr_w.groupby("Application").size()
                .reset_index(name="Count")
                .sort_values("Count", ascending=False).head(10))
    fig = px.bar(top, x="Application", y="Count",
                 color="Count", color_continuous_scale=["#b3d9f0", PLOTLY_NAVY])
    fig.update_coloraxes(showscale=False)
    fig.update_xaxes(tickangle=-30)
    return chart_layout(fig, "Top Incident Categories (Assignment Group)", height=340)


# ─────────────────────────────────────────────
#  EXCEL EXPORT
# ─────────────────────────────────────────────
def build_excel(cr, re, throughput_df, open_df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        cr.to_excel(writer, sheet_name="Created Incidents", index=False)
        re.to_excel(writer, sheet_name="Resolved Incidents", index=False)
        throughput_df.to_excel(writer, sheet_name="Weekly Throughput", index=False)
        open_df[["Number", "Application", "Priority_Norm", "Created On",
                 "Age Days", "Age Bucket"]].to_excel(writer, sheet_name="Open Ageing", index=False)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────
#  SIDEBAR – file uploaders only (filters added after data loads)
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Upload Data")
    st.markdown("---")
    created_file  = st.file_uploader("Incident **Created** File (CSV / Excel)", type=["csv","xlsx","xls"], key="cr")
    resolved_file = st.file_uploader("Incident **Resolved** File (CSV / Excel)", type=["csv","xlsx","xls"], key="re")
    st.markdown("---")
    st.caption("AMS Operations Dashboard v1.0")


# ─────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="ams-header">
  <div>
    <h1>📊 AMS Operations Dashboard</h1>
    <p>Abbott Application Management Support · ServiceNow Incident Analytics</p>
  </div>
  <span class="ams-badge">LIVE REPORT</span>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  GUARD: wait for files
# ─────────────────────────────────────────────
if not created_file or not resolved_file:
    st.info("👈  Upload both **Incident Created** and **Incident Resolved** files in the sidebar to get started.")
    st.stop()


# ─────────────────────────────────────────────
#  LOAD DATA
# ─────────────────────────────────────────────
with st.spinner("Processing data…"):
    cr_raw, re_raw = load_and_preprocess(created_file, resolved_file)

# ── Build filter option lists from loaded data ──
apps   = sorted(cr_raw["Application"].dropna().unique())
pris   = sorted(cr_raw["Priority_Norm"].dropna().unique())
groups = sorted(cr_raw["Assignment Group"].dropna().unique()) if "Assignment Group" in cr_raw.columns else []

min_date = cr_raw["Created On"].min().date()
max_date = cr_raw["Created On"].max().date()

# ── Render filters once (same keys, populated options) ──
with st.sidebar:
    st.markdown("### ⚙️ Filters")
    sel_apps   = st.multiselect("Application",      apps,   default=[], key="f_app")
    sel_pris   = st.multiselect("Priority",         pris,   default=[], key="f_pri")
    sel_groups = st.multiselect("Assignment Group", groups, default=[], key="f_grp")
    date_range = st.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        key="f_date",
    )

# ── Apply filters ──
cr = cr_raw.copy()
re = re_raw.copy()

if sel_apps:
    cr = cr[cr["Application"].isin(sel_apps)]
    re = re[re["Application"].isin(sel_apps)]
if sel_pris:
    cr = cr[cr["Priority_Norm"].isin(sel_pris)]
    re = re[re["Priority_Norm"].isin(sel_pris)]
if sel_groups and "Assignment Group" in cr.columns:
    cr = cr[cr["Assignment Group"].isin(sel_groups)]
    re = re[re["Assignment Group"].isin(sel_groups)]

if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    s, e = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
    cr = cr[(cr["Created On"] >= s) & (cr["Created On"] <= e)]
    re = re[(re["Resolved On"] >= s) & (re["Resolved On"] <= e)]


# ─────────────────────────────────────────────
#  KPI CARDS
# ─────────────────────────────────────────────
total_created  = len(cr)
total_resolved = len(re)
open_count     = cr["Is Open"].sum()
throughput_pct = round(total_resolved / total_created * 100, 1) if total_created else 0
avg_res        = cr["Resolution Days"].mean() if "Resolution Days" in cr.columns else 0

tp_class = "green" if throughput_pct >= 100 else ("amber" if throughput_pct >= 85 else "red")

st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">Total Created</div>
    <div class="kpi-value">{total_created:,}</div>
    <div class="kpi-sub">All logged incidents</div>
  </div>
  <div class="kpi-card green">
    <div class="kpi-label">Total Resolved</div>
    <div class="kpi-value">{total_resolved:,}</div>
    <div class="kpi-sub">Closed incidents</div>
  </div>
  <div class="kpi-card red">
    <div class="kpi-label">Open / Backlog</div>
    <div class="kpi-value">{int(open_count):,}</div>
    <div class="kpi-sub">Pending resolution</div>
  </div>
  <div class="kpi-card {tp_class}">
    <div class="kpi-label">Overall Throughput</div>
    <div class="kpi-value">{throughput_pct}%</div>
    <div class="kpi-sub">Resolved ÷ Created</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🏠 Overview",
    "📅 Weekly Throughput",
    "⏳ Ageing Analysis",
    "📈 Incident Analytics",
    "⬇️ Download",
])


# ═══════════════════════════════════════════════
#  TAB 1 – OVERVIEW
# ═══════════════════════════════════════════════
with tab1:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-title">Open vs Resolved</div>', unsafe_allow_html=True)
        st.plotly_chart(open_closed_fig(cr), use_container_width=True, config={"displayModeBar": False}, key="ov_open_closed")

    with c2:
        st.markdown('<div class="section-title">Priority Breakdown</div>', unsafe_allow_html=True)
        st.plotly_chart(priority_fig(cr), use_container_width=True, config={"displayModeBar": False}, key="ov_priority")

    st.markdown('<div class="section-title">Daily Incident Volume Trend</div>', unsafe_allow_html=True)
    st.plotly_chart(trend_fig(cr), use_container_width=True, config={"displayModeBar": False}, key="ov_trend")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="section-title">Top Applications</div>', unsafe_allow_html=True)
        st.plotly_chart(top_apps_fig(cr), use_container_width=True, config={"displayModeBar": False}, key="ov_top_apps")
    with c4:
        st.markdown('<div class="section-title">Resolution Time</div>', unsafe_allow_html=True)
        st.plotly_chart(resolution_time_fig(cr), use_container_width=True, config={"displayModeBar": False}, key="ov_res_time")


# ═══════════════════════════════════════════════
#  TAB 2 – WEEKLY THROUGHPUT
# ═══════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">Incident Weekly Throughput – Created vs Resolved</div>',
                unsafe_allow_html=True)

    fig_tp, tp_df = weekly_throughput_fig(cr, re)
    st.plotly_chart(fig_tp, use_container_width=True, config={"displayModeBar": False}, key="wk_throughput")

    st.markdown('<div class="section-title">Throughput Table</div>', unsafe_allow_html=True)

    def fmt_tp(v):
        cls = "tp-green" if v >= 100 else ("tp-yellow" if v >= 85 else "tp-red")
        return f'<span class="throughput-badge {cls}">{v:.0f}%</span>'

    display_tp = tp_df[["Week Label", "Created", "Resolved", "Throughput"]].copy()
    display_tp["Throughput"] = display_tp["Throughput"].apply(fmt_tp)
    st.markdown(
        display_tp.to_html(index=False, escape=False,
                           classes="dataframe", border=0),
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════
#  TAB 3 – AGEING ANALYSIS
# ═══════════════════════════════════════════════
with tab3:
    open_df = cr[cr["Is Open"]].copy()

    kA, kB, kC, kD = st.columns(4)
    for col, bucket in zip([kA, kB, kC, kD], AGE_ORDER):
        cnt = (open_df["Age Bucket"] == bucket).sum()
        col.metric(bucket, f"{cnt:,}")

    st.markdown('<div class="section-title">Open Incident Ageing by Application</div>',
                unsafe_allow_html=True)
    if open_df.empty:
        st.success("✅ No open incidents found.")
    else:
        st.plotly_chart(ageing_fig(open_df), use_container_width=True, config={"displayModeBar": False}, key="ag_ageing")

    st.markdown('<div class="section-title">Ageing Summary Table</div>', unsafe_allow_html=True)
    age_pivot = (open_df.groupby(["Application", "Age Bucket"], observed=True)
                        .size().reset_index(name="Count")
                        .pivot(index="Application", columns="Age Bucket", values="Count")
                        .fillna(0).astype(int))
    # Ensure all columns present
    for b in AGE_ORDER:
        if b not in age_pivot.columns:
            age_pivot[b] = 0
    age_pivot = age_pivot[AGE_ORDER]
    age_pivot["Grand Total"] = age_pivot.sum(axis=1)
    age_pivot = age_pivot.sort_values("Grand Total", ascending=False)
    st.dataframe(age_pivot, use_container_width=True)

    st.markdown('<div class="section-title">SLA Breach Analysis</div>', unsafe_allow_html=True)
    sla_f = sla_breach_fig(cr)
    if sla_f:
        st.plotly_chart(sla_f, use_container_width=True, config={"displayModeBar": False}, key="ag_sla")
    else:
        st.info("No SLA breaches detected in the current filter selection.")


# ═══════════════════════════════════════════════
#  TAB 4 – INCIDENT ANALYTICS
# ═══════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">Incident Category Distribution</div>', unsafe_allow_html=True)
    st.plotly_chart(category_fig(cr), use_container_width=True, config={"displayModeBar": False}, key="an_category")

    c5, c6 = st.columns(2)
    with c5:
        st.markdown('<div class="section-title">Priority Distribution</div>', unsafe_allow_html=True)
        st.plotly_chart(priority_fig(cr), use_container_width=True, config={"displayModeBar": False}, key="an_priority")
    with c6:
        st.markdown('<div class="section-title">Avg Resolution by Priority</div>', unsafe_allow_html=True)
        if "Resolution Days" in cr.columns:
            rp = (cr.dropna(subset=["Resolution Days"])
                    .groupby("Priority_Norm")["Resolution Days"]
                    .mean().reset_index()
                    .rename(columns={"Resolution Days": "Avg Days"}))
            rp["Color"] = rp["Priority_Norm"].map(PRIORITY_COLORS).fillna("#aaa")
            fig_rp = px.bar(rp, x="Priority_Norm", y="Avg Days",
                            color="Priority_Norm", color_discrete_map=PRIORITY_COLORS,
                            text=rp["Avg Days"].round(1))
            fig_rp.update_traces(texttemplate="%{text} d", textposition="outside")
            st.plotly_chart(chart_layout(fig_rp, height=300),
                            use_container_width=True, config={"displayModeBar": False}, key="an_res_pri")

    # Weekly heatmap
    st.markdown('<div class="section-title">Creation Heatmap (Weekday × Week)</div>', unsafe_allow_html=True)
    hm = cr.copy()
    hm["Weekday"] = hm["Created On"].dt.day_name()
    hm["WeekOf"]  = hm["Created On"].dt.to_period("W").dt.start_time.dt.strftime("%-d %b")
    pivot_hm = (hm.groupby(["Weekday", "WeekOf"]).size().reset_index(name="Count")
                  .pivot(index="Weekday", columns="WeekOf", values="Count").fillna(0))
    day_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    pivot_hm = pivot_hm.reindex([d for d in day_order if d in pivot_hm.index])
    fig_hm = px.imshow(pivot_hm, color_continuous_scale=["#f0f8ff", PLOTLY_NAVY],
                       aspect="auto")
    fig_hm.update_coloraxes(showscale=True)
    st.plotly_chart(chart_layout(fig_hm, height=300),
                    use_container_width=True, config={"displayModeBar": False}, key="an_heatmap")

    # Raw data viewer
    with st.expander("🔍 View Raw Created Incidents"):
        st.dataframe(cr[["Number","Application","Priority_Norm","Created On",
                          "Is Open","Age Days","Age Bucket"]].sort_values("Created On", ascending=False),
                     use_container_width=True, height=320)


# ═══════════════════════════════════════════════
#  TAB 5 – DOWNLOAD
# ═══════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-title">Export Reports</div>', unsafe_allow_html=True)
    st.write("Download the processed data and summary reports as an Excel workbook.")

    _, tp_df2 = weekly_throughput_fig(cr, re)
    open_export = cr[cr["Is Open"]].copy()
    excel_bytes = build_excel(cr, re, tp_df2, open_export)

    st.download_button(
        label="⬇️ Download Full Report (Excel)",
        data=excel_bytes,
        file_name=f"AMS_Dashboard_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")
    st.markdown("**Workbook contains:**")
    st.markdown("""
    - **Created Incidents** – Full created dataset with calculated fields  
    - **Resolved Incidents** – Resolved dataset  
    - **Weekly Throughput** – Week-by-week created / resolved / throughput %  
    - **Open Ageing** – Open backlog with age bucket classification  
    """)
