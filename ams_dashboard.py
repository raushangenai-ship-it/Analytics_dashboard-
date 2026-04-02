"""
AMS (Abbott Application Management Support) Operational Dashboard
Run with: streamlit run ams_dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import datetime
import io

# ─────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AMS Operations Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────
#  THEME & GLOBAL CSS
# ─────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

/* ── SIDEBAR ── */
[data-testid="stSidebar"] { background:#0d1b2a; border-right:2px solid #0078b6; }
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,[data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stMarkdown,[data-testid="stSidebar"] .stCaption
    { color:#d6e8f5 !important; }
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stDateInput label,
[data-testid="stSidebar"] .stSelectbox label
    { color:#7eb8d8 !important; font-size:0.72rem !important;
      font-weight:600 !important; text-transform:uppercase !important;
      letter-spacing:0.08em !important; }
[data-testid="stSidebar"] [data-testid="stFileUploader"]
    { background:#132233; border:1.5px dashed #2a6080; border-radius:8px; padding:0.5rem; }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] *
    { color:#8bbdd6 !important; }
[data-testid="stSidebar"] [data-testid="stFileUploader"] button
    { background:#0078b6 !important; color:#fff !important;
      border:none !important; border-radius:5px !important; }
[data-testid="stSidebar"] [data-testid="stFileUploaderFileName"],
[data-testid="stSidebar"] [data-testid="stFileUploaderFileData"]
    { color:#d6e8f5 !important; }
[data-testid="stSidebar"] [data-baseweb="select"] > div
    { background:#132233 !important; border-color:#2a6080 !important; }
[data-testid="stSidebar"] [data-baseweb="select"] *
    { color:#d6e8f5 !important; }
[data-testid="stSidebar"] [data-baseweb="tag"]
    { background:#0078b6 !important; }
[data-testid="stSidebar"] [data-baseweb="tag"] span { color:#fff !important; }
[data-testid="stSidebar"] [data-testid="stDateInput"] input
    { background:#132233 !important; color:#d6e8f5 !important;
      border-color:#2a6080 !important; }
[data-testid="stSidebar"] hr { border-color:#1e3348; }
[data-testid="stSidebar"] .stCheckbox label span { color:#d6e8f5 !important; }
[data-testid="stSidebar"] .stRadio label span { color:#d6e8f5 !important; }

/* ── MAIN HEADER ── */
.ams-header {
    background:linear-gradient(135deg,#003f6b 0%,#0078b6 55%,#00afd8 100%);
    padding:1.2rem 2rem; border-radius:10px; margin-bottom:1.2rem;
    display:flex; align-items:center; justify-content:space-between;
    box-shadow:0 4px 20px rgba(0,120,182,0.25);
}
.ams-header h1 { margin:0; color:#fff; font-size:1.55rem; font-weight:700; letter-spacing:-0.02em; }
.ams-header p  { margin:0.2rem 0 0; color:rgba(255,255,255,0.72); font-size:0.8rem;
    font-family:'IBM Plex Mono',monospace; }
.ams-right { text-align:right; }
.ams-badge { background:rgba(255,255,255,0.15); color:#fff; padding:0.22rem 0.9rem;
    border-radius:20px; font-size:0.73rem; font-weight:700; letter-spacing:0.07em;
    display:inline-block; margin-bottom:0.3rem; }
.ams-date  { color:rgba(255,255,255,0.65); font-size:0.75rem;
    font-family:'IBM Plex Mono',monospace; }

/* ── KPI STRIP ── */
.kpi-row { display:grid; grid-template-columns:repeat(6,1fr);
    gap:0.75rem; margin-bottom:1.2rem; }
.kpi { background:#fff; border:1px solid #e0eaf3; border-radius:10px;
    padding:0.9rem 1rem; border-top:3px solid #0078b6;
    box-shadow:0 1px 5px rgba(0,0,0,0.05); }
.kpi.red    { border-top-color:#d62828; }
.kpi.green  { border-top-color:#1a9e4e; }
.kpi.amber  { border-top-color:#d97706; }
.kpi.purple { border-top-color:#7c3aed; }
.kpi.teal   { border-top-color:#0891b2; }
.kpi-label  { font-size:0.68rem; font-weight:700; text-transform:uppercase;
    letter-spacing:0.09em; color:#5a7a94; margin-bottom:0.25rem; }
.kpi-value  { font-size:1.85rem; font-weight:700; color:#0d1b2a; line-height:1; }
.kpi-delta  { font-size:0.72rem; margin-top:0.25rem; }
.delta-up   { color:#d62828; }
.delta-dn   { color:#1a9e4e; }
.delta-neu  { color:#5a7a94; }

/* ── SECTION TITLES ── */
.sec { font-size:0.72rem; font-weight:700; text-transform:uppercase;
    letter-spacing:0.1em; color:#003f6b;
    border-bottom:2px solid #0078b6; padding-bottom:0.3rem;
    margin:0.9rem 0 0.7rem 0; }

/* ── ALERT BANNERS ── */
.alert-critical { background:#fff0f0; border-left:4px solid #d62828;
    padding:0.7rem 1rem; border-radius:6px; margin-bottom:0.6rem; }
.alert-warn     { background:#fffbeb; border-left:4px solid #d97706;
    padding:0.7rem 1rem; border-radius:6px; margin-bottom:0.6rem; }
.alert-ok       { background:#f0faf4; border-left:4px solid #1a9e4e;
    padding:0.7rem 1rem; border-radius:6px; margin-bottom:0.6rem; }
.alert-title    { font-weight:700; font-size:0.85rem; }
.alert-body     { font-size:0.8rem; color:#444; margin-top:0.2rem; }

/* ── INCIDENT CARDS ── */
.inc-card { background:#fff; border:1px solid #dde8f0; border-radius:8px;
    padding:0.75rem 1rem; margin-bottom:0.5rem;
    box-shadow:0 1px 4px rgba(0,0,0,0.04); }
.inc-card.high   { border-left:4px solid #d62828; }
.inc-card.mod    { border-left:4px solid #d97706; }
.inc-card.low    { border-left:4px solid #0078b6; }
.inc-num  { font-family:'IBM Plex Mono',monospace; font-size:0.8rem;
    font-weight:600; color:#0078b6; }
.inc-app  { font-size:0.82rem; color:#0d1b2a; font-weight:500; }
.inc-meta { font-size:0.73rem; color:#6b8599; margin-top:0.2rem; }
.inc-age  { float:right; font-weight:700; font-size:0.82rem; }
.age-red  { color:#d62828; }
.age-amb  { color:#d97706; }
.age-grn  { color:#1a9e4e; }

/* ── THROUGHPUT TABLE ── */
.tp-table { width:100%; border-collapse:collapse; font-size:0.82rem; }
.tp-table th { background:#003f6b; color:#fff; padding:0.5rem 0.8rem;
    text-align:center; font-weight:600; font-size:0.73rem;
    text-transform:uppercase; letter-spacing:0.07em; }
.tp-table td { padding:0.45rem 0.8rem; text-align:center;
    border-bottom:1px solid #e8eef4; }
.tp-table tr:hover td { background:#f0f7fc; }
.tp-table tr:last-child td { font-weight:700; background:#e8f4fb; }
.badge { display:inline-block; padding:0.18rem 0.65rem; border-radius:4px;
    font-size:0.78rem; font-weight:700; font-family:'IBM Plex Mono',monospace; }
.bg-g { background:#d4f7dc; color:#145a2e; }
.bg-y { background:#fff3cd; color:#7a5200; }
.bg-r { background:#fde8e8; color:#8b1a1a; }

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"] { gap:6px; border-bottom:2px solid #dde8f0; }
.stTabs [data-baseweb="tab"] {
    background:transparent; border:none; border-radius:6px 6px 0 0;
    color:#5a7a94; font-weight:600; font-size:0.8rem;
    letter-spacing:0.03em; padding:0.45rem 1.1rem; }
.stTabs [aria-selected="true"] { background:#0078b6 !important; color:#fff !important; }

/* ── DOWNLOAD BUTTON ── */
.stDownloadButton > button { background:#003f6b; color:#fff; border:none;
    border-radius:6px; font-weight:600; font-size:0.8rem; padding:0.5rem 1.2rem; }
.stDownloadButton > button:hover { background:#0078b6; }

[data-testid="stDataFrame"] { border-radius:8px; overflow:hidden; }
.streamlit-expanderHeader { font-weight:600; color:#003f6b; font-size:0.85rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────────────────
EXCEL_BASE   = datetime.datetime(1899, 12, 30)
TODAY        = pd.Timestamp.today().normalize()
AGE_ORDER    = ["0-30 Days", "31-60 Days", "61-90 Days", "91-120 Days", "120+ Days"]
AGE_COLORS   = ["#1a9e4e", "#0078b6", "#d97706", "#d62828", "#6b0f1a"]
PRIORITY_ORD = ["High", "Moderate", "Low"]
PRI_COLORS   = {"High": "#d62828", "Moderate": "#d97706", "Low": "#0078b6"}
PRI_SLA      = {"High": 1, "Moderate": 3, "Low": 7}
C_NAVY       = "#003f6b"
C_BLUE       = "#0078b6"
C_TEAL       = "#00afd8"


# ─────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────
def parse_dates(series):
    for fmt in ("%m-%d-%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return pd.to_datetime(series, format=fmt, errors="raise")
        except Exception:
            pass
    return pd.to_datetime(series, infer_datetime_format=True, errors="coerce")

def serial_to_ts(val):
    try:
        return EXCEL_BASE + datetime.timedelta(days=int(val))
    except Exception:
        return pd.NaT

def extract_app(s):
    if pd.isnull(s): return "Unknown"
    parts = [p.strip() for p in str(s).split("-") if p.strip()]
    skip = {"IBM","GLOBAL","Global","AppSpprt","Appspprt","Appsupport",
            "ColTech","Comm","MEDDEV","Non","Critical","AppSupport"}
    filtered = [p for p in parts if p not in skip]
    return filtered[-1] if filtered else str(s).split("-")[-1].strip()

def age_bucket(d):
    if d <= 30:   return "0-30 Days"
    elif d <= 60:  return "31-60 Days"
    elif d <= 90:  return "61-90 Days"
    elif d <= 120: return "91-120 Days"
    return "120+ Days"

def is_auto(opener):
    kw = ["panda","integration","ai ","bot","auto","system","monitor"]
    return any(k in str(opener).lower() for k in kw)

def chart_style(fig, title="", height=340, legend=True):
    fig.update_layout(
        title=dict(text=title, font=dict(size=12.5, color="#0d1b2a",
                   family="IBM Plex Sans"), x=0, pad=dict(l=0)),
        height=height,
        margin=dict(l=8, r=8, t=42 if title else 18, b=8),
        paper_bgcolor="white", plot_bgcolor="#f5f9fd",
        font=dict(family="IBM Plex Sans", size=11, color="#3a4a5c"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)) if legend else dict(visible=False),
        xaxis=dict(gridcolor="#dde8f0", linecolor="#c0d0de", zeroline=False),
        yaxis=dict(gridcolor="#dde8f0", linecolor="#c0d0de", zeroline=False),
    )
    return fig


# ─────────────────────────────────────────────────────────
#  DATA LOADING
# ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load(created_bytes, resolved_bytes):
    def _read(f):
        name = getattr(f, "name", "")
        return pd.read_csv(f) if name.endswith(".csv") else pd.read_excel(f)

    cr = _read(created_bytes)
    re = _read(resolved_bytes)
    cr.columns = cr.columns.str.strip()
    re.columns = re.columns.str.strip()

    cr["Created On"]  = parse_dates(cr["Created On"])
    if "Resolved On" in cr.columns:
        cr["Resolved On"] = parse_dates(cr["Resolved On"])
    re["Created On"]  = parse_dates(re["Created On"])
    re["Resolved On"] = parse_dates(re["Resolved On"])

    for df, col in [(cr, "Created Week Ending"), (re, "Resolved Week Ending")]:
        if col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].apply(serial_to_ts)
            else:
                df[col] = parse_dates(df[col])

    for df in (cr, re):
        df["App"] = df["Assignment Group"].apply(extract_app)

    resolved_set = set(re["Number"].dropna())
    cr["Is Open"] = ~cr["Number"].isin(resolved_set)
    cr["Age Days"] = (TODAY - cr["Created On"]).dt.days.clip(lower=0)
    cr["Age Bucket"] = cr["Age Days"].apply(age_bucket)
    cr["Age Bucket"] = pd.Categorical(cr["Age Bucket"], categories=AGE_ORDER, ordered=True)

    if "Resolved On" in cr.columns:
        cr["Res Days"] = (cr["Resolved On"] - cr["Created On"]).dt.days
    else:
        cr["Res Days"] = np.nan

    if "Opened By" in cr.columns:
        cr["Source"] = cr["Opened By"].apply(
            lambda x: "Automated" if is_auto(x) else "Manual")
    else:
        cr["Source"] = "Manual"

    cr["SLA Target"]  = cr["Priority"].map(PRI_SLA).fillna(7)
    cr["SLA Breached"] = cr["Is Open"] & (cr["Age Days"] > cr["SLA Target"])
    pri_order = {"High": 0, "Moderate": 1, "Low": 2}
    cr["Pri Sort"] = cr["Priority"].map(pri_order).fillna(3)

    return cr, re


# ─────────────────────────────────────────────────────────
#  TOWER MAPPING LOADER
# ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_mapping(mapping_bytes):
    """Load SNOW Queue → Tower mapping file."""
    name = getattr(mapping_bytes, "name", "")
    df = pd.read_csv(mapping_bytes) if name.endswith(".csv") else pd.read_excel(mapping_bytes)
    df.columns = df.columns.str.strip()
    # Normalise column names flexibly
    col_map = {}
    for c in df.columns:
        cl = c.lower().strip()
        if "queue" in cl or "snow" in cl or "assignment" in cl:
            col_map[c] = "SNOW Queue Name"
        elif "tower" in cl:
            col_map[c] = "Tower"
    df = df.rename(columns=col_map)
    if "SNOW Queue Name" not in df.columns or "Tower" not in df.columns:
        st.warning("Mapping file must have columns for Queue Name and Tower.")
        return {}
    return dict(zip(df["SNOW Queue Name"].str.strip(), df["Tower"].str.strip()))

def apply_tower_mapping(df, mapping_dict):
    """Add Tower column to dataframe using Assignment Group lookup."""
    if not mapping_dict:
        df["Tower"] = "Unmapped"
        return df
    df["Tower"] = df["Assignment Group"].map(mapping_dict).fillna("Unmapped")
    return df

TOWER_COLORS = {
    "Collaboration":   "#0078b6",
    "SAP":             "#d97706",
    "Finance":         "#1a9e4e",
    "Manufacturing":   "#7c3aed",
    "Quality":         "#d62828",
    "Integration":     "#0891b2",
    "DBA":             "#374151",
    "Legacy ERP":      "#92400e",
    "BPCS":            "#be185d",
    "Unmapped":        "#9ca3af",
}


# ─────────────────────────────────────────────────────────
#  CHART FUNCTIONS
# ─────────────────────────────────────────────────────────
def fig_weekly_throughput(cr, re):
    cr_w = cr.groupby("Created Week Ending").size().reset_index(name="Created")
    re_w = re.groupby("Resolved Week Ending").size().reset_index(name="Resolved")
    cr_w = cr_w.rename(columns={"Created Week Ending": "Week"})
    re_w = re_w.rename(columns={"Resolved Week Ending": "Week"})
    weeks = sorted(set(cr_w["Week"]).union(re_w["Week"]))
    df = (pd.DataFrame({"Week": weeks})
            .merge(cr_w, on="Week", how="left")
            .merge(re_w, on="Week", how="left")
            .fillna(0))
    df["Throughput"] = (df["Resolved"] / df["Created"].replace(0, np.nan) * 100).round(1)
    df["WLabel"] = df["Week"].dt.strftime("%-d %b")
    df = df.sort_values("Week").reset_index(drop=True)

    fig = go.Figure()
    fig.add_bar(x=df["WLabel"], y=df["Created"],
                name="Created", marker_color=C_NAVY, opacity=0.88)
    fig.add_bar(x=df["WLabel"], y=df["Resolved"],
                name="Resolved", marker_color=C_TEAL, opacity=0.88)
    fig.add_scatter(x=df["WLabel"], y=df["Throughput"],
                    name="Throughput %", yaxis="y2",
                    mode="lines+markers+text",
                    line=dict(color="#d62828", width=2.5),
                    marker=dict(size=7, color="#d62828"),
                    text=[f"{v:.0f}%" for v in df["Throughput"]],
                    textposition="top center",
                    textfont=dict(size=9, color="#d62828"))
    fig.update_layout(
        barmode="group",
        yaxis=dict(title="Incidents"),
        yaxis2=dict(title="Throughput %", overlaying="y", side="right",
                    showgrid=False,
                    range=[0, max(df["Throughput"].max() * 1.5, 130)]),
    )
    return chart_style(fig, height=400), df


def fig_backlog_trend(cr):
    daily_created  = cr.groupby(cr["Created On"].dt.date).size().rename("Created")
    daily_resolved = (cr[~cr["Is Open"]]
                      .groupby(cr["Resolved On"].dt.date).size().rename("Resolved"))
    all_dates = pd.date_range(cr["Created On"].min().date(), TODAY.date(), freq="D")
    df = pd.DataFrame(index=all_dates)
    df.index.name = "Date"
    df = df.join(daily_created, how="left").join(daily_resolved, how="left").fillna(0)
    df["Backlog"] = (df["Created"].cumsum() - df["Resolved"].cumsum()).clip(lower=0)
    df = df.reset_index()
    fig = go.Figure()
    fig.add_scatter(x=df["Date"], y=df["Backlog"],
                    fill="tozeroy", fillcolor="rgba(214,40,40,0.10)",
                    line=dict(color="#d62828", width=2), name="Cumulative Open Backlog",
                    showlegend=True)
    return chart_style(fig, "Running Open Backlog Over Time", height=280)


def fig_ageing_bar(open_df):
    grp = (open_df.groupby(["App", "Age Bucket"], observed=True)
                  .size().reset_index(name="Count"))
    piv = (grp.pivot(index="App", columns="Age Bucket", values="Count")
              .fillna(0).reset_index())
    piv["Total"] = piv[[c for c in AGE_ORDER if c in piv.columns]].sum(axis=1)
    piv = piv.sort_values("Total", ascending=True).tail(18)
    fig = go.Figure()
    for bucket, color in zip(AGE_ORDER, AGE_COLORS):
        if bucket in piv.columns:
            fig.add_bar(y=piv["App"], x=piv[bucket],
                        name=bucket, orientation="h", marker_color=color)
    fig.update_layout(barmode="stack", xaxis_title="Open Incidents",
                      yaxis=dict(tickfont=dict(size=9)))
    return chart_style(fig, height=max(320, len(piv) * 26 + 80))


def fig_resolution_dist(cr):
    df = cr.dropna(subset=["Res Days"]).copy()
    df = df[df["Res Days"] >= 0]
    fig = go.Figure()
    for pri in PRIORITY_ORD:
        sub = df[df["Priority"] == pri]["Res Days"]
        if len(sub) > 0:
            fig.add_box(y=sub, name=pri, marker_color=PRI_COLORS.get(pri, C_BLUE),
                        boxmean="sd", line_width=1.5)
    cap = df["Res Days"].quantile(0.95) * 1.3
    fig.update_yaxes(title="Days to Resolve", range=[0, min(cap, 100)])
    return chart_style(fig, "Resolution Time Distribution by Priority", height=320)


def fig_team_volume(cr):
    grp = (cr.groupby("App")
             .agg(Total=("Number","count"),
                  Open=("Is Open","sum"),
                  Avg_Res=("Res Days","mean"))
             .reset_index()
             .sort_values("Total", ascending=False).head(15))
    grp["Closed"] = grp["Total"] - grp["Open"]
    grp = grp.sort_values("Total", ascending=True)
    fig = go.Figure()
    fig.add_bar(y=grp["App"], x=grp["Closed"], name="Resolved",
                orientation="h", marker_color=C_TEAL, opacity=0.9)
    fig.add_bar(y=grp["App"], x=grp["Open"], name="Open",
                orientation="h", marker_color="#d62828", opacity=0.9)
    fig.update_layout(barmode="stack", xaxis_title="Incidents",
                      yaxis=dict(tickfont=dict(size=9)))
    return chart_style(fig, "Incident Volume by Application (Open vs Resolved)", height=420)


def fig_avg_resolution_by_app(cr):
    df = (cr[cr["Res Days"].notna() & (cr["Res Days"] >= 0)]
            .groupby("App")["Res Days"].mean()
            .reset_index()
            .rename(columns={"Res Days": "Avg Days"})
            .sort_values("Avg Days", ascending=False).head(15))
    df["Avg Days"] = df["Avg Days"].round(1)
    df["Speed"] = df["Avg Days"].apply(
        lambda x: "Slow  >14d" if x > 14 else ("Medium 7-14d" if x > 7 else "Fast  ≤7d"))
    df["Color"] = df["Avg Days"].apply(
        lambda x: "#d62828" if x > 14 else ("#d97706" if x > 7 else "#1a9e4e"))
    fig = px.bar(df, x="App", y="Avg Days", color="Speed",
                 color_discrete_map={
                     "Fast  ≤7d": "#1a9e4e",
                     "Medium 7-14d": "#d97706",
                     "Slow  >14d": "#d62828"},
                 category_orders={"Speed": ["Fast  ≤7d","Medium 7-14d","Slow  >14d"]},
                 text=df["Avg Days"].apply(lambda x: f"{x}d"))
    fig.update_traces(textposition="outside")
    fig.update_xaxes(tickangle=-35)
    return chart_style(fig, "Avg Resolution Time per Application (days)", height=360)


def fig_priority_donut(cr):
    p = cr["Priority"].value_counts().reset_index()
    p.columns = ["Priority", "Count"]
    fig = px.pie(p, names="Priority", values="Count",
                 color="Priority", color_discrete_map=PRI_COLORS, hole=0.55)
    fig.update_traces(textinfo="label+percent", textfont_size=11)
    return chart_style(fig, height=300, legend=False)


def fig_daily_trend(cr):
    daily = cr.groupby(cr["Created On"].dt.date).size().reset_index(name="Count")
    daily.columns = ["Date", "Count"]
    daily = daily.sort_values("Date")
    daily["MA7"] = daily["Count"].rolling(7, min_periods=1).mean().round(1)
    fig = go.Figure()
    fig.add_bar(x=daily["Date"], y=daily["Count"],
                name="Daily Incident Count", marker_color=C_BLUE, opacity=0.55)
    fig.add_scatter(x=daily["Date"], y=daily["MA7"],
                    name="7-Day Rolling Avg", line=dict(color="#d62828", width=2.2), mode="lines")
    return chart_style(fig, "Daily Incident Volume + 7-day Rolling Average", height=300)


def fig_source_split(cr):
    src = cr.groupby(["Created Week Ending", "Source"]).size().reset_index(name="Count")
    src["WLabel"] = src["Created Week Ending"].dt.strftime("%-d %b")
    src = src.sort_values("Created Week Ending")
    fig = px.bar(src, x="WLabel", y="Count", color="Source",
                 color_discrete_map={"Automated": C_TEAL, "Manual": C_NAVY},
                 barmode="stack",
                 labels={"WLabel": "Week Ending", "Count": "Incidents", "Source": "Ticket Source"})
    return chart_style(fig, "Incident Source: Manual vs Automated (Weekly)", height=300)


def fig_heatmap(cr):
    hm = cr.copy()
    hm["Weekday"] = hm["Created On"].dt.day_name()
    hm["WeekOf"]  = hm["Created On"].dt.to_period("W").dt.start_time.dt.strftime("%-d %b")
    piv = (hm.groupby(["Weekday","WeekOf"]).size().reset_index(name="Count")
             .pivot(index="Weekday", columns="WeekOf", values="Count").fillna(0))
    order = [d for d in ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
             if d in piv.index]
    piv = piv.reindex(order)
    fig = px.imshow(piv, color_continuous_scale=["#e8f4fb", C_NAVY],
                    aspect="auto", labels=dict(color="Count"))
    return chart_style(fig, "Creation Heatmap - Weekday x Week", height=280)


def fig_sla_breach(cr):
    df = cr[cr["Is Open"]].copy()
    sla = df.groupby(["App", "SLA Breached"]).size().reset_index(name="Count")
    breached = sla[sla["SLA Breached"]].sort_values("Count", ascending=True).tail(12)
    if breached.empty:
        return None
    fig = px.bar(breached, x="Count", y="App", orientation="h",
                 color_discrete_sequence=["#d62828"], text="Count")
    fig.update_traces(textposition="outside", name="Breached Incidents",
                      showlegend=True)
    return chart_style(fig, "Applications with Most SLA-Breached Open Incidents", height=340)


def fig_opener_bar(cr):
    if "Opened By" not in cr.columns:
        return None
    top = (cr.groupby("Opened By").size()
             .reset_index(name="Count")
             .sort_values("Count", ascending=False).head(12))
    top["Type"] = top["Opened By"].apply(lambda x: "Automated" if is_auto(x) else "Manual")
    top = top.sort_values("Count", ascending=True)
    fig = px.bar(top, x="Count", y="Opened By", orientation="h",
                 color="Type",
                 color_discrete_map={"Automated": C_TEAL, "Manual": C_NAVY},
                 text="Count")
    fig.update_traces(textposition="outside")
    return chart_style(fig, "Top Ticket Openers", height=380)


def fig_tower_volume(cr):
    """Bar chart: incident volume per tower, stacked open/closed."""
    grp = (cr.groupby("Tower")
             .agg(Total=("Number","count"), Open=("Is Open","sum"))
             .reset_index())
    grp["Closed"] = grp["Total"] - grp["Open"]
    grp = grp.sort_values("Total", ascending=True)
    colors = [TOWER_COLORS.get(t, "#9ca3af") for t in grp["Tower"]]
    fig = go.Figure()
    fig.add_bar(y=grp["Tower"], x=grp["Closed"], name="Resolved",
                orientation="h",
                marker=dict(color=[c for c in colors], opacity=0.6))
    fig.add_bar(y=grp["Tower"], x=grp["Open"], name="Open",
                orientation="h",
                marker=dict(color=[c for c in colors], opacity=1.0))
    fig.update_layout(barmode="stack", xaxis_title="Incidents",
                      yaxis=dict(tickfont=dict(size=10)))
    return chart_style(fig, "Incident Volume by Tower (Open vs Resolved)", height=380)


def fig_tower_throughput(cr, re):
    """Weekly throughput table grouped by Tower."""
    if "Tower" not in cr.columns or "Tower" not in re.columns:
        return None
    cr_t = cr.groupby(["Created Week Ending","Tower"]).size().reset_index(name="Created")
    re_t = re.groupby(["Resolved Week Ending","Tower"]).size().reset_index(name="Resolved")
    cr_t = cr_t.rename(columns={"Created Week Ending":"Week"})
    re_t = re_t.rename(columns={"Resolved Week Ending":"Week"})
    merged = pd.merge(cr_t, re_t, on=["Week","Tower"], how="outer").fillna(0)
    merged["Throughput"] = (merged["Resolved"]/merged["Created"].replace(0,np.nan)*100).round(1)
    merged["WLabel"] = pd.to_datetime(merged["Week"]).dt.strftime("%-d %b")
    merged = merged.sort_values(["Tower","Week"])
    fig = px.line(merged, x="WLabel", y="Throughput", color="Tower",
                  color_discrete_map=TOWER_COLORS, markers=True,
                  line_shape="spline")
    fig.add_hline(y=100, line_dash="dash", line_color="#d62828",
                  annotation_text="100% target", annotation_position="bottom right")
    fig.update_yaxes(title="Throughput %", range=[0, 160])
    return chart_style(fig, "Weekly Throughput % by Tower", height=360)


def fig_tower_ageing(open_df):
    """Stacked bar: open ticket ageing by Tower."""
    grp = (open_df.groupby(["Tower","Age Bucket"], observed=True)
                  .size().reset_index(name="Count"))
    piv = (grp.pivot(index="Tower", columns="Age Bucket", values="Count")
              .fillna(0).reset_index())
    piv["Total"] = piv[[c for c in AGE_ORDER if c in piv.columns]].sum(axis=1)
    piv = piv.sort_values("Total", ascending=True)
    fig = go.Figure()
    for bucket, color in zip(AGE_ORDER, AGE_COLORS):
        if bucket in piv.columns:
            fig.add_bar(y=piv["Tower"], x=piv[bucket],
                        name=bucket, orientation="h", marker_color=color)
    fig.update_layout(barmode="stack", xaxis_title="Open Incidents",
                      yaxis=dict(tickfont=dict(size=10)))
    return chart_style(fig, "Open Incident Ageing by Tower", height=340)


def fig_tower_resolution(cr):
    """Box plot: resolution time distribution per tower."""
    df = cr.dropna(subset=["Res Days"]).copy()
    df = df[df["Res Days"] >= 0]
    towers = df.groupby("Tower")["Res Days"].median().sort_values().index.tolist()
    fig = go.Figure()
    for t in towers:
        sub = df[df["Tower"] == t]["Res Days"]
        fig.add_box(y=sub, name=t,
                    marker_color=TOWER_COLORS.get(t,"#9ca3af"),
                    boxmean="sd", line_width=1.5)
    cap = df["Res Days"].quantile(0.95) * 1.3
    fig.update_yaxes(title="Days to Resolve", range=[0, min(cap, 80)])
    return chart_style(fig, "Resolution Time by Tower", height=360)


def fig_tower_sla(cr):
    """SLA breach % per tower."""
    df = cr[cr["Is Open"]].copy()
    if df.empty or "Tower" not in df.columns:
        return None
    t = (df.groupby("Tower")
           .agg(Total=("Number","count"), Breached=("SLA Breached","sum"))
           .reset_index())
    t["Breach %"] = (t["Breached"]/t["Total"]*100).round(1)
    t["Color"] = t["Breach %"].apply(
        lambda x: "#d62828" if x>30 else ("#d97706" if x>10 else "#1a9e4e"))
    t = t.sort_values("Breach %", ascending=True)
    t["Risk"] = t["Breach %"].apply(
        lambda x: "High Risk >30%" if x > 30 else ("Medium Risk 10-30%" if x > 10 else "Low Risk ≤10%"))
    risk_colors = {"High Risk >30%": "#d62828", "Medium Risk 10-30%": "#d97706", "Low Risk ≤10%": "#1a9e4e"}
    fig = px.bar(t, x="Breach %", y="Tower", orientation="h",
                 color="Risk",
                 color_discrete_map=risk_colors,
                 category_orders={"Risk": ["High Risk >30%","Medium Risk 10-30%","Low Risk ≤10%"]},
                 text=t["Breach %"].apply(lambda x: f"{x}%"))
    fig.update_traces(textposition="outside")
    return chart_style(fig, "SLA Breach Rate % by Tower (Open Tickets)", height=320)


# ─────────────────────────────────────────────────────────
#  EXCEL EXPORT
# ─────────────────────────────────────────────────────────
def build_export(cr, re, tp_df):
    buf = io.BytesIO()
    open_df = cr[cr["Is Open"]].copy()
    sla_df  = cr[cr["SLA Breached"]].copy()

    age_pivot = (open_df.groupby(["App","Age Bucket"], observed=True)
                        .size().reset_index(name="Count")
                        .pivot(index="App", columns="Age Bucket", values="Count")
                        .fillna(0).astype(int))
    for b in AGE_ORDER:
        if b not in age_pivot.columns:
            age_pivot[b] = 0
    age_pivot = age_pivot[AGE_ORDER]
    age_pivot["Grand Total"] = age_pivot.sum(axis=1)
    age_pivot = age_pivot.sort_values("Grand Total", ascending=False).reset_index()

    app_summary = (cr.groupby("App")
                     .agg(Total=("Number","count"),
                          Open=("Is Open","sum"),
                          SLA_Breached=("SLA Breached","sum"),
                          Avg_Res_Days=("Res Days","mean"))
                     .round(1).reset_index())

    open_export_cols = [c for c in
        ["Number","Tower","App","Priority","Impact","Created On","Age Days",
         "Age Bucket","SLA Breached","Assignment Group","Opened By"]
        if c in open_df.columns]

    sla_cols = [c for c in
        ["Number","Tower","App","Priority","Created On","Age Days","Assignment Group"]
        if c in sla_df.columns]

    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tp_df[["WLabel","Created","Resolved","Throughput"]].to_excel(
            w, sheet_name="Weekly Throughput", index=False)
        age_pivot.to_excel(w, sheet_name="Open Ageing", index=False)
        open_df[open_export_cols].sort_values(
            ["Priority","Age Days"], ascending=[True,False]).to_excel(
            w, sheet_name="Open Incidents", index=False)
        sla_df[sla_cols].to_excel(w, sheet_name="SLA Breached", index=False)
        app_summary.to_excel(w, sheet_name="App Summary", index=False)
        if "Tower" in cr.columns:
            tower_summary = (cr.groupby("Tower")
                               .agg(Total=("Number","count"),
                                    Open=("Is Open","sum"),
                                    SLA_Breached=("SLA Breached","sum"),
                                    Avg_Res_Days=("Res Days","mean"))
                               .round(1).reset_index())
            tower_summary["Closure Rate %"] = (
                (tower_summary["Total"]-tower_summary["Open"])/tower_summary["Total"]*100
            ).round(1)
            tower_summary.to_excel(w, sheet_name="Tower Summary", index=False)
        cr.to_excel(w, sheet_name="All Created", index=False)
        re.to_excel(w, sheet_name="All Resolved", index=False)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────
#  ME MONTHLY DATA LOADER
# ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_me_monthly(f):
    """Parse the ME Monthly Deck Reporting Excel into a clean DataFrame."""
    name = getattr(f, "name", "")
    if name.endswith(".csv"):
        raw = pd.read_csv(f, header=None)
    else:
        xl = pd.ExcelFile(f)
        first_sheet = xl.sheet_names[0]
        raw = xl.parse(first_sheet, header=None)

    records = []
    i = 4  # data rows start at row 4
    while i < len(raw):
        r0 = raw.iloc[i]
        r1 = raw.iloc[i + 1] if i + 1 < len(raw) else None
        r2 = raw.iloc[i + 2] if i + 2 < len(raw) else None

        if pd.notna(r0.iloc[0]):
            whitemark = r0.iloc[5] if pd.notna(r0.iloc[5]) else None
            if whitemark is None and r2 is not None and pd.notna(r2.iloc[5]):
                whitemark = r2.iloc[5]

            sub_cat   = r1.iloc[1] if r1 is not None and pd.notna(r1.iloc[1]) else None
            on_time_sub = r1.iloc[2] if r1 is not None else None
            abt_sub   = r1.iloc[3] if r1 is not None else None
            ibm_sub   = r1.iloc[4] if r1 is not None else None
            pct_sub   = r1.iloc[6] if r1 is not None else None

            records.append({
                "Month":              pd.to_datetime(r0.iloc[0]),
                "On Time":            r0.iloc[2],
                "ABT Delay":          r0.iloc[3],
                "IBM Delay":          r0.iloc[4],
                "Whitemark":          whitemark,
                "On Time All %":      r0.iloc[6],
                "On Time Critical %": r0.iloc[7],
                "Sub Category":       sub_cat,
                "On Time Sub":        on_time_sub,
                "ABT Delay Sub":      abt_sub,
                "IBM Delay Sub":      ibm_sub,
                "On Time Sub %":      pct_sub,
            })
        i += 3

    df = pd.DataFrame(records)
    df["Month"] = pd.to_datetime(df["Month"])
    df["Total Changes"] = df["On Time"].fillna(0) + df["ABT Delay"].fillna(0) + df["IBM Delay"].fillna(0)
    df["Year"]  = df["Month"].dt.year
    df["MonthName"] = df["Month"].dt.strftime("%b %Y")
    df["MonthShort"] = df["Month"].dt.strftime("%b")
    df["Month Num"] = df["Month"].dt.month
    return df.sort_values("Month").reset_index(drop=True)


# ─────────────────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Data Upload")
    st.markdown("---")
    created_file  = st.file_uploader(
        "Incident **Created** File", type=["csv","xlsx","xls"], key="cr")
    resolved_file = st.file_uploader(
        "Incident **Resolved** File", type=["csv","xlsx","xls"], key="re")
    mapping_file  = st.file_uploader(
        "Queue → Tower **Mapping** File *(optional)*",
        type=["csv","xlsx","xls"], key="mf",
        help="Upload the SNOW Queue vs Tower mapping Excel. If not uploaded, Tower filter will not be available.")
    me_monthly_file = st.file_uploader(
        "ME Monthly Deck Report *(optional)*",
        type=["csv","xlsx","xls"], key="me",
        help="Upload the ME Monthly Deck Reporting Excel for the Change On-Time tab.")
    st.markdown("---")
    st.caption("AMS Operations Dashboard v2.0")


# ─────────────────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────────────────
# Tower pill for header
sel_tow=""
_tower_pill = (f'<span style="background:rgba(0,175,216,0.25);color:#b3e8f7;'
               f'padding:0.18rem 0.7rem;border-radius:12px;font-size:0.72rem;'
               f'font-weight:700;letter-spacing:0.06em;margin-right:0.4rem;">'
               f'🏗 {sel_tow}</span>' if sel_tow != "All" else "")

st.markdown(f"""
<div class="ams-header">
  <div>
    <h1>📊 AMS Operations Dashboard</h1>
    <p>Abbott Application Management Support · ServiceNow Incident Analytics</p>
  </div>
  <div class="ams-right">
    {_tower_pill}<span class="ams-badge">LIVE REPORT</span><br>
    <span class="ams-date">{TODAY.strftime('%A, %d %b %Y')}</span>
  </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────
#  WAIT FOR FILES
# ─────────────────────────────────────────────────────────
if not created_file or not resolved_file:
    _, c2, _ = st.columns([1,2,1])
    with c2:
        st.markdown("""
        <div style="text-align:center;padding:3rem 2rem;background:#f5f9fd;
             border-radius:12px;border:2px dashed #b0cfe0;margin-top:2rem;">
          <div style="font-size:3rem">📁</div>
          <div style="font-size:1.2rem;font-weight:700;color:#003f6b;margin:0.5rem 0">
            Upload Your ServiceNow Data</div>
          <div style="color:#5a7a94;font-size:0.88rem">
            Use the sidebar to upload the <b>Incident Created</b> and
            <b>Incident Resolved</b> CSV or Excel files to get started.
          </div>
        </div>""", unsafe_allow_html=True)
    st.stop()


# ─────────────────────────────────────────────────────────
#  LOAD
# ─────────────────────────────────────────────────────────
with st.spinner("Processing ServiceNow data…"):
    cr_raw, re_raw = load(created_file, resolved_file)

# ── Load tower mapping ──
tower_map = {}
if mapping_file:
    tower_map = load_mapping(mapping_file)
cr_raw = apply_tower_mapping(cr_raw, tower_map)
re_raw = apply_tower_mapping(re_raw, tower_map)

# ── Sidebar filters ──
apps   = ["All"] + sorted(cr_raw["App"].dropna().unique())
pris   = ["All"] + sorted(cr_raw["Priority"].dropna().unique(),
                           key=lambda x: {"High":0,"Moderate":1,"Low":2}.get(x,9))
groups = ["All"] + sorted(cr_raw["Assignment Group"].dropna().unique())
towers = ["All"] + sorted([t for t in cr_raw["Tower"].dropna().unique() if t != "Unmapped"]) + (
    ["Unmapped"] if "Unmapped" in cr_raw["Tower"].values else [])
min_d  = cr_raw["Created On"].min().date()
max_d  = cr_raw["Created On"].max().date()

with st.sidebar:
    st.markdown("### ⚙️ Filters")

    # Tower — primary slice (most important)
    if mapping_file:
        st.markdown("""<div style="background:#0f3a5e;border-left:3px solid #00afd8;
            border-radius:6px;padding:0.45rem 0.75rem;margin-bottom:0.6rem;">
            <span style="color:#7eb8d8;font-size:0.68rem;font-weight:700;
            text-transform:uppercase;letter-spacing:0.08em;">🏗 Tower</span>
            </div>""", unsafe_allow_html=True)
        sel_tow = st.selectbox("Tower", towers, key="f_tow")
    else:
        sel_tow = "All"
        st.markdown("""<div style="background:#1e3348;border-radius:6px;
            padding:0.5rem 0.75rem;margin-bottom:0.6rem;font-size:0.75rem;color:#7eb8d8;">
            📌 Upload mapping file to enable Tower filter</div>""",
            unsafe_allow_html=True)

    st.markdown("---")
    sel_app   = st.selectbox("Application",      apps,   key="f_app")
    sel_pri   = st.selectbox("Priority",         pris,   key="f_pri")
    sel_grp   = st.selectbox("Assignment Group", groups, key="f_grp")
    date_rng  = st.date_input("Date Range",
                               value=(min_d, max_d),
                               min_value=min_d, max_value=max_d,
                               key="f_date")
    show_auto = st.checkbox("Include Auto-created tickets", value=True, key="f_auto")

    # Show active filter summary
    active = [f for f in [
        f"Tower: {sel_tow}" if sel_tow!="All" else "",
        f"App: {sel_app}" if sel_app!="All" else "",
        f"Priority: {sel_pri}" if sel_pri!="All" else "",
        f"Group: {sel_grp[:20]}..." if sel_grp!="All" else "",
    ] if f]
    if active:
        st.markdown("**Active filters:**")
        for a in active:
            st.markdown(f"<span style='background:#0f3a5e;color:#7eb8d8;padding:0.15rem 0.5rem;"
                        f"border-radius:4px;font-size:0.73rem;display:inline-block;"
                        f"margin:0.1rem 0'>{a}</span>", unsafe_allow_html=True)

# ── Apply filters ──
cr = cr_raw.copy()
re = re_raw.copy()

if sel_tow != "All":
    cr = cr[cr["Tower"] == sel_tow]
    re = re[re["Tower"] == sel_tow]
if sel_app != "All":
    cr = cr[cr["App"] == sel_app]
    re = re[re["App"] == sel_app]
if sel_pri != "All":
    cr = cr[cr["Priority"] == sel_pri]
    re = re[re["Priority"] == sel_pri]
if sel_grp != "All":
    cr = cr[cr["Assignment Group"] == sel_grp]
    re = re[re["Assignment Group"] == sel_grp]
if not show_auto:
    cr = cr[cr["Source"] == "Manual"]
if isinstance(date_rng, (list, tuple)) and len(date_rng) == 2:
    s, e = pd.Timestamp(date_rng[0]), pd.Timestamp(date_rng[1])
    cr = cr[(cr["Created On"] >= s) & (cr["Created On"] <= e)]
    re = re[(re["Resolved On"] >= s) & (re["Resolved On"] <= e)]

open_df = cr[cr["Is Open"]].copy()

# ── KPIs ──
total_cr  = len(cr)
total_re  = len(re)
total_op  = int(cr["Is Open"].sum())
pct_tp    = round(total_re / total_cr * 100, 1) if total_cr else 0
avg_res_v = cr["Res Days"].mean()
avg_res   = round(avg_res_v, 1) if not np.isnan(avg_res_v) else 0
sla_br    = int(cr["SLA Breached"].sum())
high_open = int((open_df["Priority"] == "High").sum())
mod_open  = int((open_df["Priority"] == "Moderate").sum())

cr["WeekNum"] = cr["Created On"].dt.isocalendar().week
wks = sorted(cr["WeekNum"].dropna().unique())
wow_delta = (len(cr[cr["WeekNum"] == wks[-1]]) -
             len(cr[cr["WeekNum"] == wks[-2]])) if len(wks) >= 2 else 0

tp_class = "green" if pct_tp >= 100 else ("amber" if pct_tp >= 85 else "red")
tower_ctx = f" &nbsp;·&nbsp; 🏗 Tower: <b>{sel_tow}</b>" if sel_tow != "All" else ""

def delta_html(v, invert=False):
    if v == 0:
        return '<span class="delta-neu">No change vs prev week</span>'
    arrow = "↑" if v > 0 else "↓"
    cls = ("delta-up" if v > 0 else "delta-dn") if not invert else ("delta-dn" if v > 0 else "delta-up")
    return f'<span class="{cls}">{arrow} {abs(v):,} vs prev week</span>'

st.markdown(f"""
<div class="kpi-row">
  <div class="kpi">
    <div class="kpi-label">Total Created</div>
    <div class="kpi-value">{total_cr:,}</div>
    <div class="kpi-delta">{delta_html(wow_delta)}</div>
  </div>
  <div class="kpi green">
    <div class="kpi-label">Total Resolved</div>
    <div class="kpi-value">{total_re:,}</div>
    <div class="kpi-delta"><span class="delta-neu">Across date range</span></div>
  </div>
  <div class="kpi red">
    <div class="kpi-label">Open Backlog</div>
    <div class="kpi-value">{total_op:,}</div>
    <div class="kpi-delta"><span class="{'delta-up' if high_open>0 else 'delta-neu'}">
      {high_open} High &middot; {mod_open} Moderate</span></div>
  </div>
  <div class="kpi {tp_class}">
    <div class="kpi-label">Throughput</div>
    <div class="kpi-value">{pct_tp}%</div>
    <div class="kpi-delta"><span class="{'delta-dn' if pct_tp>=100 else 'delta-up'}">
      Resolved / Created</span></div>
  </div>
  <div class="kpi teal">
    <div class="kpi-label">Avg Resolution</div>
    <div class="kpi-value">{avg_res}d</div>
    <div class="kpi-delta"><span class="{'delta-up' if avg_res>10 else 'delta-dn'}">
      {'Above' if avg_res>10 else 'Within'} 10-day target</span></div>
  </div>
  <div class="kpi {'red' if sla_br>0 else 'green'}">
    <div class="kpi-label">SLA Breached</div>
    <div class="kpi-value">{sla_br:,}</div>
    <div class="kpi-delta"><span class="{'delta-up' if sla_br>0 else 'delta-dn'}">
      {'Needs attention' if sla_br>0 else 'All within SLA'}</span></div>
  </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────
#  TABS
# ─────────────────────────────────────────────────────────
(tab_action, tab_weekly, tab_age,
 tab_team, tab_trends, tab_drill, tab_me, tab_export) = st.tabs([
    "🚨 Action Required",
    "📅 Weekly Throughput",
    "⏳ Ageing & Backlog",
    "👥 Team & App Performance",
    "📈 Trends & Analytics",
    "🔍 Incident Drill-down",
    "📋 ME Monthly Report",
    "⬇️ Export",
])


# ═══════════════════════════════════════════════
#  TAB 1 – ACTION REQUIRED
# ═══════════════════════════════════════════════
with tab_action:
    st.markdown('<div class="sec">Immediate Attention Required</div>',
                unsafe_allow_html=True)

    if high_open > 0:
        st.markdown(f"""
        <div class="alert-critical">
          <div class="alert-title">🔴 {high_open} HIGH Priority Incident(s) Open</div>
          <div class="alert-body">Requires immediate escalation — target resolution within 1 business day.</div>
        </div>""", unsafe_allow_html=True)

    if sla_br > 0:
        st.markdown(f"""
        <div class="alert-warn">
          <div class="alert-title">⚠️ {sla_br} Incident(s) Have Breached SLA Target</div>
          <div class="alert-body">These tickets exceed priority-based resolution targets.
            Review and update stakeholders immediately.</div>
        </div>""", unsafe_allow_html=True)

    old_60 = int((open_df["Age Days"] > 60).sum())
    if old_60 > 0:
        st.markdown(f"""
        <div class="alert-warn">
          <div class="alert-title">📅 {old_60} Open Ticket(s) Older than 60 Days</div>
          <div class="alert-body">Long-running incidents may be stale.
            Confirm still active or close if resolved.</div>
        </div>""", unsafe_allow_html=True)

    if high_open == 0 and sla_br == 0 and old_60 == 0:
        st.markdown("""
        <div class="alert-ok">
          <div class="alert-title">✅ No Critical Actions Required</div>
          <div class="alert-body">All open incidents are within acceptable parameters.</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="sec">Open Incidents — Priority View (Top 30)</div>',
                unsafe_allow_html=True)

    if open_df.empty:
        st.success("No open incidents in the current selection.")
    else:
        top_open = (open_df
                    .sort_values(["Pri Sort","Age Days"], ascending=[True,False])
                    .head(30))
        for _, row in top_open.iterrows():
            pri = row["Priority"]
            age = int(row["Age Days"])
            card_cls = "high" if pri=="High" else ("mod" if pri=="Moderate" else "low")
            age_cls  = "age-red" if age>60 else ("age-amb" if age>30 else "age-grn")
            pri_icon = "🔴" if pri=="High" else ("🟡" if pri=="Moderate" else "🔵")
            created_str = row["Created On"].strftime("%d %b %Y") if pd.notna(row["Created On"]) else "N/A"
            sla_flag = "&nbsp;|&nbsp; 🚨 <b>SLA BREACHED</b>" if row.get("SLA Breached") else ""
            grp_str  = str(row.get("Assignment Group",""))[:55]
            st.markdown(f"""
            <div class="inc-card {card_cls}">
              <span class="inc-num">{row['Number']}</span>
              <span class="inc-age {age_cls}">{age}d old</span><br>
              <span class="inc-app">{row['App']}</span><br>
              <span class="inc-meta">
                {pri_icon} {pri} &nbsp;|&nbsp; 🗂 {grp_str}
                &nbsp;|&nbsp; 📅 {created_str}{sla_flag}
              </span>
            </div>""", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sec">Age Distribution of Open Tickets</div>',
                    unsafe_allow_html=True)
        if not open_df.empty:
            age_cnt = open_df["Age Bucket"].value_counts().reindex(AGE_ORDER).fillna(0)
            fig_aged = go.Figure()
            for bucket, color in zip(AGE_ORDER, AGE_COLORS):
                val = age_cnt.get(bucket, 0)
                fig_aged.add_bar(
                    x=[bucket], y=[val],
                    name=bucket, marker_color=color,
                    text=[int(val)], textposition="outside",
                    showlegend=True)
            st.plotly_chart(chart_style(fig_aged, height=280),
                            use_container_width=True,
                            config={"displayModeBar": False},
                            key="act_age_dist")
    with c2:
        st.markdown('<div class="sec">SLA Breach by Application</div>',
                    unsafe_allow_html=True)
        sf = fig_sla_breach(cr)
        if sf:
            st.plotly_chart(sf, use_container_width=True,
                            config={"displayModeBar": False}, key="act_sla")
        else:
            st.success("No SLA breaches in current selection.")


# ═══════════════════════════════════════════════
#  TAB 2 – WEEKLY THROUGHPUT
# ═══════════════════════════════════════════════
with tab_weekly:
    st.markdown('<div class="sec">Weekly Throughput — Created vs Resolved</div>',
                unsafe_allow_html=True)
    fig_tp, tp_df = fig_weekly_throughput(cr, re)
    st.plotly_chart(fig_tp, use_container_width=True,
                    config={"displayModeBar": False}, key="wk_main")

    st.markdown('<div class="sec">Week-by-Week Summary</div>',
                unsafe_allow_html=True)
    rows = []
    for _, r in tp_df.iterrows():
        tp  = r["Throughput"]
        cls = "bg-g" if tp >= 100 else ("bg-y" if tp >= 85 else "bg-r")
        delta = int(r["Created"]) - int(r["Resolved"])
        delta_str = f'+{delta}' if delta > 0 else str(delta)
        rows.append(f"""<tr>
          <td>{r['WLabel']}</td>
          <td>{int(r['Created'])}</td>
          <td>{int(r['Resolved'])}</td>
          <td style="color:{'#d62828' if delta>0 else '#1a9e4e'};font-weight:600">{delta_str}</td>
          <td><span class="badge {cls}">{tp:.0f}%</span></td>
        </tr>""")
    tot_c = int(tp_df["Created"].sum())
    tot_r = int(tp_df["Resolved"].sum())
    tot_t = round(tot_r/tot_c*100,1) if tot_c else 0
    tot_d = tot_c - tot_r
    cls_t = "bg-g" if tot_t>=100 else ("bg-y" if tot_t>=85 else "bg-r")
    rows.append(f"""<tr>
      <td><b>TOTAL</b></td><td><b>{tot_c}</b></td><td><b>{tot_r}</b></td>
      <td style="color:{'#d62828' if tot_d>0 else '#1a9e4e'};font-weight:700">
        {'+' if tot_d>0 else ''}{tot_d}</td>
      <td><span class="badge {cls_t}"><b>{tot_t:.0f}%</b></span></td>
    </tr>""")

    st.markdown(f"""
    <table class="tp-table">
      <thead><tr>
        <th>Week Ending</th><th>Created</th><th>Resolved</th>
        <th>Net (Cr-Re)</th><th>Throughput</th>
      </tr></thead>
      <tbody>{"".join(rows)}</tbody>
    </table>""", unsafe_allow_html=True)

    st.markdown('<div class="sec">Incident Source Split — Manual vs Automated</div>',
                unsafe_allow_html=True)
    st.plotly_chart(fig_source_split(cr), use_container_width=True,
                    config={"displayModeBar": False}, key="wk_source")


# ═══════════════════════════════════════════════
#  TAB 3 – AGEING & BACKLOG
# ═══════════════════════════════════════════════
with tab_age:
    cols5 = st.columns(5)
    for col, bucket, color in zip(cols5, AGE_ORDER, AGE_COLORS):
        cnt = int((open_df["Age Bucket"] == bucket).sum()) if not open_df.empty else 0
        col.markdown(f"""
        <div style="background:#fff;border-top:3px solid {color};border-radius:8px;
             padding:0.7rem 0.9rem;text-align:center;border:1px solid #dde8f0;
             box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <div style="font-size:0.63rem;font-weight:700;color:#5a7a94;
               text-transform:uppercase;letter-spacing:0.08em;margin-bottom:0.2rem;">{bucket}</div>
          <div style="font-size:1.7rem;font-weight:700;color:#0d1b2a;">{cnt}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="sec">Running Open Backlog Trend</div>',
                unsafe_allow_html=True)
    st.plotly_chart(fig_backlog_trend(cr), use_container_width=True,
                    config={"displayModeBar": False}, key="age_backlog")

    if mapping_file:
        st.markdown('<div class="sec">Open Incident Ageing by Tower</div>',
                    unsafe_allow_html=True)
        if not open_df.empty:
            st.plotly_chart(fig_tower_ageing(open_df), use_container_width=True,
                            config={"displayModeBar": False}, key="age_tower_bar")

    st.markdown('<div class="sec">Open Incident Ageing by Application</div>',
                unsafe_allow_html=True)
    if open_df.empty:
        st.success("No open incidents.")
    else:
        st.plotly_chart(fig_ageing_bar(open_df), use_container_width=True,
                        config={"displayModeBar": False}, key="age_bar")

    st.markdown('<div class="sec">Ageing Summary Table</div>',
                unsafe_allow_html=True)
    if not open_df.empty:
        age_piv = (open_df.groupby(["App","Age Bucket"], observed=True)
                          .size().reset_index(name="Count")
                          .pivot(index="App", columns="Age Bucket", values="Count")
                          .fillna(0).astype(int))
        for b in AGE_ORDER:
            if b not in age_piv.columns:
                age_piv[b] = 0
        age_piv = age_piv[AGE_ORDER]
        age_piv["Grand Total"] = age_piv.sum(axis=1)
        st.dataframe(age_piv.sort_values("Grand Total", ascending=False),
                     use_container_width=True)


# ═══════════════════════════════════════════════
#  TAB 4 – TEAM & APP PERFORMANCE
# ═══════════════════════════════════════════════
with tab_team:
    # ── Tower overview section (only when mapping loaded) ──
    if mapping_file and sel_tow == "All":
        st.markdown('<div class="sec">Tower-Level Overview</div>', unsafe_allow_html=True)
        tow_c1, tow_c2 = st.columns(2)
        with tow_c1:
            st.plotly_chart(fig_tower_volume(cr), use_container_width=True,
                            config={"displayModeBar": False}, key="tow_vol")
        with tow_c2:
            sf_tow = fig_tower_sla(cr)
            if sf_tow:
                st.plotly_chart(sf_tow, use_container_width=True,
                                config={"displayModeBar": False}, key="tow_sla")

        tow_c3, tow_c4 = st.columns(2)
        with tow_c3:
            st.plotly_chart(fig_tower_ageing(open_df), use_container_width=True,
                            config={"displayModeBar": False}, key="tow_age")
        with tow_c4:
            st.plotly_chart(fig_tower_resolution(cr), use_container_width=True,
                            config={"displayModeBar": False}, key="tow_res")

        tp_tow = fig_tower_throughput(cr, re)
        if tp_tow:
            st.plotly_chart(tp_tow, use_container_width=True,
                            config={"displayModeBar": False}, key="tow_tp")

        # Tower summary table
        st.markdown('<div class="sec">Tower Performance Summary</div>',
                    unsafe_allow_html=True)
        tow_sum = (cr.groupby("Tower")
                     .agg(Total=("Number","count"),
                          Open=("Is Open","sum"),
                          SLA_Breached=("SLA Breached","sum"),
                          Avg_Res_Days=("Res Days","mean"))
                     .round(1).reset_index()
                     .sort_values("Total", ascending=False))
        tow_sum["Closure Rate %"] = (
            (tow_sum["Total"]-tow_sum["Open"])/tow_sum["Total"]*100
        ).round(1)
        tow_sum["Open"] = tow_sum["Open"].astype(int)
        tow_sum["SLA_Breached"] = tow_sum["SLA_Breached"].astype(int)
        st.dataframe(tow_sum, use_container_width=True)
        st.markdown("---")

    st.plotly_chart(fig_team_volume(cr), use_container_width=True,
                    config={"displayModeBar": False}, key="team_vol")

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(fig_avg_resolution_by_app(cr), use_container_width=True,
                        config={"displayModeBar": False}, key="team_resapp")
    with c2:
        st.plotly_chart(fig_resolution_dist(cr), use_container_width=True,
                        config={"displayModeBar": False}, key="team_resdist")

    opener_f = fig_opener_bar(cr)
    if opener_f:
        st.plotly_chart(opener_f, use_container_width=True,
                        config={"displayModeBar": False}, key="team_opener")

    st.markdown('<div class="sec">Application Performance Summary Table</div>',
                unsafe_allow_html=True)
    app_sum = (cr.groupby("App")
                 .agg(Total=("Number","count"),
                      Open=("Is Open","sum"),
                      SLA_Breached=("SLA Breached","sum"),
                      Avg_Res_Days=("Res Days","mean"))
                 .round(1).reset_index()
                 .sort_values("Total", ascending=False))
    app_sum["Closure Rate %"] = (
        (app_sum["Total"] - app_sum["Open"]) / app_sum["Total"] * 100
    ).round(1)
    app_sum["Open"] = app_sum["Open"].astype(int)
    app_sum["SLA_Breached"] = app_sum["SLA_Breached"].astype(int)
    st.dataframe(app_sum, use_container_width=True, height=340)


# ═══════════════════════════════════════════════
#  TAB 5 – TRENDS & ANALYTICS
# ═══════════════════════════════════════════════
with tab_trends:
    st.plotly_chart(fig_daily_trend(cr), use_container_width=True,
                    config={"displayModeBar": False}, key="tr_daily")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sec">Priority Distribution</div>',
                    unsafe_allow_html=True)
        st.plotly_chart(fig_priority_donut(cr), use_container_width=True,
                        config={"displayModeBar": False}, key="tr_pri")

        pw = (cr.groupby(["Created Week Ending","Priority"])
                .size().reset_index(name="Count"))
        pw["WLabel"] = pw["Created Week Ending"].dt.strftime("%-d %b")
        pw = pw.sort_values("Created Week Ending")
        fig_pw = px.line(pw, x="WLabel", y="Count", color="Priority",
                         color_discrete_map=PRI_COLORS, markers=True)
        st.plotly_chart(chart_style(fig_pw, "Priority Volume by Week", height=280),
                        use_container_width=True,
                        config={"displayModeBar": False}, key="tr_priweek")

    with c2:
        st.markdown('<div class="sec">Creation Heatmap</div>',
                    unsafe_allow_html=True)
        st.plotly_chart(fig_heatmap(cr), use_container_width=True,
                        config={"displayModeBar": False}, key="tr_heat")

        if "Impact" in cr.columns:
            imp = cr["Impact"].value_counts().reset_index()
            imp.columns = ["Impact", "Count"]
            fig_imp = px.bar(imp, x="Impact", y="Count", color="Impact",
                             color_discrete_map={
                                 "Critical":"#d62828","High":"#d97706",
                                 "Moderate":"#0078b6","Low":"#1a9e4e"},
                             category_orders={"Impact": ["Critical","High","Moderate","Low"]},
                             labels={"Count": "Incidents", "Impact": "Impact Level"},
                             text="Count")
            fig_imp.update_traces(textposition="outside")
            st.plotly_chart(chart_style(fig_imp, "Impact Distribution", height=280),
                            use_container_width=True,
                            config={"displayModeBar": False}, key="tr_impact")


# ═══════════════════════════════════════════════
#  TAB 6 – INCIDENT DRILL-DOWN
# ═══════════════════════════════════════════════
with tab_drill:
    st.markdown('<div class="sec">Interactive Incident Explorer</div>',
                unsafe_allow_html=True)

    d1, d2, d3 = st.columns(3)
    with d1:
        view_mode = st.radio("View Mode",
                             ["Open Only","All Incidents","Resolved Only"],
                             horizontal=True, key="drill_mode")
    with d2:
        drill_pri = st.multiselect(
            "Priority Filter",
            cr["Priority"].dropna().unique().tolist(),
            default=[], key="drill_pri")
    with d3:
        max_age = int(cr["Age Days"].max()) if len(cr) > 0 else 200
        age_min, age_max = st.slider(
            "Age Range (days)", 0, max_age, (0, max_age), key="drill_age")

    drill_df = cr.copy()
    if view_mode == "Open Only":
        drill_df = drill_df[drill_df["Is Open"]]
    elif view_mode == "Resolved Only":
        drill_df = drill_df[~drill_df["Is Open"]]
    if drill_pri:
        drill_df = drill_df[drill_df["Priority"].isin(drill_pri)]
    drill_df = drill_df[
        (drill_df["Age Days"] >= age_min) & (drill_df["Age Days"] <= age_max)]

    st.caption(f"Showing **{len(drill_df):,}** incidents")

    show_cols = [c for c in
        ["Number","Tower","App","Priority","Impact","Created On",
         "Age Days","Age Bucket","SLA Breached","Is Open",
         "Assignment Group","Opened By"]
        if c in drill_df.columns]

    st.dataframe(
        drill_df[show_cols]
        .sort_values(["Priority","Age Days"], ascending=[True,False])
        .reset_index(drop=True),
        use_container_width=True, height=480)

    if len(drill_df) > 0:
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Selected Count",  f"{len(drill_df):,}")
        s2.metric("Open",            f"{drill_df['Is Open'].sum():,}")
        s3.metric("SLA Breached",    f"{drill_df['SLA Breached'].sum():,}")
        avg_r = drill_df["Res Days"].mean()
        s4.metric("Avg Res Days", f"{avg_r:.1f}d" if not np.isnan(avg_r) else "N/A")


# ═══════════════════════════════════════════════
#  TAB 7 – EXPORT
# ═══════════════════════════════════════════════
# ═══════════════════════════════════════════════
#  TAB ME – ME MONTHLY REPORT
# ═══════════════════════════════════════════════
with tab_me:
    if not me_monthly_file:
        st.markdown("""
        <div style="text-align:center;padding:3rem 2rem;background:#f5f9fd;
             border-radius:12px;border:2px dashed #b0cfe0;margin-top:1rem;">
          <div style="font-size:2.5rem">📋</div>
          <div style="font-size:1.1rem;font-weight:700;color:#003f6b;margin:0.5rem 0">
            Upload ME Monthly Deck Report</div>
          <div style="color:#5a7a94;font-size:0.85rem">
            Use the sidebar to upload <b>ME_Monthly_Deck_Reporting.xlsx</b>
            to enable the Change On-Time executive reporting tab.
          </div>
        </div>""", unsafe_allow_html=True)
    else:
        with st.spinner("Parsing ME Monthly data…"):
            me = load_me_monthly(me_monthly_file)

        # ── Filters row ──────────────────────────────────────────────────
        years_avail = sorted(me["Year"].unique())
        mf1, mf2, mf3 = st.columns([2, 2, 3])
        with mf1:
            yr_opts = ["All Years"] + [str(y) for y in years_avail]
            sel_yr = st.selectbox("Year", yr_opts, key="me_yr")
        with mf2:
            view_n = st.selectbox("Show Last N Months",
                                  ["All", "12 Months", "24 Months", "36 Months"],
                                  key="me_n")
        with mf3:
            sub_cats = ["All"] + sorted(me["Sub Category"].dropna().unique().tolist())
            sel_sub = st.selectbox("Category Filter", sub_cats, key="me_sub")

        me_f = me.copy()
        if sel_yr != "All Years":
            me_f = me_f[me_f["Year"] == int(sel_yr)]
        if view_n == "12 Months":
            me_f = me_f.tail(12)
        elif view_n == "24 Months":
            me_f = me_f.tail(24)
        elif view_n == "36 Months":
            me_f = me_f.tail(36)

        # ── KPI strip ────────────────────────────────────────────────────
        tot_changes   = int(me_f["Total Changes"].sum())
        tot_ontime    = int(me_f["On Time"].fillna(0).sum())
        tot_abt_delay = int(me_f["ABT Delay"].fillna(0).sum())
        tot_ibm_delay = int(me_f["IBM Delay"].fillna(0).sum())
        avg_monthly   = round(me_f["Total Changes"].mean(), 1)
        pct_ontime    = round(tot_ontime / tot_changes * 100, 1) if tot_changes else 0
        whitemark_val = me_f["Whitemark"].dropna().iloc[-1] if me_f["Whitemark"].notna().any() else "N/A"

        st.markdown(f"""
        <div class="kpi-row" style="grid-template-columns:repeat(5,1fr);">
          <div class="kpi green">
            <div class="kpi-label">Total Changes</div>
            <div class="kpi-value">{tot_changes:,}</div>
            <div class="kpi-delta"><span class="delta-neu">In selected period</span></div>
          </div>
          <div class="kpi green">
            <div class="kpi-label">On Time</div>
            <div class="kpi-value">{tot_ontime:,}</div>
            <div class="kpi-delta"><span class="delta-dn">{pct_ontime}% on-time rate</span></div>
          </div>
          <div class="kpi {'red' if tot_abt_delay>0 else 'green'}">
            <div class="kpi-label">ABT Delays</div>
            <div class="kpi-value">{tot_abt_delay:,}</div>
            <div class="kpi-delta"><span class="{'delta-up' if tot_abt_delay>0 else 'delta-dn'}">
              Client-side delays</span></div>
          </div>
          <div class="kpi {'red' if tot_ibm_delay>0 else 'green'}">
            <div class="kpi-label">IBM Delays</div>
            <div class="kpi-value">{tot_ibm_delay:,}</div>
            <div class="kpi-delta"><span class="{'delta-up' if tot_ibm_delay>0 else 'delta-dn'}">
              Vendor delays</span></div>
          </div>
          <div class="kpi teal">
            <div class="kpi-label">Avg / Month</div>
            <div class="kpi-value">{avg_monthly}</div>
            <div class="kpi-delta"><span class="delta-neu">Whitemark: {whitemark_val}</span></div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # ── Chart 1: Monthly Change Volume (bar) + On-Time % (line) ──────
        st.markdown('<div class="sec">Monthly Change Volume & On-Time Performance</div>',
                    unsafe_allow_html=True)

        fig_me_main = go.Figure()
        fig_me_main.add_bar(x=me_f["MonthName"], y=me_f["On Time"],
                            name="✅ On Time", marker_color="#1a9e4e", opacity=0.88)
        fig_me_main.add_bar(x=me_f["MonthName"], y=me_f["ABT Delay"],
                            name="⚠️ ABT Delay (Client)", marker_color="#d97706", opacity=0.9)
        fig_me_main.add_bar(x=me_f["MonthName"], y=me_f["IBM Delay"],
                            name="🔴 IBM Delay (Vendor)", marker_color="#d62828", opacity=0.9)
        fig_me_main.add_scatter(
            x=me_f["MonthName"], y=me_f["On Time All %"],
            name="On Time All %", yaxis="y2", mode="lines+markers",
            line=dict(color=C_BLUE, width=2.5),
            marker=dict(size=6, color=C_BLUE))
        fig_me_main.add_scatter(
            x=me_f["MonthName"], y=me_f["On Time Critical %"],
            name="On Time Critical %", yaxis="y2", mode="lines+markers",
            line=dict(color="#7c3aed", width=2, dash="dot"),
            marker=dict(size=5, color="#7c3aed"))
        # Whitemark reference line + legend entry
        if me_f["Whitemark"].notna().any():
            wm = me_f["Whitemark"].dropna().iloc[-1]
            fig_me_main.add_hline(y=wm, line_dash="dash",
                                  line_color="#9b1d20", line_width=1.5)
            # Dummy scatter so Whitemark appears in the legend
            fig_me_main.add_scatter(
                x=[None], y=[None], mode="lines",
                name=f"── Whitemark ({wm})",
                line=dict(color="#9b1d20", width=2, dash="dash"),
                yaxis="y2", showlegend=True)
        fig_me_main.update_layout(
            barmode="stack",
            yaxis=dict(title="Number of Changes", gridcolor="#dde8f0"),
            yaxis2=dict(title="On-Time %", overlaying="y", side="right",
                        showgrid=False, range=[0, 115]),
            xaxis=dict(tickangle=-40, gridcolor="#dde8f0"),
            height=440,
            margin=dict(l=8, r=8, t=60, b=8),
            paper_bgcolor="white", plot_bgcolor="#f5f9fd",
            font=dict(family="IBM Plex Sans", size=11, color="#3a4a5c"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02,
                        xanchor="right", x=1, font=dict(size=10),
                        title=dict(text="", font=dict(size=10))),
        )
        st.plotly_chart(fig_me_main, use_container_width=True,
                        config={"displayModeBar": False}, key="me_main")

        # ── Charts row 2 ──────────────────────────────────────────────────
        mc1, mc2 = st.columns(2)

        with mc1:
            st.markdown('<div class="sec">Annual Change Volume (Year-over-Year)</div>',
                        unsafe_allow_html=True)
            yr_grp = (me_f.groupby("Year")
                          .agg(Total=("Total Changes","sum"),
                               OnTime=("On Time","sum"),
                               ABT=("ABT Delay","sum"),
                               IBM=("IBM Delay","sum"),
                               Months=("Month","count"))
                          .reset_index())
            yr_grp["Avg/Month"] = (yr_grp["Total"] / yr_grp["Months"]).round(1)
            fig_yr = go.Figure()
            fig_yr.add_bar(x=yr_grp["Year"].astype(str), y=yr_grp["OnTime"],
                           name="✅ On Time", marker_color="#1a9e4e", opacity=0.88)
            fig_yr.add_bar(x=yr_grp["Year"].astype(str), y=yr_grp["ABT"],
                           name="⚠️ ABT Delay", marker_color="#d97706")
            fig_yr.add_bar(x=yr_grp["Year"].astype(str), y=yr_grp["IBM"],
                           name="🔴 IBM Delay", marker_color="#d62828")
            fig_yr.add_scatter(x=yr_grp["Year"].astype(str), y=yr_grp["Avg/Month"],
                               name="📈 Avg/Month", yaxis="y2", mode="lines+markers",
                               line=dict(color=C_BLUE, width=2),
                               marker=dict(size=7, color=C_BLUE))
            fig_yr.update_layout(
                barmode="stack",
                yaxis=dict(title="Total Changes"),
                yaxis2=dict(title="Avg per Month", overlaying="y", side="right",
                            showgrid=False),
                height=340, margin=dict(l=8,r=8,t=20,b=8),
                paper_bgcolor="white", plot_bgcolor="#f5f9fd",
                font=dict(family="IBM Plex Sans", size=11, color="#3a4a5c"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1, font=dict(size=10)),
            )
            st.plotly_chart(fig_yr, use_container_width=True,
                            config={"displayModeBar": False}, key="me_yoy")

        with mc2:
            st.markdown('<div class="sec">Monthly Seasonality — Avg Changes by Calendar Month</div>',
                        unsafe_allow_html=True)
            season = (me_f.groupby("Month Num")["Total Changes"].mean().reset_index())
            month_labels = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                            7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
            season["Month Label"] = season["Month Num"].map(month_labels)
            # Bucket each month into volume tiers for a named legend
            q_lo = season["Total Changes"].quantile(0.33)
            q_hi = season["Total Changes"].quantile(0.67)
            def _vol_tier(v):
                if v >= q_hi: return "High Volume"
                elif v >= q_lo: return "Medium Volume"
                return "Low Volume"
            season["Volume"] = season["Total Changes"].apply(_vol_tier)
            vol_colors = {"Low Volume": "#b3d9f0", "Medium Volume": C_BLUE,
                          "High Volume": C_NAVY}
            vol_order  = ["Low Volume", "Medium Volume", "High Volume"]
            fig_sea = px.bar(season, x="Month Label", y="Total Changes",
                             color="Volume",
                             color_discrete_map=vol_colors,
                             category_orders={"Volume": vol_order,
                                              "Month Label": list(month_labels.values())},
                             labels={"Total Changes": "Avg Monthly Changes",
                                     "Month Label": "Calendar Month",
                                     "Volume": "Volume Tier"},
                             text=season["Total Changes"].round(1))
            fig_sea.update_traces(textposition="outside", texttemplate="%{text:.1f}")
            fig_sea.update_layout(
                height=340, margin=dict(l=8,r=8,t=55,b=8),
                paper_bgcolor="white", plot_bgcolor="#f5f9fd",
                font=dict(family="IBM Plex Sans", size=11, color="#3a4a5c"),
                xaxis=dict(gridcolor="#dde8f0"),
                yaxis=dict(title="Avg Changes", gridcolor="#dde8f0"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1, font=dict(size=10),
                            title=dict(text="Volume Tier", font=dict(size=10))),
                showlegend=True,
            )
            st.plotly_chart(fig_sea, use_container_width=True,
                            config={"displayModeBar": False}, key="me_season")

        # ── Charts row 3: Delay analysis + Sub-category ───────────────────
        mc3, mc4 = st.columns(2)

        with mc3:
            st.markdown('<div class="sec">Delay Root-Cause — ABT vs IBM (Monthly)</div>',
                        unsafe_allow_html=True)
            delay_df = me_f[(me_f["ABT Delay"] > 0) | (me_f["IBM Delay"] > 0)].copy()
            if delay_df.empty:
                st.markdown("""
                <div class="alert-ok">
                  <div class="alert-title">✅ Zero Delays in Selected Period</div>
                  <div class="alert-body">No ABT or IBM delays recorded.</div>
                </div>""", unsafe_allow_html=True)
            else:
                fig_delay = go.Figure()
                fig_delay.add_bar(x=delay_df["MonthName"], y=delay_df["ABT Delay"],
                                  name="ABT Delay (Client-side)", marker_color="#d97706")
                fig_delay.add_bar(x=delay_df["MonthName"], y=delay_df["IBM Delay"],
                                  name="IBM Delay (Vendor-side)", marker_color="#d62828")
                fig_delay.update_layout(
                    barmode="group",
                    height=300, margin=dict(l=8,r=8,t=20,b=8),
                    paper_bgcolor="white", plot_bgcolor="#f5f9fd",
                    font=dict(family="IBM Plex Sans", size=11, color="#3a4a5c"),
                    xaxis=dict(tickangle=-35, gridcolor="#dde8f0"),
                    yaxis=dict(title="Delay Count", gridcolor="#dde8f0"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02,
                                xanchor="right", x=1, font=dict(size=10)),
                )
                st.plotly_chart(fig_delay, use_container_width=True,
                                config={"displayModeBar": False}, key="me_delay")

        with mc4:
            st.markdown('<div class="sec">Critical / High Sub-Category On-Time %</div>',
                        unsafe_allow_html=True)
            sub_df = me_f[me_f["On Time Sub %"].notna()].copy()
            if sub_df.empty:
                st.info("No sub-category data available in selected period.")
            else:
                sub_pivot = sub_df[["MonthName","Sub Category","On Time Sub","On Time Sub %"]].copy()
                color_map = {"Critical": "#d62828", "High": "#d97706"}
                fig_sub = go.Figure()
                for cat in sub_pivot["Sub Category"].dropna().unique():
                    cat_df = sub_pivot[sub_pivot["Sub Category"] == cat]
                    fig_sub.add_scatter(
                        x=cat_df["MonthName"], y=cat_df["On Time Sub %"],
                        name=f"{cat} On-Time %",
                        mode="lines+markers",
                        line=dict(color=color_map.get(cat, C_BLUE), width=2),
                        marker=dict(size=6))
                fig_sub.add_hline(y=100, line_dash="dash", line_color="#1a9e4e",
                                  line_width=1.5)
                # Dummy scatter so the target line appears in the legend
                fig_sub.add_scatter(
                    x=[None], y=[None], mode="lines",
                    name="── 100% Target",
                    line=dict(color="#1a9e4e", width=2, dash="dash"),
                    showlegend=True)
                fig_sub.update_layout(
                    height=300, margin=dict(l=8,r=8,t=20,b=8),
                    paper_bgcolor="white", plot_bgcolor="#f5f9fd",
                    font=dict(family="IBM Plex Sans", size=11, color="#3a4a5c"),
                    xaxis=dict(tickangle=-35, gridcolor="#dde8f0"),
                    yaxis=dict(title="On-Time %", range=[0, 115], gridcolor="#dde8f0"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02,
                                xanchor="right", x=1, font=dict(size=10)),
                )
                st.plotly_chart(fig_sub, use_container_width=True,
                                config={"displayModeBar": False}, key="me_sub_pct")

        # ── Monthly Summary Table (matches %On Time sheet format) ─────────
        st.markdown('<div class="sec">Monthly Summary Table — Last 12 Months</div>',
                    unsafe_allow_html=True)
        tbl = me_f.tail(12)[["MonthName","On Time","ABT Delay","IBM Delay",
                               "Total Changes","On Time All %","On Time Critical %"]].copy()
        tbl.columns = ["Month","On Time","ABT Delay","IBM Delay",
                       "Total","On Time All %","On Time Critical %"]

        def pct_badge(v):
            if pd.isna(v): return "—"
            cls = "bg-g" if v >= 100 else ("bg-y" if v >= 90 else "bg-r")
            return f'<span class="badge {cls}">{v:.0f}%</span>'

        rows_html = []
        for _, r in tbl.iterrows():
            abt_style = ' style="color:#d97706;font-weight:700"' if r["ABT Delay"] > 0 else ''
            ibm_style = ' style="color:#d62828;font-weight:700"' if r["IBM Delay"] > 0 else ''
            rows_html.append(f"""<tr>
              <td>{r['Month']}</td>
              <td>{int(r['On Time']) if pd.notna(r['On Time']) else '—'}</td>
              <td{abt_style}>{int(r['ABT Delay']) if pd.notna(r['ABT Delay']) else '—'}</td>
              <td{ibm_style}>{int(r['IBM Delay']) if pd.notna(r['IBM Delay']) else '—'}</td>
              <td><b>{int(r['Total']) if pd.notna(r['Total']) else '—'}</b></td>
              <td>{pct_badge(r['On Time All %'])}</td>
              <td>{pct_badge(r['On Time Critical %'])}</td>
            </tr>""")

        # Totals
        s = tbl.sum(numeric_only=True)
        tp_all  = round(tbl['On Time All %'].mean(), 1) if len(tbl) else 0
        tp_crit = round(tbl['On Time Critical %'].mean(), 1) if len(tbl) else 0
        rows_html.append(f"""<tr>
          <td><b>TOTAL / AVG</b></td>
          <td><b>{int(s['On Time'])}</b></td>
          <td><b>{int(s['ABT Delay'])}</b></td>
          <td><b>{int(s['IBM Delay'])}</b></td>
          <td><b>{int(s['Total'])}</b></td>
          <td>{pct_badge(tp_all)}</td>
          <td>{pct_badge(tp_crit)}</td>
        </tr>""")

        st.markdown(f"""
        <table class="tp-table">
          <thead><tr>
            <th>Month</th><th>On Time</th><th>ABT Delay</th>
            <th>IBM Delay</th><th>Total</th>
            <th>On Time All %</th><th>On Time Critical %</th>
          </tr></thead>
          <tbody>{"".join(rows_html)}</tbody>
        </table>""", unsafe_allow_html=True)

        # ── Rolling 12-month avg trend ─────────────────────────────────────
        st.markdown('<div class="sec">12-Month Rolling Average — Change Volume</div>',
                    unsafe_allow_html=True)
        roll = me.copy()
        roll["Rolling Avg"] = roll["Total Changes"].rolling(12, min_periods=3).mean().round(1)
        roll["Rolling Max"] = roll["Total Changes"].rolling(12, min_periods=3).max()
        roll["Rolling Min"] = roll["Total Changes"].rolling(12, min_periods=3).min()
        fig_roll = go.Figure()
        fig_roll.add_scatter(x=roll["MonthName"], y=roll["Rolling Max"],
                             fill=None, mode="lines",
                             line=dict(color="rgba(0,120,182,0.1)", width=0),
                             showlegend=False)
        fig_roll.add_scatter(x=roll["MonthName"], y=roll["Rolling Min"],
                             fill="tonexty", mode="lines",
                             line=dict(color="rgba(0,120,182,0.1)", width=0),
                             fillcolor="rgba(0,120,182,0.08)",
                             name="12-mo Min/Max Range", showlegend=True)
        fig_roll.add_bar(x=roll["MonthName"], y=roll["Total Changes"],
                         name="Monthly Total Changes", marker_color=C_TEAL, opacity=0.5)
        fig_roll.add_scatter(x=roll["MonthName"], y=roll["Rolling Avg"],
                             name="12-mo Rolling Average", mode="lines",
                             line=dict(color=C_NAVY, width=2.5))
        fig_roll.update_layout(
            height=300, margin=dict(l=8,r=8,t=20,b=8),
            paper_bgcolor="white", plot_bgcolor="#f5f9fd",
            font=dict(family="IBM Plex Sans", size=11, color="#3a4a5c"),
            xaxis=dict(tickangle=-40, gridcolor="#dde8f0",
                       tickmode="array",
                       tickvals=roll["MonthName"].iloc[::6].tolist()),
            yaxis=dict(title="Changes", gridcolor="#dde8f0"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02,
                        xanchor="right", x=1, font=dict(size=10)),
            barmode="overlay",
        )
        st.plotly_chart(fig_roll, use_container_width=True,
                        config={"displayModeBar": False}, key="me_rolling")

        # ── Raw data expander ─────────────────────────────────────────────
        with st.expander("🔍 View Raw ME Monthly Data"):
            st.dataframe(
                me_f[["MonthName","On Time","ABT Delay","IBM Delay",
                       "Total Changes","On Time All %","On Time Critical %",
                       "Sub Category","On Time Sub","On Time Sub %","Whitemark"]]
                .rename(columns={"MonthName":"Month"}),
                use_container_width=True, height=380)

        # ── Download ME data ──────────────────────────────────────────────
        me_buf = io.BytesIO()
        me_f.to_excel(me_buf, index=False, sheet_name="ME Monthly")
        me_buf.seek(0)
        st.download_button(
            "⬇️ Download ME Monthly Data (.xlsx)",
            data=me_buf.getvalue(),
            file_name=f"ME_Monthly_Report_{TODAY.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="me_dl"
        )


with tab_export:
    st.markdown('<div class="sec">Export Full AMS Report to Excel</div>',
                unsafe_allow_html=True)

    c1, c2 = st.columns([2,1])
    with c1:
        st.markdown("""
| Sheet | Contents |
|---|---|
| **Weekly Throughput** | Week-by-week created / resolved / net / throughput % |
| **Open Ageing** | Open backlog pivot — App × Age bucket |
| **Open Incidents** | Full open list with age, SLA status, priority |
| **SLA Breached** | Tickets that breached priority-based SLA targets |
| **Tower Summary** | Per-tower KPIs: volume, open, SLA breached, avg resolution, closure rate |
| **App Summary** | Per-app KPIs: volume, open, SLA breached, avg resolution |
| **All Created** | Full raw created dataset with computed fields |
| **All Resolved** | Full raw resolved dataset |
        """)

    with c2:
        _, tp_df_export = fig_weekly_throughput(cr, re)
        excel_bytes = build_export(cr, re, tp_df_export)
        st.download_button(
            label="⬇️  Download AMS Report (.xlsx)",
            data=excel_bytes,
            file_name=f"AMS_Report_{TODAY.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.markdown("---")
        st.metric("Created Records",  f"{len(cr):,}")
        st.metric("Resolved Records", f"{len(re):,}")
        st.metric("Open Incidents",   f"{total_op:,}")
        st.metric("SLA Breached",     f"{sla_br:,}")
