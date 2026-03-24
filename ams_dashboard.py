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
                    line=dict(color="#d62828", width=2), name="Open Backlog")
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
    df["Color"] = df["Avg Days"].apply(
        lambda x: "#d62828" if x > 14 else ("#d97706" if x > 7 else "#1a9e4e"))
    fig = px.bar(df, x="App", y="Avg Days", color="Color",
                 color_discrete_map="identity",
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
                name="Daily Count", marker_color=C_BLUE, opacity=0.55)
    fig.add_scatter(x=daily["Date"], y=daily["MA7"],
                    name="7-day Avg", line=dict(color="#d62828", width=2.2), mode="lines")
    return chart_style(fig, "Daily Incident Volume + 7-day Rolling Average", height=300)


def fig_source_split(cr):
    src = cr.groupby(["Created Week Ending", "Source"]).size().reset_index(name="Count")
    src["WLabel"] = src["Created Week Ending"].dt.strftime("%-d %b")
    src = src.sort_values("Created Week Ending")
    fig = px.bar(src, x="WLabel", y="Count", color="Source",
                 color_discrete_map={"Automated": C_TEAL, "Manual": C_NAVY},
                 barmode="stack")
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
    fig.update_traces(textposition="outside")
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
        ["Number","App","Priority","Impact","Created On","Age Days",
         "Age Bucket","SLA Breached","Assignment Group","Opened By"]
        if c in open_df.columns]

    sla_cols = [c for c in
        ["Number","App","Priority","Created On","Age Days","Assignment Group"]
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
        cr.to_excel(w, sheet_name="All Created", index=False)
        re.to_excel(w, sheet_name="All Resolved", index=False)
    buf.seek(0)
    return buf.getvalue()


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
    st.markdown("---")
    st.caption("AMS Operations Dashboard v2.0")


# ─────────────────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────────────────
st.markdown(f"""
<div class="ams-header">
  <div>
    <h1>📊 AMS Operations Dashboard</h1>
    <p>Abbott Application Management Support · ServiceNow Incident Analytics</p>
  </div>
  <div class="ams-right">
    <span class="ams-badge">LIVE REPORT</span><br>
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

# ── Sidebar filters ──
apps   = ["All"] + sorted(cr_raw["App"].dropna().unique())
pris   = ["All"] + sorted(cr_raw["Priority"].dropna().unique(),
                           key=lambda x: {"High":0,"Moderate":1,"Low":2}.get(x,9))
groups = ["All"] + sorted(cr_raw["Assignment Group"].dropna().unique())
min_d  = cr_raw["Created On"].min().date()
max_d  = cr_raw["Created On"].max().date()

with st.sidebar:
    st.markdown("### ⚙️ Filters")
    sel_app   = st.selectbox("Application",      apps,   key="f_app")
    sel_pri   = st.selectbox("Priority",         pris,   key="f_pri")
    sel_grp   = st.selectbox("Assignment Group", groups, key="f_grp")
    date_rng  = st.date_input("Date Range",
                               value=(min_d, max_d),
                               min_value=min_d, max_value=max_d,
                               key="f_date")
    show_auto = st.checkbox("Include Auto-created tickets", value=True, key="f_auto")

# ── Apply filters ──
cr = cr_raw.copy()
re = re_raw.copy()

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
 tab_team, tab_trends, tab_drill, tab_export) = st.tabs([
    "🚨 Action Required",
    "📅 Weekly Throughput",
    "⏳ Ageing & Backlog",
    "👥 Team & App Performance",
    "📈 Trends & Analytics",
    "🔍 Incident Drill-down",
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
            fig_aged = go.Figure(go.Bar(
                x=age_cnt.index, y=age_cnt.values,
                marker_color=AGE_COLORS,
                text=age_cnt.values.astype(int),
                textposition="outside"))
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
        ["Number","App","Priority","Impact","Created On",
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
