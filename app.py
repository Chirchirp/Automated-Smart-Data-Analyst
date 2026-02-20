"""
SokoData Solutions — Smart Data Explorer v2
============================================
A fully self-contained, API-free intelligent data analysis app.

Features:
  • CSV / Excel upload with auto-encoding detection
  • Data quality report  (missing values, duplicates, outliers)
  • Descriptive & distributional statistics
  • Auto-generated smart charts (bar, pie, line, histogram, scatter, heatmap)
  • Time-series detection & trend line
  • Correlation matrix with annotations
  • One-click downloadable HTML report
  • Auto-written plain-English insights — no API required

Run locally:
  pip install -r requirements.txt
  streamlit run app.py

Deploy to Streamlit Cloud:
  Push this folder to GitHub → share.streamlit.io → New app → select repo/branch/app.py
"""

import io
import base64
import warnings
import datetime
import itertools

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────
# 0.  Page config & brand colours
# ─────────────────────────────────────────
BRAND_BLUE   = "#1A73E8"
BRAND_GREEN  = "#34A853"
BRAND_AMBER  = "#FBBC05"
BRAND_DARK   = "#0F172A"
BRAND_LIGHT  = "#F8FAFC"
DANGER_RED   = "#EA4335"

st.set_page_config(
    page_title  = "SokoData — Smart Data Explorer",
    page_icon   = "📊",
    layout      = "wide",
    initial_sidebar_state = "expanded",
)

# ─────────────────────────────────────────
# Custom CSS — SokoData brand
# ─────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500&display=swap');

  html, body, [class*="css"] {{ font-family: 'DM Sans', sans-serif; }}
  h1,h2,h3 {{ font-family: 'Syne', sans-serif !important; }}

  /* Top header bar */
  .soko-header {{
    background: linear-gradient(135deg, {BRAND_DARK} 0%, #1e3a5f 100%);
    padding: 1.5rem 2rem;
    border-radius: 1rem;
    margin-bottom: 1.5rem;
    display: flex; align-items: center; gap: 1.2rem;
  }}
  .soko-logo {{ font-family:'Syne',sans-serif; font-size:1.6rem; font-weight:800; color:#fff; }}
  .soko-tagline {{ color:rgba(255,255,255,.6); font-size:.9rem; margin-top:.15rem; }}

  /* KPI metric cards */
  .kpi-row {{ display:flex; gap:.65rem; flex-wrap:wrap; margin-bottom:1.5rem; }}
  .kpi-card {{
    flex:1; min-width:100px;
    background:#fff; border-radius:.875rem;
    border:1px solid #e2e8f0;
    padding:.75rem .9rem;
    box-shadow:0 2px 8px rgba(0,0,0,.05);
  }}
  .kpi-label {{ font-size:.62rem; font-weight:700; letter-spacing:.07em; text-transform:uppercase; color:#94a3b8; white-space:nowrap; }}
  .kpi-value {{ font-family:'Syne',sans-serif; font-size:1.45rem; font-weight:800; margin:.15rem 0 0; line-height:1.1; }}
  .kpi-blue   {{ color:{BRAND_BLUE}; }}
  .kpi-green  {{ color:{BRAND_GREEN}; }}
  .kpi-amber  {{ color:{BRAND_AMBER}; }}
  .kpi-red    {{ color:{DANGER_RED}; }}

  /* Section pills */
  .section-pill {{
    display:inline-block;
    background:rgba(26,115,232,.1); color:{BRAND_BLUE};
    font-size:.72rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase;
    padding:.28rem .85rem; border-radius:2rem; margin-bottom:.6rem;
  }}

  /* Insight cards */
  .insight-card {{
    background: linear-gradient(135deg,rgba(26,115,232,.06),rgba(52,168,83,.04));
    border-left:4px solid {BRAND_BLUE};
    border-radius:0 .75rem .75rem 0;
    padding:.9rem 1.2rem; margin:.5rem 0;
    font-size:.92rem; color:#334155;
  }}
  .insight-warn {{
    border-left-color:{BRAND_AMBER};
    background:linear-gradient(135deg,rgba(251,188,5,.07),rgba(251,188,5,.02));
  }}
  .insight-danger {{
    border-left-color:{DANGER_RED};
    background:linear-gradient(135deg,rgba(234,67,53,.07),rgba(234,67,53,.02));
  }}
  .insight-success {{
    border-left-color:{BRAND_GREEN};
    background:linear-gradient(135deg,rgba(52,168,83,.07),rgba(52,168,83,.02));
  }}

  /* Data quality bar */
  .quality-bar-wrap {{ background:#e2e8f0; border-radius:2rem; height:.65rem; width:100%; }}
  .quality-bar-fill {{ height:.65rem; border-radius:2rem; background:linear-gradient(90deg,{BRAND_GREEN},{BRAND_BLUE}); }}

  /* Download button override */
  .stDownloadButton > button {{
    background:{BRAND_BLUE} !important; color:#fff !important;
    border-radius:.6rem !important; font-weight:600 !important;
    border:none !important;
  }}

  /* Sidebar */
  section[data-testid="stSidebar"] > div:first-child {{
    background:{BRAND_DARK};
  }}
  section[data-testid="stSidebar"] * {{ color:#e2e8f0 !important; }}
  section[data-testid="stSidebar"] .stFileUploader label {{ color:#93c5fd !important; }}

  /* Expander headers */
  .streamlit-expanderHeader {{ font-family:'Syne',sans-serif !important; font-size:1rem !important; font-weight:700 !important; }}
  
  /* Hide Streamlit branding */
  #MainMenu, footer {{ visibility:hidden; }}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# 1.  Utility functions
# ─────────────────────────────────────────

@st.cache_data(show_spinner=False)
def read_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    name = file_name.lower()
    buf  = io.BytesIO(file_bytes)
    if name.endswith((".xls", ".xlsx")):
        return pd.read_excel(buf)
    for enc in ("utf-8", "latin1", "cp1252"):
        try:
            buf.seek(0)
            return pd.read_csv(buf, encoding=enc)
        except Exception:
            continue
    buf.seek(0)
    return pd.read_csv(buf, sep=None, engine="python")


def numeric_cols(df: pd.DataFrame):
    return df.select_dtypes(include="number").columns.tolist()

def cat_cols(df: pd.DataFrame):
    obj = df.select_dtypes(include=["object","category"]).columns.tolist()
    low_card_num = [c for c in numeric_cols(df) if 2 < df[c].nunique() <= 15]
    return list(dict.fromkeys(obj + low_card_num))

def bool_cols(df: pd.DataFrame):
    return df.select_dtypes(include="bool").columns.tolist()

def detect_datetime(df: pd.DataFrame):
    dt = df.select_dtypes(include="datetime").columns.tolist()
    for col in df.select_dtypes(include="object").columns:
        sample = df[col].dropna().astype(str).head(30)
        if sample.empty: continue
        try:
            parsed = pd.to_datetime(sample, infer_datetime_format=True, errors="coerce")
            if parsed.notna().mean() > 0.7:
                dt.append(col)
        except Exception:
            pass
    return list(dict.fromkeys(dt))

def outlier_count(series: pd.Series) -> int:
    q1, q3 = series.quantile(0.25), series.quantile(0.75)
    iqr = q3 - q1
    return int(((series < q1 - 1.5*iqr) | (series > q3 + 1.5*iqr)).sum())

def memory_mb(df: pd.DataFrame) -> float:
    return round(df.memory_usage(deep=True).sum() / 1e6, 3)

def completeness_score(df: pd.DataFrame) -> float:
    total = df.shape[0] * df.shape[1]
    if total == 0: return 100.0
    return round((1 - df.isna().sum().sum() / total) * 100, 1)

def missing_report(df: pd.DataFrame) -> pd.DataFrame:
    miss   = df.isna().sum()
    pct    = (miss / len(df) * 100).round(2)
    dtype  = df.dtypes.astype(str)
    report = pd.DataFrame({"Missing #": miss, "Missing %": pct, "dtype": dtype})
    return report[report["Missing #"] > 0].sort_values("Missing #", ascending=False)

def auto_insights(df: pd.DataFrame) -> list[dict]:
    """
    Generate plain-English insights with no external API.
    Returns list of dicts: {text, level}  level ∈ {info, warn, danger, success}
    """
    insights = []
    nc = numeric_cols(df)
    cc = cat_cols(df)
    dt = detect_datetime(df)

    # ── Size
    insights.append({"text": f"Dataset contains **{df.shape[0]:,} rows** and **{df.shape[1]} columns**, occupying **{memory_mb(df)} MB** in memory.", "level": "info"})

    # ── Completeness
    comp = completeness_score(df)
    if comp == 100:
        insights.append({"text": "✅ Dataset is **100% complete** — no missing values found.", "level": "success"})
    elif comp >= 90:
        insights.append({"text": f"⚠️ Completeness is **{comp}%** — minor missing data. Consider imputation for modelling.", "level": "warn"})
    else:
        insights.append({"text": f"🔴 Completeness is only **{comp}%** — significant missing data. Data cleaning strongly advised before analysis.", "level": "danger"})

    # ── Duplicates
    dups = int(df.duplicated().sum())
    if dups == 0:
        insights.append({"text": "✅ No duplicate rows detected.", "level": "success"})
    else:
        pct = round(dups / len(df) * 100, 1)
        insights.append({"text": f"⚠️ **{dups} duplicate rows** ({pct}% of dataset). Remove duplicates before statistical analysis.", "level": "warn"})

    # ── Numeric insights
    for col in nc[:8]:
        s = df[col].dropna()
        if s.empty: continue
        skew = s.skew()
        cv   = (s.std() / s.mean() * 100) if s.mean() != 0 else 0
        out  = outlier_count(s)
        if abs(skew) > 1.5:
            direction = "right (positively)" if skew > 0 else "left (negatively)"
            insights.append({"text": f"Column **{col}** is {direction} skewed (skew={skew:.2f}). A log or Box-Cox transform may normalise it.", "level": "warn"})
        if out > 0:
            insights.append({"text": f"Column **{col}** has **{out} outlier(s)** (IQR method). Verify if these are data errors or genuine extremes.", "level": "warn"})
        if abs(cv) > 100:
            insights.append({"text": f"Column **{col}** has very high variability (CV={cv:.0f}%). Values span a wide range — consider segmenting the analysis.", "level": "info"})

    # ── Correlation
    if len(nc) >= 2:
        corr = df[nc].corr()
        pairs = [(corr.columns[i], corr.columns[j], corr.iloc[i,j])
                 for i in range(len(corr)) for j in range(i+1,len(corr))]
        strong = [(a,b,r) for a,b,r in pairs if abs(r) >= 0.8]
        for a,b,r in strong[:3]:
            label = "strong positive" if r > 0 else "strong negative"
            insights.append({"text": f"**{a}** and **{b}** have a {label} correlation (r={r:.2f}). These columns may be redundant in predictive models.", "level": "info"})

    # ── Categorical
    for col in cc[:4]:
        vc = df[col].value_counts(dropna=False)
        dom_pct = vc.iloc[0] / len(df) * 100
        if dom_pct > 80:
            insights.append({"text": f"Column **{col}** is dominated by **'{vc.index[0]}'** ({dom_pct:.0f}% of records). Low variance may limit its predictive utility.", "level": "warn"})

    # ── Time series
    if dt:
        insights.append({"text": f"✅ Datetime column(s) detected: **{', '.join(dt)}**. Scroll to the Time-Series section for trend analysis.", "level": "success"})

    return insights


def chart_palette(n: int):
    palette = [BRAND_BLUE, BRAND_GREEN, BRAND_AMBER, "#8B5CF6", "#EC4899", "#14B8A6", "#F97316", DANGER_RED]
    return list(itertools.islice(itertools.cycle(palette), n))


# ─────────────────────────────────────────
# 2.  Sidebar — upload & options
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:1.2rem 0 .5rem;">
      <svg width="40" height="40" viewBox="0 0 32 32" fill="none">
        <path d="M16 2L30 16L16 30L2 16L16 2Z" stroke="#93c5fd" stroke-width="2"/>
        <path d="M16 8L24 16L16 24L8 16L16 8Z" stroke="#60a5fa" stroke-width="2"/>
      </svg>
      <div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.2rem;color:#fff;margin-top:.4rem;">SokoData</div>
      <div style="color:#93c5fd;font-size:.8rem;">Smart Data Explorer</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    uploaded = st.file_uploader(
        "📂 Upload CSV or Excel",
        type=["csv","xls","xlsx"],
        help="Max 200 MB. Supports CSV (UTF-8/Latin-1) and Excel (.xls/.xlsx)"
    )

    st.markdown("---")
    st.markdown("**⚙️ Display options**")
    max_cat_bars  = st.slider("Max categories in bar/pie charts", 5, 20, 10)
    max_hist_cols = st.slider("Max numeric histograms", 1, 12, 6)
    show_raw      = st.checkbox("Show raw data table", value=True)
    show_types    = st.checkbox("Show column types table", value=True)

    st.markdown("---")
    st.markdown("**🔗 Try demo data**")
    use_demo = st.button("Load demo dataset")

    st.markdown("---")
    st.caption("No data leaves your browser. 100% local analysis — no API key required.")


# ─────────────────────────────────────────
# 3.  Load data
# ─────────────────────────────────────────
DEMO_CSV = """customer_id,age,monthly_spend_ksh,tenure_months,data_usage_gb,calls_per_week,sms_per_week,county,segment,churn_risk
C001,34,2400,18,12.5,45,30,Nairobi,SME,low
C002,27,850,3,3.2,12,8,Mombasa,Consumer,high
C003,45,5200,36,28.1,67,12,Nairobi,Enterprise,low
C004,22,400,1,1.1,5,20,Kisumu,Consumer,high
C005,38,3100,24,18.7,55,40,Nakuru,SME,medium
C006,52,7800,60,35.4,80,5,Nairobi,Enterprise,low
C007,29,1200,6,8.9,22,55,Eldoret,Consumer,medium
C008,41,2900,30,22.0,60,18,Mombasa,SME,low
C009,31,1800,12,9.8,34,44,Thika,Consumer,medium
C010,55,9500,72,48.2,90,3,Nairobi,Enterprise,low
C011,24,300,1,0.8,4,10,Kisumu,Consumer,high
C012,47,4200,42,25.3,70,8,Nakuru,SME,low
C013,36,2100,15,11.2,40,25,Nairobi,SME,medium
C014,28,700,4,4.5,15,35,Mombasa,Consumer,high
C015,43,6300,48,32.8,75,6,Nairobi,Enterprise,low
"""

if use_demo:
    df = pd.read_csv(io.StringIO(DEMO_CSV))
    st.session_state["df"]       = df
    st.session_state["filename"] = "demo_dataset.csv"

if uploaded is not None:
    with st.spinner("Reading file …"):
        df = read_file(uploaded.read(), uploaded.name)
    st.session_state["df"]       = df
    st.session_state["filename"] = uploaded.name

if "df" not in st.session_state:
    # ── Landing screen
    st.markdown("""
    <div class="soko-header">
      <svg width="44" height="44" viewBox="0 0 32 32" fill="none">
        <path d="M16 2L30 16L16 30L2 16L16 2Z" stroke="white" stroke-width="2"/>
        <path d="M16 8L24 16L16 24L8 16L16 8Z" stroke="rgba(255,255,255,.6)" stroke-width="2"/>
      </svg>
      <div>
        <div class="soko-logo">SokoData — Smart Data Explorer</div>
        <div class="soko-tagline">Upload a CSV or Excel file to get instant, intelligent analytics — no API key needed.</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    c1,c2,c3 = st.columns(3)
    for col, icon, title, desc in [
        (c1,"📊","Auto Charts","Bar, pie, line, histogram & scatter plots generated intelligently from your data."),
        (c2,"🔍","Data Quality","Missing values, duplicates, outlier detection & completeness score."),
        (c3,"💡","Smart Insights","Plain-English findings written automatically — no AI API required."),
    ]:
        col.markdown(f"""
        <div style="background:#fff;border-radius:1rem;border:1px solid #e2e8f0;padding:1.5rem;text-align:center;">
          <div style="font-size:2rem;">{icon}</div>
          <h3 style="font-family:Syne,sans-serif;margin:.5rem 0 .3rem;">{title}</h3>
          <p style="color:#64748b;font-size:.88rem;">{desc}</p>
        </div>""", unsafe_allow_html=True)
    st.stop()

df: pd.DataFrame = st.session_state["df"]
filename: str    = st.session_state.get("filename","dataset")

# Column type lists
NC = numeric_cols(df)
CC = cat_cols(df)
DT = detect_datetime(df)

# ─────────────────────────────────────────
# 4.  Page header
# ─────────────────────────────────────────
st.markdown(f"""
<div class="soko-header">
  <svg width="40" height="40" viewBox="0 0 32 32" fill="none">
    <path d="M16 2L30 16L16 30L2 16L16 2Z" stroke="white" stroke-width="2"/>
    <path d="M16 8L24 16L16 24L8 16L16 8Z" stroke="rgba(255,255,255,.6)" stroke-width="2"/>
  </svg>
  <div>
    <div class="soko-logo">Smart Data Explorer</div>
    <div class="soko-tagline">Analysing: <strong style="color:#93c5fd;">{filename}</strong> &nbsp;|&nbsp; {df.shape[0]:,} rows · {df.shape[1]} columns</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# 5.  KPI metric row
# ─────────────────────────────────────────
comp   = completeness_score(df)
dups   = int(df.duplicated().sum())
miss   = int(df.isna().sum().sum())
total_out = sum(outlier_count(df[c].dropna()) for c in NC)
comp_color = "kpi-green" if comp == 100 else ("kpi-amber" if comp >= 85 else "kpi-red")
dup_color  = "kpi-green" if dups == 0 else "kpi-amber"

st.markdown(f"""
<div class="kpi-row">
  <div class="kpi-card"><div class="kpi-label">Rows</div><div class="kpi-value kpi-blue">{df.shape[0]:,}</div></div>
  <div class="kpi-card"><div class="kpi-label">Columns</div><div class="kpi-value kpi-blue">{df.shape[1]}</div></div>
  <div class="kpi-card"><div class="kpi-label">Completeness</div><div class="kpi-value {comp_color}">{comp}%</div></div>
  <div class="kpi-card"><div class="kpi-label">Missing Cells</div><div class="kpi-value kpi-amber">{miss:,}</div></div>
  <div class="kpi-card"><div class="kpi-label">Duplicate Rows</div><div class="kpi-value {dup_color}">{dups:,}</div></div>
  <div class="kpi-card"><div class="kpi-label">Outliers (IQR)</div><div class="kpi-value kpi-amber">{total_out:,}</div></div>
  <div class="kpi-card"><div class="kpi-label">Numeric Cols</div><div class="kpi-value kpi-blue">{len(NC)}</div></div>
  <div class="kpi-card"><div class="kpi-label">Categorical Cols</div><div class="kpi-value kpi-green">{len(CC)}</div></div>
</div>
""", unsafe_allow_html=True)

# Completeness bar
st.markdown(f"""
<div style="margin-bottom:1.5rem;">
  <div style="font-size:.8rem;color:#64748b;margin-bottom:.3rem;">Data completeness — {comp}%</div>
  <div class="quality-bar-wrap"><div class="quality-bar-fill" style="width:{comp}%;"></div></div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# 6.  Auto Insights
# ─────────────────────────────────────────
with st.expander("💡 Automatic Insights", expanded=True):
    st.markdown('<span class="section-pill">AI-free intelligence</span>', unsafe_allow_html=True)
    insights = auto_insights(df)
    level_map = {"info":"insight-card","warn":"insight-card insight-warn","danger":"insight-card insight-danger","success":"insight-card insight-success"}
    for ins in insights:
        st.markdown(f'<div class="{level_map[ins["level"]]}">{ins["text"]}</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────
# 7.  Data Preview
# ─────────────────────────────────────────
if show_raw:
    with st.expander("🗂 Data Preview & Column Types", expanded=True):
        st.markdown('<span class="section-pill">Raw data</span>', unsafe_allow_html=True)

        tab_prev, tab_tail, tab_types = st.tabs(["First 10 rows","Last 10 rows","Column types"])
        with tab_prev:
            st.dataframe(df.head(10), use_container_width=True)
        with tab_tail:
            st.dataframe(df.tail(10), use_container_width=True)
        if show_types:
            with tab_types:
                tdf = pd.DataFrame({
                    "dtype":  df.dtypes.astype(str),
                    "unique": df.nunique(dropna=True),
                    "missing %": (df.isna().mean()*100).round(2),
                    "sample": [str(df[c].dropna().iloc[0]) if df[c].dropna().size else "—" for c in df.columns]
                })
                st.dataframe(tdf, use_container_width=True)


# ─────────────────────────────────────────
# 8.  Data Quality Detail
# ─────────────────────────────────────────
with st.expander("🔍 Data Quality Report", expanded=False):
    st.markdown('<span class="section-pill">Quality checks</span>', unsafe_allow_html=True)

    q1, q2 = st.columns(2)

    with q1:
        st.subheader("Missing Values")
        mr = missing_report(df)
        if mr.empty:
            st.success("No missing values! 🎉")
        else:
            st.dataframe(mr, use_container_width=True)
            fig = px.bar(
                mr.reset_index().rename(columns={"index":"Column"}),
                x="Column", y="Missing %",
                color="Missing %",
                color_continuous_scale=["#34A853","#FBBC05","#EA4335"],
                title="Missing % by Column",
            )
            fig.update_layout(showlegend=False, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    with q2:
        st.subheader("Duplicate Rows")
        if dups == 0:
            st.success("No duplicates detected! 🎉")
        else:
            st.warning(f"{dups} duplicate rows ({round(dups/len(df)*100,1)}%)")
            st.dataframe(df[df.duplicated(keep="first")].head(10), use_container_width=True)

    # Outlier report
    if NC:
        st.subheader("Outlier Summary (IQR method)")
        out_data = []
        for col in NC:
            s  = df[col].dropna()
            q1v, q3v = s.quantile(.25), s.quantile(.75)
            iqr = q3v - q1v
            n   = outlier_count(s)
            out_data.append({"Column": col, "Outliers": n, "Outlier %": round(n/len(s)*100,1),
                              "Lower fence": round(q1v-1.5*iqr,3), "Upper fence": round(q3v+1.5*iqr,3)})
        st.dataframe(pd.DataFrame(out_data).sort_values("Outliers",ascending=False), use_container_width=True)


# ─────────────────────────────────────────
# 9.  Descriptive Statistics
# ─────────────────────────────────────────
with st.expander("📐 Descriptive Statistics", expanded=True):
    st.markdown('<span class="section-pill">Summary stats</span>', unsafe_allow_html=True)

    tab_num, tab_cat = st.tabs(["Numeric summary","Categorical summary"])

    with tab_num:
        if not NC:
            st.info("No numeric columns.")
        else:
            num_stats = df[NC].describe(percentiles=[.1,.25,.5,.75,.9]).T
            num_stats["skew"]     = df[NC].skew().round(3)
            num_stats["kurtosis"] = df[NC].kurtosis().round(3)
            num_stats["outliers"] = [outlier_count(df[c].dropna()) for c in NC]
            # background_gradient requires matplotlib — not in requirements.
            # Use a Plotly table instead (zero extra dependencies).
            fig_tbl = go.Figure(data=[go.Table(
                header=dict(
                    values=["<b>Column</b>"] + [f"<b>{c}</b>" for c in num_stats.columns],
                    fill_color="#0F172A", font=dict(color="white", size=12),
                    align="left", line_color="#1e293b", height=32,
                ),
                cells=dict(
                    values=[num_stats.index.tolist()] +
                           [num_stats[c].round(3).tolist() for c in num_stats.columns],
                    fill_color=[["#f8fafc" if i % 2 == 0 else "#eff6ff"
                                 for i in range(len(num_stats))]],
                    font=dict(size=11, color="#0F172A"),
                    align="left", line_color="#e2e8f0", height=28,
                )
            )])
            fig_tbl.update_layout(
                margin=dict(l=0, r=0, t=0, b=0),
                paper_bgcolor="rgba(0,0,0,0)",
                height=max(200, 32 + len(num_stats) * 30),
            )
            st.plotly_chart(fig_tbl, use_container_width=True)

    with tab_cat:
        cat_only = df.select_dtypes(include=["object","category"])
        if cat_only.empty:
            st.info("No categorical columns.")
        else:
            st.dataframe(cat_only.describe().T, use_container_width=True)


# ─────────────────────────────────────────
# 10.  Numeric Distributions
# ─────────────────────────────────────────
with st.expander("📊 Numeric Distributions", expanded=True):
    st.markdown('<span class="section-pill">Histograms & box plots</span>', unsafe_allow_html=True)

    if not NC:
        st.info("No numeric columns to visualise.")
    else:
        cols_to_show = NC[:max_hist_cols]
        for i in range(0, len(cols_to_show), 2):
            row_cols = cols_to_show[i:i+2]
            grid = st.columns(len(row_cols))
            for gc, col in zip(grid, row_cols):
                with gc:
                    fig = px.histogram(df, x=col, nbins=30, marginal="box",
                                       color_discrete_sequence=[BRAND_BLUE],
                                       title=f"{col}")
                    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                      bargap=.05, showlegend=False, height=340)
                    st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────
# 11.  Categorical Breakdowns
# ─────────────────────────────────────────
with st.expander("🥧 Categorical Breakdowns", expanded=True):
    st.markdown('<span class="section-pill">Bar & pie charts</span>', unsafe_allow_html=True)

    if not CC:
        st.info("No categorical columns detected.")
    else:
        for col in CC[:6]:
            vc = df[col].value_counts(dropna=False).head(max_cat_bars)
            labels = vc.index.astype(str).tolist()
            values = vc.values.tolist()
            colours = chart_palette(len(labels))

            c_bar, c_pie = st.columns(2)
            with c_bar:
                fig = px.bar(x=labels, y=values, color=labels,
                             color_discrete_sequence=colours,
                             title=f"{col} — top {max_cat_bars}",
                             labels={"x": col,"y":"Count"})
                fig.update_layout(showlegend=False, plot_bgcolor="rgba(0,0,0,0)",
                                  paper_bgcolor="rgba(0,0,0,0)", height=320)
                st.plotly_chart(fig, use_container_width=True)
            with c_pie:
                fig = px.pie(names=labels, values=values,
                             color_discrete_sequence=colours,
                             title=f"{col} — distribution",
                             hole=.38)
                fig.update_traces(textposition="inside", textinfo="percent+label")
                fig.update_layout(showlegend=False, height=320,
                                  paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────
# 12.  Correlation Matrix
# ─────────────────────────────────────────
with st.expander("🔗 Correlation Matrix", expanded=False):
    st.markdown('<span class="section-pill">Numeric correlations</span>', unsafe_allow_html=True)

    if len(NC) < 2:
        st.info("Need at least 2 numeric columns for correlation.")
    else:
        corr = df[NC].corr().round(3)
        fig = px.imshow(
            corr, text_auto=True, aspect="auto",
            color_continuous_scale="RdBu_r",
            zmin=-1, zmax=1,
            title="Pearson Correlation Matrix"
        )
        fig.update_layout(height=max(350, len(NC)*50), paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

        # Strongest pairs table
        pairs = []
        for i in range(len(corr)):
            for j in range(i+1, len(corr)):
                r = corr.iloc[i,j]
                pairs.append({"Column A": corr.columns[i], "Column B": corr.columns[j], "Correlation": round(r,3), "|r|": abs(round(r,3))})
        pairs_df = pd.DataFrame(pairs).sort_values("|r|", ascending=False).drop(columns="|r|")
        st.markdown("**Strongest correlations (top 10)**")
        st.dataframe(pairs_df.head(10), use_container_width=True)


# ─────────────────────────────────────────
# 13.  Scatter / Relationship Explorer
# ─────────────────────────────────────────
with st.expander("🔵 Relationship Explorer (Scatter)", expanded=False):
    st.markdown('<span class="section-pill">Interactive scatter</span>', unsafe_allow_html=True)

    if len(NC) < 2:
        st.info("Need at least 2 numeric columns.")
    else:
        sc1, sc2, sc3 = st.columns(3)
        x_col  = sc1.selectbox("X axis",  NC, index=0)
        y_col  = sc2.selectbox("Y axis",  NC, index=min(1,len(NC)-1))
        c_col  = sc3.selectbox("Colour by (optional)", ["— none —"] + CC, index=0)
        colour = None if c_col == "— none —" else c_col

        fig = px.scatter(df, x=x_col, y=y_col, color=colour,
                         trendline="ols" if colour is None else None,
                         color_discrete_sequence=chart_palette(df[colour].nunique() if colour else 1),
                         title=f"{y_col} vs {x_col}",
                         opacity=.75)
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=420)
        st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────
# 14.  Time-Series Analysis
# ─────────────────────────────────────────
with st.expander("📈 Time-Series Analysis", expanded=False):
    st.markdown('<span class="section-pill">Trends over time</span>', unsafe_allow_html=True)

    if not DT:
        st.info("No datetime column detected. Tip: ensure your date column is formatted as YYYY-MM-DD.")
    else:
        ts1, ts2, ts3 = st.columns(3)
        date_col  = ts1.selectbox("Date column", DT)
        value_col = ts2.selectbox("Value column", NC) if NC else None
        agg_fn    = ts3.selectbox("Aggregate by", ["Day","Week","Month","Quarter","Year"])

        if value_col:
            df_ts = df[[date_col, value_col]].copy()
            df_ts[date_col] = pd.to_datetime(df_ts[date_col], errors="coerce")
            df_ts = df_ts.dropna().sort_values(date_col)

            freq_map = {"Day":"D","Week":"W","Month":"ME","Quarter":"QE","Year":"YE"}
            df_agg = df_ts.resample(freq_map[agg_fn], on=date_col)[value_col].mean().reset_index()

            fig = px.line(df_agg, x=date_col, y=value_col,
                          title=f"{value_col} — {agg_fn}ly average",
                          color_discrete_sequence=[BRAND_BLUE],
                          markers=True)
            fig.update_traces(line_width=2.5)
            # Add rolling average
            if len(df_agg) >= 4:
                df_agg["rolling"] = df_agg[value_col].rolling(3, min_periods=1).mean()
                fig.add_scatter(x=df_agg[date_col], y=df_agg["rolling"],
                                name="3-period avg", line=dict(color=BRAND_AMBER, dash="dot", width=2))
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=380)
            st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────
# 15.  Custom Chart Builder
# ─────────────────────────────────────────
with st.expander("🎨 Custom Chart Builder", expanded=False):
    st.markdown('<span class="section-pill">Build your own chart</span>', unsafe_allow_html=True)

    chart_type = st.selectbox("Chart type", ["Bar","Horizontal Bar","Line","Scatter","Box","Violin","Pie/Donut"])
    all_cols   = df.columns.tolist()

    cc1, cc2, cc3 = st.columns(3)
    cx = cc1.selectbox("X / Category", all_cols)
    cy = cc2.selectbox("Y / Value",    all_cols, index=min(1,len(all_cols)-1))
    cg = cc3.selectbox("Group / Colour (optional)", ["— none —"] + all_cols)
    cg_val = None if cg == "— none —" else cg

    if st.button("Generate Chart ▶"):
        try:
            # Build common kwargs WITHOUT df — pass df as first positional arg below
            shared = dict(
                color=cg_val,
                color_discrete_sequence=chart_palette(10),
            )
            if chart_type == "Bar":
                fig = px.bar(df, x=cx, y=cy, title=f"{cy} by {cx}", barmode="group", **shared)
            elif chart_type == "Horizontal Bar":
                fig = px.bar(df, x=cy, y=cx, title=f"{cy} by {cx}", orientation="h", **shared)
            elif chart_type == "Line":
                fig = px.line(df, x=cx, y=cy, title=f"{cy} over {cx}", markers=True, **shared)
            elif chart_type == "Scatter":
                fig = px.scatter(df, x=cx, y=cy, title=f"{cy} vs {cx}", opacity=.75,
                                 trendline="ols" if cg_val is None else None, **shared)
            elif chart_type == "Box":
                fig = px.box(df, x=cx, y=cy, title=f"{cy} by {cx}", **shared)
            elif chart_type == "Violin":
                fig = px.violin(df, x=cx, y=cy, title=f"{cy} by {cx}", box=True, **shared)
            else:  # Pie / Donut
                vc = df[cx].value_counts().head(max_cat_bars)
                fig = px.pie(
                    names=vc.index.astype(str), values=vc.values,
                    color_discrete_sequence=chart_palette(len(vc)),
                    hole=.38, title=f"{cx} — distribution",
                )
                fig.update_traces(textposition="inside", textinfo="percent+label")

            fig.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                height=450,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            )
            st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.error(f"Could not render chart: {e}")
            st.info("Tip: make sure your Y column contains numeric data for bar/line/scatter charts.")


# ─────────────────────────────────────────
# 16.  Downloads
# ─────────────────────────────────────────
with st.expander("⬇️ Download Reports", expanded=False):
    st.markdown('<span class="section-pill">Export</span>', unsafe_allow_html=True)

    d1, d2, d3 = st.columns(3)

    # Missing report CSV
    with d1:
        mr_csv = missing_report(df).to_csv().encode("utf-8")
        st.download_button("📥 Missing Values CSV", data=mr_csv,
                           file_name="missing_report.csv", mime="text/csv")

    # Numeric summary CSV
    with d2:
        ns = df[NC].describe(percentiles=[.1,.25,.5,.75,.9]).T if NC else pd.DataFrame()
        ns_csv = ns.to_csv().encode("utf-8")
        st.download_button("📥 Numeric Summary CSV", data=ns_csv,
                           file_name="numeric_summary.csv", mime="text/csv")

    # Cleaned CSV (drop duplicates)
    with d3:
        clean_csv = df.drop_duplicates().to_csv(index=False).encode("utf-8")
        st.download_button("📥 De-duplicated Dataset CSV", data=clean_csv,
                           file_name="cleaned_data.csv", mime="text/csv")

    # Full HTML report
    st.markdown("---")
    if st.button("📄 Generate Full HTML Report"):
        insights_html = "".join(
            f'<li style="margin:.4rem 0;color:{"#dc2626" if i["level"]=="danger" else "#b45309" if i["level"]=="warn" else "#16a34a" if i["level"]=="success" else "#1d4ed8"}">'
            f'{i["text"].replace("**","<strong>").replace("**","</strong>")}</li>'
            for i in auto_insights(df)
        )
        num_html = df[NC].describe().T.to_html(classes="table") if NC else "<p>No numeric columns.</p>"
        cat_html = df.select_dtypes(include=["object","category"]).describe().T.to_html(classes="table") if CC else "<p>No categorical columns.</p>"
        miss_html = missing_report(df).to_html(classes="table") if not missing_report(df).empty else "<p>No missing values.</p>"

        report = f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>SokoData Report — {filename}</title>
<style>
  body{{font-family:'Segoe UI',sans-serif;margin:0;padding:0;background:#f8fafc;color:#0f172a}}
  header{{background:linear-gradient(135deg,#0f172a,#1e3a5f);color:#fff;padding:2rem 3rem}}
  header h1{{margin:0;font-size:1.8rem}} header p{{color:rgba(255,255,255,.65);margin:.3rem 0 0}}
  .container{{max-width:1100px;margin:2rem auto;padding:0 2rem}}
  h2{{font-size:1.2rem;border-left:4px solid #1A73E8;padding-left:.75rem;margin-top:2.5rem}}
  .kpi-row{{display:flex;gap:1rem;flex-wrap:wrap;margin:1.5rem 0}}
  .kpi{{flex:1;min-width:120px;background:#fff;border-radius:.75rem;border:1px solid #e2e8f0;padding:1rem 1.25rem;text-align:center}}
  .kpi-label{{font-size:.7rem;text-transform:uppercase;letter-spacing:.08em;color:#94a3b8}}
  .kpi-value{{font-size:2rem;font-weight:800;color:#1A73E8;margin-top:.2rem}}
  .table{{width:100%;border-collapse:collapse;background:#fff;border-radius:.75rem;overflow:hidden;font-size:.85rem}}
  .table th{{background:#f1f5f9;padding:.55rem .85rem;text-align:left;font-size:.78rem;color:#475569}}
  .table td{{padding:.5rem .85rem;border-top:1px solid #f1f5f9;color:#334155}}
  ul.insights{{padding-left:1.2rem}} ul.insights li{{margin:.4rem 0}}
  footer{{text-align:center;color:#94a3b8;font-size:.8rem;padding:2rem;margin-top:3rem;border-top:1px solid #e2e8f0}}
</style></head><body>
<header>
  <h1>SokoData — Smart Data Explorer Report</h1>
  <p>File: {filename} &nbsp;|&nbsp; Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} &nbsp;|&nbsp; {df.shape[0]:,} rows · {df.shape[1]} columns</p>
</header>
<div class="container">
  <div class="kpi-row">
    <div class="kpi"><div class="kpi-label">Rows</div><div class="kpi-value">{df.shape[0]:,}</div></div>
    <div class="kpi"><div class="kpi-label">Columns</div><div class="kpi-value">{df.shape[1]}</div></div>
    <div class="kpi"><div class="kpi-label">Completeness</div><div class="kpi-value">{comp}%</div></div>
    <div class="kpi"><div class="kpi-label">Missing cells</div><div class="kpi-value">{miss:,}</div></div>
    <div class="kpi"><div class="kpi-label">Duplicates</div><div class="kpi-value">{dups:,}</div></div>
  </div>
  <h2>Automatic Insights</h2><ul class="insights">{insights_html}</ul>
  <h2>Missing Values</h2>{miss_html}
  <h2>Numeric Summary</h2>{num_html}
  <h2>Categorical Summary</h2>{cat_html}
  <footer>Generated by SokoData Solutions · sokodatasolutions.com</footer>
</div></body></html>"""

        st.download_button("📥 Download HTML Report", data=report.encode("utf-8"),
                           file_name=f"sokodata_report_{filename}.html", mime="text/html")

# ─────────────────────────────────────────
# 17.  Footer
# ─────────────────────────────────────────
st.markdown("---")
st.markdown(f"""
<div style="text-align:center;color:#94a3b8;font-size:.82rem;padding:.5rem 0 1.5rem;">
  Built by <strong style="color:{BRAND_BLUE};">SokoData Solutions</strong> ·
  100% local analysis · no API key required ·
  <a href="https://sokodatasolutions.com" style="color:{BRAND_BLUE};">sokodatasolutions.com</a>
</div>
""", unsafe_allow_html=True)
