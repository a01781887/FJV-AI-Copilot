import re
import html
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st


# ============================================================
# PAGE
# ============================================================

st.set_page_config(
    st.markdown("""
<div style="
    background:#173F46;
    padding:28px 32px;
    border-radius:22px;
    margin-bottom:18px;
">
    <div style="
        color:#F3D777;
        font-size:11px;
        font-weight:700;
        letter-spacing:0.15em;
        text-transform:uppercase;
        margin-bottom:8px;
    ">
        Fundación Jorge Vergara · Impact Intelligence
    </div>

    <div style="
        color:white;
        font-size:38px;
        font-weight:750;
        letter-spacing:-0.045em;
    ">
        FJV Decision Copilot
    </div>

    <div style="
        color:rgba(255,255,255,.70);
        font-size:14px;
        margin-top:8px;
    ">
        KPI analysis · related signals · management recommendations · 2027 forecast
    </div>
</div>
""", unsafe_allow_html=True)
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.markdown("""
<style>

/* ===== FJV COLORS ===== */
:root {
    --fjv-purple: #6D43A2;
    --fjv-orange: #F18A3A;
    --fjv-teal: #173F46;
    --fjv-sky: #8FCBE6;
    --fjv-yellow: #F3D777;
    --fjv-cream: #F7F2E9;
    --fjv-white: #FFFDF9;
    --fjv-ink: #263438;
    --fjv-muted: #758185;
    --fjv-border: #E6DFD5;
}

/* PAGE */
[data-testid="stAppViewContainer"] {
    background: #F7F2E9;
}

.block-container {
    max-width: 1320px;
    padding-top: 2rem;
    padding-bottom: 4rem;
}

/* Remove unnecessary Streamlit chrome */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

/* TITLES */
h1 {
    color: #173F46;
    font-size: 2.4rem !important;
    letter-spacing: -0.04em;
}

h2, h3 {
    color: #263438;
    letter-spacing: -0.025em;
}

/* KPI SELECTOR */
div[data-baseweb="select"] > div {
    background: #FFFDF9 !important;
    border: 1px solid #E6DFD5 !important;
    border-radius: 14px !important;
    min-height: 52px;
}

/* METRIC CARDS */
[data-testid="stMetric"] {
    background: #FFFDF9;
    border: 1px solid #E6DFD5;
    padding: 18px;
    border-radius: 18px;
    min-height: 125px;
    box-shadow: 0px 6px 20px rgba(30,40,42,0.04);
}

[data-testid="stMetricLabel"] {
    color: #758185;
    font-size: 0.78rem;
    font-weight: 650;
}

[data-testid="stMetricValue"] {
    color: #173F46;
    font-weight: 750;
    letter-spacing: -0.04em;
}

/* Give the forecast card a little distinction */
div[data-testid="column"]:nth-of-type(5) [data-testid="stMetric"] {
    background: #173F46;
    border-color: #173F46;
}

div[data-testid="column"]:nth-of-type(5) [data-testid="stMetricLabel"],
div[data-testid="column"]:nth-of-type(5) [data-testid="stMetricValue"] {
    color: white !important;
}

/* INFO BANNER */
[data-testid="stAlert"] {
    background: #EEE7F6;
    color: #503279;
    border: none;
    border-radius: 14px;
}

/* TABS */
button[data-baseweb="tab"] {
    font-weight: 650;
    font-size: 0.9rem;
}

div[data-baseweb="tab-highlight"] {
    background: #6D43A2 !important;
}

/* SUCCESS / RECOMMENDED ACTION */
[data-testid="stAlert"][data-baseweb] {
    border-radius: 16px;
}

/* DATAFRAME */
[data-testid="stDataFrame"] {
    background: #FFFDF9;
    border: 1px solid #E6DFD5;
    border-radius: 16px;
    overflow: hidden;
}

/* CHAT */
[data-testid="stChatInput"] {
    border-radius: 16px;
}

[data-testid="stChatMessage"] {
    background: #FFFDF9;
    border-radius: 16px;
    padding: 8px 14px;
    margin-bottom: 8px;
}

/* DIVIDERS */
hr {
    border-color: #E6DFD5;
}

/* GENERAL TEXT */
p {
    color: #536267;
}

</style>
""", unsafe_allow_html=True)
DATA_FILE = Path(__file__).parent / "FJV_AI_Copilot_Context.xlsx"


# ============================================================
# FJV-INSPIRED DESIGN TOKENS
# Palette intentionally approximates the visual language seen
# in FJV public-facing material. It is not presented as an
# official brand manual / exact HEX specification.
# ============================================================

INK = "#1D2B2E"
TEAL = "#173F46"
TEAL_2 = "#24555B"
PURPLE = "#6D43A2"
ORANGE = "#F18A3A"
SKY = "#8FCBE6"
YELLOW = "#F3D777"
CREAM = "#F7F2E9"
PAPER = "#FFFDF9"
WHITE = "#FFFFFF"
MUTED = "#738084"
BORDER = "#E8E0D5"
GREEN = "#2F7C66"
RED = "#B65754"
SOFT_PURPLE = "#EEE7F6"
SOFT_ORANGE = "#FCEADB"
SOFT_SKY = "#E7F4FA"
SOFT_YELLOW = "#FBF3D5"


# ============================================================
# CSS
# ============================================================

st.markdown(
    f"""
<style>
/* ------------------------------------------------------------
   FOUNDATION
------------------------------------------------------------ */
html, body, [class*="css"] {{
    font-family: Inter, ui-sans-serif, -apple-system, BlinkMacSystemFont,
                 "Segoe UI", sans-serif;
}}

[data-testid="stAppViewContainer"] {{
    background:
        radial-gradient(circle at 97% 3%, rgba(143,203,230,.22), transparent 18%),
        radial-gradient(circle at 3% 28%, rgba(109,67,162,.09), transparent 17%),
        {CREAM};
    color: {INK};
}}

[data-testid="stHeader"] {{
    background: rgba(247,242,233,.80);
    backdrop-filter: blur(10px);
}}

.block-container {{
    max-width: 1380px;
    padding-top: 1.15rem;
    padding-bottom: 4rem;
}}

#MainMenu {{ visibility: hidden; }}
footer {{ visibility: hidden; }}

h1, h2, h3 {{
    color: {INK};
    letter-spacing: -.035em;
}}

/* ------------------------------------------------------------
   BRAND NAV
------------------------------------------------------------ */
.fjv-nav {{
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:16px;
    margin-bottom:12px;
}}

.fjv-lockup {{
    display:flex;
    align-items:center;
    gap:11px;
    color:{TEAL};
}}

.fjv-mark {{
    width:34px;
    height:34px;
    border-radius:11px 11px 11px 3px;
    background:{PURPLE};
    position:relative;
    box-shadow:8px 7px 0 {ORANGE};
}}

.fjv-lockup-copy {{
    line-height:1.0;
}}

.fjv-lockup-copy b {{
    display:block;
    font-size:12px;
    letter-spacing:.12em;
}}

.fjv-lockup-copy span {{
    display:block;
    color:{MUTED};
    font-size:9px;
    letter-spacing:.11em;
    margin-top:4px;
}}

.prototype-tag {{
    display:inline-flex;
    align-items:center;
    gap:7px;
    padding:7px 11px;
    border-radius:999px;
    border:1px solid {BORDER};
    background:rgba(255,255,255,.62);
    color:{MUTED};
    font-size:9px;
    font-weight:750;
    letter-spacing:.06em;
    text-transform:uppercase;
}}

.prototype-dot {{
    width:7px;
    height:7px;
    border-radius:50%;
    background:{ORANGE};
}}

/* ------------------------------------------------------------
   HERO
------------------------------------------------------------ */
.hero {{
    position:relative;
    min-height:318px;
    overflow:hidden;
    border-radius:30px;
    background:{TEAL};
    color:white;
    padding:40px 44px 34px 44px;
    box-shadow:0 24px 60px rgba(23,63,70,.16);
    margin-bottom:16px;
}}

.hero-copy {{
    position:relative;
    z-index:5;
    max-width:820px;
}}

.hero-kicker {{
    color:{YELLOW};
    font-size:10px;
    font-weight:850;
    letter-spacing:.19em;
    text-transform:uppercase;
    margin-bottom:13px;
}}

.hero-title {{
    font-size:clamp(36px,5.2vw,67px);
    line-height:.94;
    font-weight:800;
    letter-spacing:-.058em;
    max-width:850px;
}}

.hero-title em {{
    font-family:Georgia, "Times New Roman", serif;
    font-weight:500;
    color:{SKY};
}}

.hero-sub {{
    margin-top:19px;
    max-width:650px;
    color:rgba(255,255,255,.76);
    font-size:14px;
    line-height:1.65;
}}

.hero-tags {{
    display:flex;
    flex-wrap:wrap;
    gap:8px;
    margin-top:24px;
}}

.hero-tag {{
    font-size:9px;
    font-weight:800;
    text-transform:uppercase;
    letter-spacing:.08em;
    padding:8px 11px;
    border-radius:999px;
    border:1px solid rgba(255,255,255,.16);
    background:rgba(255,255,255,.07);
    color:rgba(255,255,255,.90);
}}

/* decorative organic pieces */
.blob {{
    position:absolute;
    z-index:1;
    pointer-events:none;
}}

.blob-purple {{
    width:280px;
    height:280px;
    right:-75px;
    top:-95px;
    background:{PURPLE};
    border-radius:43% 57% 58% 42% / 54% 38% 62% 46%;
    transform:rotate(18deg);
}}

.blob-orange {{
    width:185px;
    height:115px;
    right:170px;
    bottom:-55px;
    background:{ORANGE};
    border-radius:58% 42% 49% 51% / 60% 50% 50% 40%;
    transform:rotate(-13deg);
}}

.blob-yellow {{
    width:76px;
    height:76px;
    right:315px;
    top:48px;
    background:{YELLOW};
    border-radius:50% 50% 12% 50%;
    transform:rotate(18deg);
}}

.blob-sky {{
    width:125px;
    height:45px;
    right:80px;
    bottom:74px;
    background:{SKY};
    border-radius:999px;
    transform:rotate(-22deg);
}}

.spark {{
    position:absolute;
    right:342px;
    bottom:66px;
    z-index:3;
    color:white;
    font-size:44px;
    font-family:Georgia,serif;
    transform:rotate(15deg);
}}

/* ------------------------------------------------------------
   JOURNEY STRIP
------------------------------------------------------------ */
.journey {{
    display:grid;
    grid-template-columns:repeat(5,1fr);
    gap:0;
    margin:2px 0 23px 0;
    border:1px solid {BORDER};
    border-radius:19px;
    overflow:hidden;
    background:rgba(255,255,255,.67);
}}

.journey-step {{
    position:relative;
    padding:13px 14px;
    border-right:1px solid {BORDER};
}}

.journey-step:last-child {{
    border-right:none;
}}

.journey-no {{
    font-size:8px;
    font-weight:850;
    letter-spacing:.12em;
    color:{MUTED};
}}

.journey-label {{
    font-size:11px;
    font-weight:780;
    color:{INK};
    margin-top:3px;
}}

.journey-step:nth-child(1) {{ box-shadow: inset 0 4px 0 {PURPLE}; }}
.journey-step:nth-child(2) {{ box-shadow: inset 0 4px 0 {SKY}; }}
.journey-step:nth-child(3) {{ box-shadow: inset 0 4px 0 {YELLOW}; }}
.journey-step:nth-child(4) {{ box-shadow: inset 0 4px 0 {ORANGE}; }}
.journey-step:nth-child(5) {{ box-shadow: inset 0 4px 0 {TEAL}; }}

/* ------------------------------------------------------------
   SECTION HEADERS + SELECT
------------------------------------------------------------ */
.section-overline {{
    font-size:9px;
    font-weight:850;
    letter-spacing:.16em;
    text-transform:uppercase;
    color:{PURPLE};
    margin:0 0 7px 2px;
}}

.section-title {{
    font-size:24px;
    font-weight:790;
    letter-spacing:-.035em;
    color:{INK};
    margin-bottom:6px;
}}

.section-description {{
    color:{MUTED};
    font-size:12px;
    line-height:1.55;
    margin-bottom:12px;
}}

div[data-baseweb="select"] > div {{
    min-height:52px;
    background:{PAPER} !important;
    border:1px solid {BORDER} !important;
    border-radius:16px !important;
    box-shadow:0 8px 25px rgba(39,52,55,.045);
}}

/* ------------------------------------------------------------
   KPI STORY
------------------------------------------------------------ */
.kpi-story {{
    background:{PAPER};
    border:1px solid {BORDER};
    border-radius:23px;
    padding:22px 23px 20px 23px;
    margin:12px 0 13px 0;
    box-shadow:0 12px 34px rgba(40,49,51,.045);
}}

.kpi-mini {{
    font-size:9px;
    font-weight:850;
    letter-spacing:.14em;
    color:{ORANGE};
    text-transform:uppercase;
    margin-bottom:7px;
}}

.kpi-big {{
    font-size:clamp(23px,3vw,35px);
    line-height:1.07;
    font-weight:800;
    letter-spacing:-.045em;
    color:{INK};
    max-width:1080px;
}}

.kpi-objective {{
    color:{MUTED};
    margin-top:9px;
    font-size:12px;
    line-height:1.55;
}}

.chips {{
    display:flex;
    flex-wrap:wrap;
    gap:7px;
    margin-top:15px;
}}

.chip {{
    background:#F1ECE5;
    color:#566469;
    padding:6px 9px;
    border-radius:999px;
    font-size:9px;
    font-weight:750;
}}

/* ------------------------------------------------------------
   METRIC CARDS
------------------------------------------------------------ */
.metric-grid {{
    display:grid;
    grid-template-columns:repeat(5,minmax(0,1fr));
    gap:10px;
    margin:10px 0 24px 0;
}}

.m-card {{
    min-height:128px;
    background:{PAPER};
    border:1px solid {BORDER};
    border-radius:20px;
    padding:17px 17px 15px 17px;
    position:relative;
    overflow:hidden;
    box-shadow:0 10px 30px rgba(40,49,51,.04);
}}

.m-accent {{
    position:absolute;
    height:7px;
    width:53px;
    left:17px;
    top:0;
    border-radius:0 0 999px 999px;
}}

.m-label {{
    margin-top:6px;
    color:{MUTED};
    font-size:8.5px;
    font-weight:850;
    letter-spacing:.10em;
    text-transform:uppercase;
}}

.m-value {{
    color:{INK};
    font-weight:800;
    letter-spacing:-.045em;
    font-size:clamp(22px,2.3vw,34px);
    line-height:1.02;
    margin-top:16px;
}}

.m-note {{
    color:#8E989B;
    font-size:9px;
    line-height:1.35;
    margin-top:8px;
}}

.forecast-card {{
    background:{TEAL};
    border-color:{TEAL};
}}

.forecast-card .m-label,
.forecast-card .m-note {{
    color:rgba(255,255,255,.65);
}}

.forecast-card .m-value {{
    color:white;
}}

.forecast-card:after {{
    content:"";
    position:absolute;
    width:85px;
    height:85px;
    background:{PURPLE};
    border-radius:50%;
    right:-44px;
    bottom:-45px;
}}

/* ------------------------------------------------------------
   CHART SHELL
------------------------------------------------------------ */
.chart-shell {{
    border:1px solid {BORDER};
    background:{PAPER};
    border-radius:24px;
    padding:17px 19px 8px 19px;
    box-shadow:0 12px 34px rgba(40,49,51,.04);
    margin-bottom:13px;
}}

.chart-head {{
    display:flex;
    justify-content:space-between;
    align-items:flex-end;
    gap:14px;
    margin-bottom:0;
}}

.chart-title {{
    color:{INK};
    font-size:18px;
    font-weight:790;
    letter-spacing:-.025em;
}}

.chart-note {{
    color:{MUTED};
    font-size:9px;
    text-align:right;
    max-width:360px;
    line-height:1.4;
}}

.forecast-pill {{
    display:inline-flex;
    gap:6px;
    align-items:center;
    padding:5px 8px;
    border-radius:999px;
    background:{SOFT_YELLOW};
    color:#7A6924;
    font-size:8px;
    font-weight:850;
    letter-spacing:.05em;
    text-transform:uppercase;
}}

/* ------------------------------------------------------------
   SIGNAL RADAR
------------------------------------------------------------ */
.signal-grid {{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:12px;
    margin:9px 0 17px 0;
}}

.signal-panel {{
    border-radius:22px;
    border:1px solid {BORDER};
    padding:17px;
    background:{PAPER};
}}

.signal-panel.attention {{
    background:linear-gradient(145deg,#FFF9F5,{SOFT_ORANGE});
}}

.signal-panel.positive {{
    background:linear-gradient(145deg,#FAFCFD,{SOFT_SKY});
}}

.signal-title {{
    font-size:10px;
    font-weight:850;
    text-transform:uppercase;
    letter-spacing:.10em;
    margin-bottom:11px;
}}

.signal-row {{
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:10px;
    padding:10px 0;
    border-top:1px solid rgba(40,55,58,.08);
}}

.signal-row:first-of-type {{
    border-top:none;
}}

.signal-name {{
    font-size:10px;
    font-weight:720;
    color:{INK};
    line-height:1.28;
}}

.signal-code {{
    color:{MUTED};
    font-size:8px;
    margin-top:3px;
}}

.signal-val {{
    white-space:nowrap;
    font-size:12px;
    font-weight:820;
}}

/* ------------------------------------------------------------
   DECISION BOARD
------------------------------------------------------------ */
.decision-grid {{
    display:grid;
    grid-template-columns:1.05fr 1.05fr 1.45fr;
    gap:11px;
    margin-top:9px;
    margin-bottom:13px;
}}

.d-card {{
    border:1px solid {BORDER};
    background:{PAPER};
    border-radius:22px;
    padding:18px 18px 17px 18px;
    min-height:210px;
    box-shadow:0 10px 28px rgba(40,49,51,.035);
}}

.d-card.action {{
    background:{PURPLE};
    border-color:{PURPLE};
    color:white;
    position:relative;
    overflow:hidden;
}}

.d-card.action:after {{
    content:"";
    position:absolute;
    width:135px;
    height:135px;
    right:-70px;
    bottom:-70px;
    background:{ORANGE};
    border-radius:50%;
}}

.d-kicker {{
    font-size:8.5px;
    font-weight:850;
    letter-spacing:.13em;
    text-transform:uppercase;
    color:{MUTED};
    margin-bottom:11px;
}}

.d-card.action .d-kicker {{
    color:{YELLOW};
}}

.d-copy {{
    color:#405055;
    font-size:11px;
    line-height:1.65;
    position:relative;
    z-index:2;
}}

.d-card.action .d-copy {{
    color:rgba(255,255,255,.93);
    font-size:12px;
    font-weight:550;
}}

.d-bullet {{
    padding:8px 0;
    border-top:1px solid rgba(38,52,55,.07);
}}

.d-bullet:first-child {{
    border-top:none;
    padding-top:0;
}}

/* ------------------------------------------------------------
   CONFIDENCE / GAP
------------------------------------------------------------ */
.confidence-wrap {{
    border-radius:22px;
    border:1px solid {BORDER};
    background:{TEAL};
    color:white;
    padding:18px 20px;
    margin:8px 0 16px 0;
}}

.confidence-head {{
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:10px;
}}

.conf-title {{
    font-size:16px;
    font-weight:780;
}}

.conf-pill {{
    padding:6px 9px;
    border-radius:999px;
    background:rgba(255,255,255,.10);
    border:1px solid rgba(255,255,255,.15);
    color:{SKY};
    font-size:9px;
    font-weight:850;
    text-transform:uppercase;
    letter-spacing:.07em;
}}

.conf-copy {{
    margin-top:11px;
    color:rgba(255,255,255,.72);
    font-size:10px;
    line-height:1.6;
}}

/* ------------------------------------------------------------
   TABS / TABLE / CHAT
------------------------------------------------------------ */
button[data-baseweb="tab"] {{
    font-size:11px !important;
    font-weight:800 !important;
}}

div[data-baseweb="tab-highlight"] {{
    background:{PURPLE} !important;
}}

[data-testid="stDataFrame"] {{
    border:1px solid {BORDER};
    border-radius:18px;
    overflow:hidden;
}}

.chat-shell {{
    margin-top:22px;
    background:{PAPER};
    border:1px solid {BORDER};
    border-radius:26px;
    padding:22px 23px 12px 23px;
    box-shadow:0 12px 34px rgba(40,49,51,.04);
}}

.chat-kicker {{
    color:{ORANGE};
    font-size:9px;
    font-weight:850;
    text-transform:uppercase;
    letter-spacing:.13em;
}}

.chat-title {{
    color:{INK};
    font-size:25px;
    font-weight:800;
    letter-spacing:-.035em;
    margin-top:5px;
}}

.chat-copy {{
    color:{MUTED};
    font-size:11px;
    line-height:1.55;
    margin:5px 0 8px 0;
}}

[data-testid="stChatInput"] {{
    border-radius:16px !important;
}}

.quick-prompts {{
    display:flex;
    gap:7px;
    flex-wrap:wrap;
    margin:10px 0 3px 0;
}}

.qp {{
    font-size:8.5px;
    font-weight:750;
    color:{TEAL};
    background:{SOFT_SKY};
    padding:6px 8px;
    border-radius:999px;
}}

/* ------------------------------------------------------------
   FOOTER
------------------------------------------------------------ */
.brand-footer {{
    margin-top:30px;
    border-top:1px solid {BORDER};
    padding:22px 3px 3px 3px;
    display:flex;
    justify-content:space-between;
    align-items:flex-end;
    gap:20px;
    flex-wrap:wrap;
}}

.brand-footer-big {{
    color:{TEAL};
    font-family:Georgia,serif;
    font-style:italic;
    font-size:20px;
}}

.brand-footer-small {{
    color:{MUTED};
    font-size:8.5px;
    text-align:right;
    line-height:1.5;
}}

/* ------------------------------------------------------------
   RESPONSIVE
------------------------------------------------------------ */
@media (max-width: 980px) {{
    .metric-grid {{
        grid-template-columns:repeat(2,minmax(0,1fr));
    }}
    .decision-grid {{
        grid-template-columns:1fr;
    }}
    .signal-grid {{
        grid-template-columns:1fr;
    }}
}}

@media (max-width: 620px) {{
    .hero {{
        padding:28px 24px;
    }}
    .journey {{
        grid-template-columns:1fr;
    }}
    .journey-step {{
        border-right:none;
        border-bottom:1px solid {BORDER};
    }}
    .metric-grid {{
        grid-template-columns:1fr;
    }}
    .blob-yellow, .spark {{
        display:none;
    }}
}}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# DATA
# ============================================================

@st.cache_data
def load_data():
    kpis = pd.read_excel(DATA_FILE, sheet_name="01_KPI_Context")
    related = pd.read_excel(DATA_FILE, sheet_name="02_Related_KPIs")
    values = pd.read_excel(DATA_FILE, sheet_name="05_Value_Dictionary")
    return kpis, related, values


def esc(value):
    if value is None or pd.isna(value):
        return "N/D"
    return html.escape(str(value))


def fmt_value(value, unit):
    if value is None or pd.isna(value):
        return "N/D"

    unit = "" if unit is None or pd.isna(unit) else str(unit)

    if unit == "%":
        return f"{float(value):,.1f}%"

    if "MXN" in unit or "USD" in unit or "Moneda" in unit:
        return f"{float(value):,.2f} {unit}"

    if "#" in unit:
        return f"{float(value):,.0f} {unit}"

    return f"{float(value):,.1f} {unit}".strip()


def parse_number(text):
    if text is None or pd.isna(text):
        return None
    m = re.search(r"(-?\d+(?:\.\d+)?)", str(text).replace(",", ""))
    return float(m.group(1)) if m else None


# ============================================================
# TARGET + FORECAST
# ============================================================

def evaluate_target(row):
    target = "" if pd.isna(row["Target_2026"]) else str(row["Target_2026"])
    value = row["2026"]
    baseline = row["Baseline_3Y"]
    direction = str(row["Direction"])

    if pd.isna(value):
        return None, "No hay resultado 2026 suficiente para evaluar la meta."

    checks = []
    explanations = []

    le = re.search(r"(?:≤|<=)\s*(\d+(?:\.\d+)?)", target)
    ge = re.search(r"(?:≥|>=)\s*(\d+(?:\.\d+)?)", target)
    eq_pct = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%\s*", target)

    if le:
        lim = float(le.group(1))
        checks.append(float(value) <= lim)
        explanations.append(
            f"resultado {fmt_value(value, row['Unit'])} vs límite ≤{lim:g}"
        )

    elif ge:
        lim = float(ge.group(1))
        if "reducción" not in target.lower() or direction == "↑":
            checks.append(float(value) >= lim)
            explanations.append(
                f"resultado {fmt_value(value, row['Unit'])} vs mínimo ≥{lim:g}"
            )

    elif eq_pct:
        lim = float(eq_pct.group(1))
        checks.append(
            float(value) >= lim if direction == "↑" else float(value) <= lim
        )
        explanations.append(
            f"resultado {fmt_value(value, row['Unit'])} vs meta {lim:g}%"
        )

    red = re.search(
        r"(?:≥|>=)\s*(\d+(?:\.\d+)?)\s*%\s*de\s*reducci",
        target.lower(),
    )

    if red and not pd.isna(baseline) and float(baseline) != 0:
        required = float(red.group(1)) / 100.0
        actual = (float(baseline) - float(value)) / abs(float(baseline))
        checks.append(actual >= required)
        explanations.append(
            f"reducción vs baseline {actual:.1%} vs requerida {required:.1%}"
        )

    if not checks:
        return None, f"Meta textual: {target or 'N/D'}"

    return all(checks), "; ".join(explanations)


def target_numeric(row):
    target = "" if pd.isna(row["Target_2026"]) else str(row["Target_2026"])
    m = re.search(r"(?:≤|<=|≥|>=)?\s*(\d+(?:\.\d+)?)", target)
    return float(m.group(1)) if m else None


def forecast_2027(row, rel):
    values = [row["2023"], row["2024"], row["2025"], row["2026"]]
    years = [2023, 2024, 2025, 2026]

    if any(pd.isna(v) for v in values):
        return None, "Histórico insuficiente"

    x_mean = sum(years) / len(years)
    y_mean = sum(values) / len(values)

    numerator = sum(
        (x - x_mean) * (y - y_mean)
        for x, y in zip(years, values)
    )

    denominator = sum((x - x_mean) ** 2 for x in years)
    slope = numerator / denominator if denominator != 0 else 0

    forecast = y_mean + slope * (2027 - x_mean)

    # modest context adjustment
    if not rel.empty:
        related_momentum = rel["Related_Performance_Change"].dropna()

        if len(related_momentum) > 0:
            avg_related = float(related_momentum.mean())
            adjustment = max(min(avg_related * 0.20, 0.10), -0.10)

            if str(row["Direction"]) == "↑":
                forecast *= (1 + adjustment)

            elif str(row["Direction"]) == "↓":
                forecast *= (1 - adjustment)

    unit = str(row["Unit"]).lower()

    if "%" in unit:
        forecast = max(0, min(100, forecast))

    elif (
        "día" in unit
        or "day" in unit
        or "#" in unit
        or "proyecto" in unit
        or "alianza" in unit
    ):
        forecast = max(0, forecast)

    return forecast, "tendencia 2023–2026 + señal contextual de KPIs relacionados"


# ============================================================
# DATA QUALITY
# ============================================================

def suspicious_units(row, value_dict):
    issues = []

    subset = value_dict[value_dict["KPI_ID"] == row["KPI_ID"]]

    if subset.empty:
        return issues

    for _, v in subset.iterrows():
        canonical = str(v.get("Canonical_Field", "")).lower()
        unit = str(v.get("Unit", "")).lower()
        role = str(v.get("Input_Role", ""))

        if (
            ("count" in canonical or "request" in canonical)
            and ("día" in unit or "day" in unit)
        ):
            issues.append(
                f"{role}: '{v.get('Value_Name')}' está etiquetado como "
                f"'{v.get('Unit')}', aunque parece ser un conteo."
            )

        if (
            row.get("Formula_Type") == "RATIO_PCT"
            and role in ("NUM", "DEN")
            and "moneda" in unit
        ):
            issues.append(
                f"{role}: '{v.get('Value_Name')}' aparece con unidad monetaria "
                "dentro de un ratio porcentual; conviene validarla."
            )

    return issues


# ============================================================
# DECISION LOGIC
# ============================================================

def build_analysis(row, rel, value_dict):
    perf = (
        float(row["Performance_Change"])
        if not pd.isna(row["Performance_Change"])
        else None
    )

    target_met, target_detail = evaluate_target(row)

    deteriorating = rel[
        rel["Related_Performance_Change"] < 0
    ].sort_values("Related_Performance_Change")

    improving = rel[
        rel["Related_Performance_Change"] >= 0
    ].sort_values("Related_Performance_Change", ascending=False)

    if target_met is True and (perf is None or perf >= 0):
        diagnosis = (
            f"{row['KPI_ID']} está cumpliendo la meta 2026 y mantiene una señal "
            "positiva frente a su baseline. Con estos datos no necesita una "
            "corrección inmediata."
        )

    elif target_met is False and perf is not None and perf >= 0:
        diagnosis = (
            f"{row['KPI_ID']} está mejorando frente al baseline, pero todavía no "
            "cumple completamente la meta 2026. El reto es cerrar el gap restante "
            "sin perder las mejoras ya obtenidas."
        )

    elif perf is not None and perf < 0:
        diagnosis = (
            f"{row['KPI_ID']} muestra deterioro frente al baseline ajustado por "
            "dirección. Conviene investigar la desviación antes de escalar o "
            "replicar el proceso actual."
        )

    else:
        diagnosis = (
            f"{row['KPI_ID']} necesita revisión adicional porque la información "
            "disponible no permite clasificar con suficiente claridad el "
            "cumplimiento de meta y la tendencia."
        )

    evidence = [
        f"Resultado 2026: {fmt_value(row['2026'], row['Unit'])}.",
        f"Baseline 3Y: {fmt_value(row['Baseline_3Y'], row['Unit'])}.",
        f"Meta 2026: {row['Target_2026']}.",
        (
            f"Momentum ajustado por dirección: {perf:+.1%}."
            if perf is not None
            else "Momentum: N/D."
        ),
        f"Lectura de meta: {target_detail}.",
    ]

    if target_met is True:
        if not deteriorating.empty:
            action = (
                f"Mantener lo que está sosteniendo el desempeño de "
                f"{row['KPI_ID']} y usar la evidencia para protegerlo. "
                f"Como siguiente foco de gestión, revisar "
                f"{deteriorating.iloc[0]['Related_KPI_ID']}, que es la señal "
                "relacionada con mayor deterioro. Validar datos operativos antes "
                "de asumir una relación causal."
            )
        else:
            weakest = rel.sort_values("Related_Performance_Change").head(1)

            if not weakest.empty:
                action = (
                    "Mantener el desempeño actual y utilizar el control anti-gaming "
                    "como revisión periódica de calidad. Después, revisar "
                    f"{weakest.iloc[0]['Related_KPI_ID']} como el punto relativamente "
                    "más débil del contexto."
                )
            else:
                action = (
                    "Mantener el desempeño actual y realizar controles periódicos "
                    "de calidad y evidencia."
                )

    elif target_met is False and perf is not None and perf >= 0:
        action = (
            "Cerrar el gap restante a meta sin deshacer la mejora lograda. "
            "Desglosar el proceso o sus inputs para ubicar en qué etapa, segmento "
            "o causa se concentra la diferencia restante."
        )

    else:
        if not deteriorating.empty:
            action = (
                f"Priorizar el diagnóstico de la desviación y contrastarlo con "
                f"{deteriorating.iloc[0]['Related_KPI_ID']}, la señal relacionada "
                "más débil. Validar primero datos por etapa, segmento o responsable "
                "antes de modificar la estrategia."
            )
        else:
            action = (
                "Priorizar el diagnóstico de la desviación, validar la calidad del "
                "dato y después comparar el KPI con las señales relacionadas antes "
                "de tomar una decisión."
            )

    gaps = suspicious_units(row, value_dict)

    gaps.append(
        "Los resultados son agregados; para explicar causas hacen falta datos "
        "operativos más granulares por etapa, segmento, causa, responsable o periodo."
    )

    gaps.append(
        "Las relaciones entre KPIs son heurísticas de recuperación y no deben "
        "interpretarse como causalidad."
    )

    confidence = "Moderada"

    if target_met is not None and perf is not None and len(gaps) == 2:
        confidence = "Moderada-alta para desempeño · Moderada para causas"

    elif len(gaps) > 2:
        confidence = "Moderada-baja hasta validar calidad de datos"

    return {
        "diagnosis": diagnosis,
        "evidence": evidence,
        "action": action,
        "gaps": gaps,
        "confidence": confidence,
        "target_met": target_met,
        "deteriorating": deteriorating,
        "improving": improving,
    }


# ============================================================
# HTML HELPERS
# ============================================================

def metric_card(label, value, note, accent, forecast=False):
    forecast_class = " forecast-card" if forecast else ""

    return f"""
    <div class="m-card{forecast_class}">
        <div class="m-accent" style="background:{accent};"></div>
        <div class="m-label">{esc(label)}</div>
        <div class="m-value">{esc(value)}</div>
        <div class="m-note">{esc(note)}</div>
    </div>
    """


def signal_rows(df, positive=True, limit=3):
    if df.empty:
        return """
        <div class="signal-row">
            <div class="signal-name">Sin señales suficientes</div>
        </div>
        """

    rows = []

    for _, r in df.head(limit).iterrows():
        val = float(r["Related_Performance_Change"])
        val_color = GREEN if val >= 0 else RED

        rows.append(
            f"""
            <div class="signal-row">
                <div>
                    <div class="signal-name">{esc(r['Related_KPI_Name'])}</div>
                    <div class="signal-code">{esc(r['Related_KPI_ID'])}</div>
                </div>
                <div class="signal-val" style="color:{val_color};">
                    {val:+.1%}
                </div>
            </div>
            """
        )

    return "".join(rows)


def evidence_html(items):
    return "".join(
        f'<div class="d-bullet">{esc(item)}</div>'
        for item in items
    )


def gaps_html(items):
    return "<br>".join(f"• {esc(item)}" for item in items)


# ============================================================
# LOAD + HEADER
# ============================================================

kpis, related, value_dict = load_data()

st.markdown(
    """
<div class="fjv-nav">
    <div class="fjv-lockup">
        <div class="fjv-mark"></div>
        <div class="fjv-lockup-copy">
            <b>FUNDACIÓN JORGE VERGARA</b>
            <span>IMPACT INTELLIGENCE</span>
        </div>
    </div>
    <div class="prototype-tag">
        <span class="prototype-dot"></span>
        Decision-support prototype
    </div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="hero">
    <div class="blob blob-purple"></div>
    <div class="blob blob-orange"></div>
    <div class="blob blob-yellow"></div>
    <div class="blob blob-sky"></div>
    <div class="spark">✦</div>

    <div class="hero-copy">
        <div class="hero-kicker">FJV · Decision Copilot</div>

        <div class="hero-title">
            De indicadores a <em>decisiones</em><br>
            que cuidan el impacto.
        </div>

        <div class="hero-sub">
            Una capa de inteligencia para conectar desempeño, señales relacionadas,
            riesgos y acciones de gestión — sin perder de vista a las personas detrás
            de cada indicador.
        </div>

        <div class="hero-tags">
            <span class="hero-tag">Strategy</span>
            <span class="hero-tag">Performance</span>
            <span class="hero-tag">Signals</span>
            <span class="hero-tag">2027 Forecast</span>
            <span class="hero-tag">Human Review</span>
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="journey">
    <div class="journey-step">
        <div class="journey-no">01</div>
        <div class="journey-label">Estrategia</div>
    </div>
    <div class="journey-step">
        <div class="journey-no">02</div>
        <div class="journey-label">KPI</div>
    </div>
    <div class="journey-step">
        <div class="journey-no">03</div>
        <div class="journey-label">Señales</div>
    </div>
    <div class="journey-step">
        <div class="journey-no">04</div>
        <div class="journey-label">Decisión</div>
    </div>
    <div class="journey-step">
        <div class="journey-no">05</div>
        <div class="journey-label">Acción</div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ============================================================
# KPI SELECTOR
# ============================================================

st.markdown(
    """
<div class="section-overline">Explore the system</div>
<div class="section-title">¿Qué historia quieres entender?</div>
<div class="section-description">
Selecciona un KPI. El Copilot reconstruye su desempeño, su contexto y las señales
que merecen atención.
</div>
""",
    unsafe_allow_html=True,
)

labels = {
    f"{r.KPI_ID} — {r.KPI_Name}": r.KPI_ID
    for r in kpis.itertuples()
}

selected_label = st.selectbox(
    "KPI",
    list(labels.keys()),
    label_visibility="collapsed",
)

selected_id = labels[selected_label]

row = kpis[kpis["KPI_ID"] == selected_id].iloc[0]

rel = related[
    related["Focal_KPI_ID"] == selected_id
].sort_values("Rank")

analysis = build_analysis(row, rel, value_dict)
forecast_value, forecast_method = forecast_2027(row, rel)

perf = (
    float(row["Performance_Change"])
    if not pd.isna(row["Performance_Change"])
    else None
)

status_text = (
    "META ALCANZADA"
    if analysis["target_met"] is True
    else "GAP A META"
    if analysis["target_met"] is False
    else "REVISAR"
)

st.markdown(
    f"""
<div class="kpi-story">
    <div class="kpi-mini">{esc(row['KPI_ID'])} · {esc(status_text)}</div>
    <div class="kpi-big">{esc(row['KPI_Name'])}</div>
    <div class="kpi-objective">{esc(row['Strategic_Objective'])}</div>

    <div class="chips">
        <span class="chip">{esc(row['Department'])}</span>
        <span class="chip">{esc(row['BSC_Perspective'])}</span>
        <span class="chip">Owner · {esc(row['Owner'])}</span>
        <span class="chip">{esc(row['Frequency'])}</span>
        <span class="chip">Dirección · {esc(row['Direction'])}</span>
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ============================================================
# METRICS
# ============================================================

momentum_text = f"{perf:+.1%}" if perf is not None else "N/D"

forecast_text = (
    fmt_value(forecast_value, row["Unit"])
    if forecast_value is not None
    else "N/D"
)

cards = (
    metric_card(
        "Resultado 2026",
        fmt_value(row["2026"], row["Unit"]),
        "Resultado actual",
        PURPLE,
    )
    + metric_card(
        "Meta",
        str(row["Target_2026"]),
        "Umbral de éxito",
        ORANGE,
    )
    + metric_card(
        "Baseline 3Y",
        fmt_value(row["Baseline_3Y"], row["Unit"]),
        "Promedio 2023–2025",
        SKY,
    )
    + metric_card(
        "Momentum",
        momentum_text,
        "Ajustado por dirección",
        YELLOW,
    )
    + metric_card(
        "2027 Forecast",
        forecast_text,
        "Proyección estimada",
        SKY,
        forecast=True,
    )
)

st.markdown(
    f'<div class="metric-grid">{cards}</div>',
    unsafe_allow_html=True,
)


# ============================================================
# TREND + FORECAST CHART
# ============================================================

st.markdown(
    f"""
<div class="chart-shell">
    <div class="chart-head">
        <div>
            <div class="section-overline">Performance journey</div>
            <div class="chart-title">De dónde venimos · hacia dónde apunta 2027</div>
        </div>
        <div class="chart-note">
            <span class="forecast-pill">✦ Estimated forecast</span><br>
            Histórico sólido · proyección punteada · apoyo a decisión, no resultado garantizado.
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

actual = pd.DataFrame(
    {
        "Year": [2023, 2024, 2025, 2026],
        "Value": [
            row["2023"],
            row["2024"],
            row["2025"],
            row["2026"],
        ],
    }
)

future = pd.DataFrame(
    {
        "Year": [2026, 2027],
        "Value": [row["2026"], forecast_value],
    }
)

actual_line = (
    alt.Chart(actual)
    .mark_line(
        color=TEAL,
        strokeWidth=4,
        interpolate="monotone",
    )
    .encode(
        x=alt.X(
            "Year:Q",
            scale=alt.Scale(domain=[2022.8, 2027.2]),
            axis=alt.Axis(
                title=None,
                values=[2023, 2024, 2025, 2026, 2027],
                format="d",
                labelColor=MUTED,
                tickColor=BORDER,
                domain=False,
                grid=False,
            ),
        ),
        y=alt.Y(
            "Value:Q",
            title=None,
            axis=alt.Axis(
                labelColor=MUTED,
                gridColor="#EEE7DE",
                domain=False,
                tickColor="transparent",
            ),
        ),
        tooltip=[
            alt.Tooltip("Year:Q", title="Año", format=".0f"),
            alt.Tooltip("Value:Q", title="Valor", format=",.2f"),
        ],
    )
)

actual_points = (
    alt.Chart(actual)
    .mark_circle(
        size=115,
        color=PAPER,
        stroke=TEAL,
        strokeWidth=3,
    )
    .encode(
        x="Year:Q",
        y="Value:Q",
        tooltip=[
            alt.Tooltip("Year:Q", title="Año", format=".0f"),
            alt.Tooltip("Value:Q", title="Valor", format=",.2f"),
        ],
    )
)

forecast_line = (
    alt.Chart(future)
    .mark_line(
        color=ORANGE,
        strokeWidth=4,
        strokeDash=[8, 7],
    )
    .encode(
        x="Year:Q",
        y="Value:Q",
        tooltip=[
            alt.Tooltip("Year:Q", title="Año", format=".0f"),
            alt.Tooltip("Value:Q", title="Forecast", format=",.2f"),
        ],
    )
)

forecast_point = (
    alt.Chart(future.tail(1))
    .mark_circle(
        size=180,
        color=YELLOW,
        stroke=ORANGE,
        strokeWidth=3,
    )
    .encode(
        x="Year:Q",
        y="Value:Q",
        tooltip=[
            alt.Tooltip("Year:Q", title="Año", format=".0f"),
            alt.Tooltip("Value:Q", title="Forecast", format=",.2f"),
        ],
    )
)

layers = actual_line + actual_points + forecast_line + forecast_point

t_num = target_numeric(row)

if t_num is not None:
    target_df = pd.DataFrame({"target": [t_num]})

    target_rule = (
        alt.Chart(target_df)
        .mark_rule(
            color=PURPLE,
            strokeDash=[3, 5],
            opacity=.55,
        )
        .encode(y="target:Q")
    )

    layers = layers + target_rule

chart = (
    layers.properties(height=330)
    .configure_view(strokeWidth=0)
)

st.altair_chart(chart, use_container_width=True)

st.caption(
    f"2027 Estimated Forecast · {forecast_method}. "
    "Los valores actuales son sintéticos/dummy para fines de prototipo."
)


# ============================================================
# SIGNAL RADAR
# ============================================================

st.markdown(
    """
<div class="section-overline" style="margin-top:24px;">Signal radar</div>
<div class="section-title">No mires este KPI solo.</div>
<div class="section-description">
El valor está en entender qué otras señales se están moviendo alrededor del indicador.
Estas relaciones ayudan a priorizar preguntas; no demuestran causalidad.
</div>
""",
    unsafe_allow_html=True,
)

negative_html = signal_rows(analysis["deteriorating"], positive=False)
positive_html = signal_rows(analysis["improving"], positive=True)

st.markdown(
    f"""
<div class="signal-grid">
    <div class="signal-panel attention">
        <div class="signal-title" style="color:{RED};">⚑ Señales para mirar</div>
        {negative_html}
    </div>

    <div class="signal-panel positive">
        <div class="signal-title" style="color:{GREEN};">✦ Señales que acompañan</div>
        {positive_html}
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ============================================================
# DECISION BOARD
# ============================================================

st.markdown(
    """
<div class="section-overline" style="margin-top:24px;">Decision board</div>
<div class="section-title">Del dato a una conversación de gestión.</div>
<div class="section-description">
El Copilot separa lo observado, lo que merece atención y la acción sugerida.
</div>
""",
    unsafe_allow_html=True,
)

evidence = evidence_html(analysis["evidence"])

if not analysis["deteriorating"].empty:
    strongest_attention = analysis["deteriorating"].iloc[0]
    attention_copy = (
        f"La señal relacionada más débil es "
        f"{strongest_attention['Related_KPI_ID']} · "
        f"{strongest_attention['Related_KPI_Name']} "
        f"({float(strongest_attention['Related_Performance_Change']):+.1%}). "
        "Debe investigarse, no asumirse como causa."
    )
else:
    attention_copy = (
        "No hay deterioros relacionados fuertes en el paquete actual. "
        "El foco puede desplazarse a sostenibilidad del desempeño, calidad "
        "del dato y prevención de gaming."
    )

st.markdown(
    f"""
<div class="decision-grid">
    <div class="d-card">
        <div class="d-kicker">01 · Lo que sabemos</div>
        <div class="d-copy">
            <div class="d-bullet">{esc(analysis['diagnosis'])}</div>
            {evidence}
        </div>
    </div>

    <div class="d-card">
        <div class="d-kicker">02 · Lo que merece atención</div>
        <div class="d-copy">
            {esc(attention_copy)}
        </div>
    </div>

    <div class="d-card action">
        <div class="d-kicker">03 · Qué haría ahora</div>
        <div class="d-copy">
            {esc(analysis['action'])}
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    f"""
<div class="confidence-wrap">
    <div class="confidence-head">
        <div class="conf-title">¿Qué tan lejos podemos llegar con estos datos?</div>
        <div class="conf-pill">{esc(analysis['confidence'])}</div>
    </div>

    <div class="conf-copy">
        {gaps_html(analysis['gaps'])}
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ============================================================
# DEEPER CONTEXT
# ============================================================

tab1, tab2, tab3 = st.tabs(
    [
        "Related KPI universe",
        "How this KPI works",
        "Governance & evidence",
    ]
)

with tab1:
    if rel.empty:
        st.info("No hay KPIs relacionados en el paquete actual.")
    else:
        show = rel[
            [
                "Rank",
                "Related_KPI_ID",
                "Related_KPI_Name",
                "Related_2026",
                "Related_Baseline_3Y",
                "Related_Performance_Change",
                "Related_Trend_Health",
                "Relation_Type",
            ]
        ].copy()

        show["Related_Performance_Change"] = show[
            "Related_Performance_Change"
        ].map(
            lambda x: f"{x:+.1%}" if pd.notna(x) else "N/D"
        )

        st.dataframe(
            show,
            use_container_width=True,
            hide_index=True,
        )

with tab2:
    st.markdown(f"**Purpose**  \n{row['Purpose']}")
    st.markdown(f"**Formula**  \n{row['Formula']}")
    st.markdown(f"**Direction**  \n{row['Direction']}")
    st.markdown(f"**Frequency**  \n{row['Frequency']}")
    st.markdown(f"**Underlying 2026 values**  \n{row['Underlying_Values_2026']}")

with tab3:
    st.markdown(f"**Data source**  \n{row['Data_Source']}")
    st.markdown(f"**Evidence**  \n{row['Evidence']}")
    st.markdown(f"**Collection method**  \n{row['Collection_Method']}")
    st.markdown(f"**Anti-gaming control**  \n{row['Anti_Gaming_Control']}")


# ============================================================
# CHAT
# ============================================================

st.markdown(
    """
<div class="chat-shell">
    <div class="chat-kicker">Ask the copilot</div>
    <div class="chat-title">¿Qué quieres entender ahora?</div>
    <div class="chat-copy">
        En este prototipo gratuito el asistente responde con reglas de gestión
        construidas sobre el KPI seleccionado y sus señales relacionadas.
    </div>
    <div class="quick-prompts">
        <span class="qp">¿Qué debería hacer?</span>
        <span class="qp">¿Qué KPI me preocupa?</span>
        <span class="qp">¿Qué está cambiando?</span>
        <span class="qp">¿Qué datos me faltan?</span>
        <span class="qp">¿Qué dice el forecast?</span>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

question = st.chat_input(
    "Pregúntale algo al Decision Copilot…"
)

if question:
    q = question.lower()

    with st.chat_message("user"):
        st.write(question)

    with st.chat_message("assistant"):
        if any(
            x in q
            for x in [
                "qué hago",
                "que hago",
                "acción",
                "accion",
                "recomienda",
                "debería hacer",
                "deberia hacer",
            ]
        ):
            st.write(analysis["action"])

        elif any(
            x in q
            for x in [
                "relacion",
                "otros kpi",
                "otros indicadores",
                "importan",
                "preocupa",
            ]
        ):
            if not analysis["deteriorating"].empty:
                st.write("Primero revisaría estas señales:")
                for _, r in analysis["deteriorating"].head(3).iterrows():
                    st.write(
                        f"• {r['Related_KPI_ID']} — "
                        f"{r['Related_KPI_Name']}: "
                        f"{float(r['Related_Performance_Change']):+.1%}"
                    )
            else:
                st.write(
                    "No detecto deterioros fuertes entre los KPIs relacionados "
                    "del paquete actual."
                )

        elif any(
            x in q
            for x in [
                "dato",
                "falta",
                "confianza",
                "confidence",
            ]
        ):
            st.write(f"Confianza: {analysis['confidence']}")
            for g in analysis["gaps"]:
                st.write("•", g)

        elif any(
            x in q
            for x in [
                "forecast",
                "2027",
                "proye",
                "futuro",
            ]
        ):
            st.write(
                f"La proyección estimada 2027 es "
                f"{fmt_value(forecast_value, row['Unit'])}. "
                f"Se construye con {forecast_method}. "
                "Es apoyo a decisión, no un resultado garantizado."
            )

        elif any(
            x in q
            for x in [
                "por qué",
                "porque",
                "diagn",
                "qué pasa",
                "que pasa",
                "cambio",
            ]
        ):
            st.write(analysis["diagnosis"])
            for e in analysis["evidence"]:
                st.write("•", e)

        else:
            st.write(
                "Puedo ayudarte a interpretar qué está pasando, qué acción "
                "priorizar, qué KPIs relacionados revisar, qué información falta "
                "o qué indica la proyección 2027."
            )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
<div class="brand-footer">
    <div class="brand-footer-big">Gente que cuida a la Gente.</div>
    <div class="brand-footer-small">
        FJV · Decision Copilot<br>
        Prototype · synthetic data · human review required
    </div>
</div>
""",
    unsafe_allow_html=True,
)
