import re
from pathlib import Path

import pandas as pd
import streamlit as st


# =========================
# PAGE
# =========================
st.set_page_config(
    page_title="FJV KPI Decision Copilot",
    page_icon="🧭",
    layout="wide",
)

st.markdown("""
<style>
:root {
    --purple:#6D43A2;
    --orange:#F18A3A;
    --teal:#173F46;
    --sky:#8FCBE6;
    --yellow:#F3D777;
    --cream:#F7F2E9;
    --paper:#FFFDF9;
    --ink:#263438;
    --muted:#758185;
    --border:#E6DFD5;
}

[data-testid="stAppViewContainer"]{
    background:
      radial-gradient(circle at 100% 0%,rgba(143,203,230,.14),transparent 20%),
      var(--cream);
}
[data-testid="stHeader"]{
    background:rgba(247,242,233,.85);
    backdrop-filter:blur(10px);
}
.block-container{
    max-width:1320px;
    padding-top:1.5rem;
    padding-bottom:4rem;
}
#MainMenu, footer{visibility:hidden;}

html,body,[class*="css"]{
    font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
}
h1,h2,h3{color:var(--ink);letter-spacing:-.03em;}

.fjv-header{
    background:var(--teal);
    padding:28px 32px;
    border-radius:22px;
    margin-bottom:18px;
    position:relative;
    overflow:hidden;
    box-shadow:0 16px 40px rgba(23,63,70,.12);
}
.fjv-header:after{
    content:"";
    position:absolute;
    width:150px;height:150px;
    right:-55px;top:-65px;
    border-radius:50%;
    background:var(--purple);
}
.fjv-header:before{
    content:"";
    position:absolute;
    width:95px;height:95px;
    right:55px;bottom:-55px;
    border-radius:50%;
    background:var(--orange);
}
.fjv-kicker{
    position:relative;z-index:2;
    color:var(--yellow);
    font-size:11px;font-weight:800;
    letter-spacing:.15em;text-transform:uppercase;
    margin-bottom:8px;
}
.fjv-title{
    position:relative;z-index:2;
    color:white;
    font-size:38px;font-weight:760;
    line-height:1.05;letter-spacing:-.045em;
}
.fjv-sub{
    position:relative;z-index:2;
    color:rgba(255,255,255,.72);
    font-size:14px;margin-top:9px;
}

div[data-baseweb="select"]>div{
    background:var(--paper)!important;
    border:1px solid var(--border)!important;
    border-radius:14px!important;
    min-height:52px;
}

.kpi-box{
    background:var(--paper);
    border:1px solid var(--border);
    border-radius:18px;
    padding:18px 20px;
    margin:12px 0;
    box-shadow:0 8px 24px rgba(30,40,42,.035);
}
.kpi-id{
    color:var(--orange);
    font-size:10px;font-weight:800;
    letter-spacing:.12em;text-transform:uppercase;
}
.kpi-name{
    color:var(--ink);
    font-size:25px;font-weight:760;
    letter-spacing:-.035em;
    line-height:1.12;margin-top:5px;
}
.kpi-meta{
    color:var(--muted);
    font-size:12px;margin-top:8px;
}

[data-testid="stMetric"]{
    background:var(--paper);
    border:1px solid var(--border);
    padding:18px;
    border-radius:18px;
    min-height:124px;
    box-shadow:0 6px 20px rgba(30,40,42,.04);
}
[data-testid="stMetricLabel"]{
    color:var(--muted);
    font-size:.76rem;
    font-weight:700;
}
[data-testid="stMetricValue"]{
    color:var(--teal);
    font-weight:780;
    letter-spacing:-.04em;
}

.section-kicker{
    color:var(--purple);
    font-size:10px;font-weight:800;
    letter-spacing:.13em;text-transform:uppercase;
    margin-top:24px;margin-bottom:3px;
}
.section-title{
    color:var(--ink);
    font-size:22px;font-weight:760;
    letter-spacing:-.03em;margin-bottom:8px;
}

[data-testid="stAlert"]{border-radius:15px;}
button[data-baseweb="tab"]{font-weight:700;font-size:.9rem;}
div[data-baseweb="tab-highlight"]{background:var(--purple)!important;}

[data-testid="stDataFrame"]{
    background:var(--paper);
    border:1px solid var(--border);
    border-radius:16px;
    overflow:hidden;
}

[data-testid="stChatInput"]{border-radius:16px;}
[data-testid="stChatMessage"]{
    background:var(--paper);
    border:1px solid var(--border);
    border-radius:16px;
    padding:8px 14px;
    margin-bottom:8px;
}
hr{border-color:var(--border);}
p{color:#536267;}

.fjv-footer{
    margin-top:30px;
    padding-top:18px;
    border-top:1px solid var(--border);
    display:flex;
    justify-content:space-between;
    gap:20px;
    flex-wrap:wrap;
}
.fjv-footer-left{color:var(--teal);font-weight:700;font-size:12px;}
.fjv-footer-right{color:var(--muted);font-size:10px;}
</style>
""", unsafe_allow_html=True)

DATA_FILE = Path(__file__).parent / "FJV_AI_Copilot_Context.xlsx"


# =========================
# DATA
# =========================
@st.cache_data
def load_data():
    kpis = pd.read_excel(DATA_FILE, sheet_name="01_KPI_Context")
    related = pd.read_excel(DATA_FILE, sheet_name="02_Related_KPIs")
    values = pd.read_excel(DATA_FILE, sheet_name="05_Value_Dictionary")
    return kpis, related, values


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


# =========================
# TARGET
# =========================
def evaluate_target(row):
    target = "" if pd.isna(row["Target_2026"]) else str(row["Target_2026"])
    value = row["2026"]
    baseline = row["Baseline_3Y"]
    direction = str(row["Direction"])

    if pd.isna(value):
        return None, "No hay resultado 2026 suficiente para evaluar la meta."

    checks, explanations = [], []

    le = re.search(r"(?:≤|<=)\s*(\d+(?:\.\d+)?)", target)
    ge = re.search(r"(?:≥|>=)\s*(\d+(?:\.\d+)?)", target)
    eq_pct = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%\s*", target)

    if le:
        lim = float(le.group(1))
        checks.append(float(value) <= lim)
        explanations.append(f"resultado {fmt_value(value,row['Unit'])} vs límite ≤{lim:g}")
    elif ge:
        lim = float(ge.group(1))
        if "reducción" not in target.lower() or direction == "↑":
            checks.append(float(value) >= lim)
            explanations.append(f"resultado {fmt_value(value,row['Unit'])} vs mínimo ≥{lim:g}")
    elif eq_pct:
        lim = float(eq_pct.group(1))
        checks.append(float(value) >= lim if direction == "↑" else float(value) <= lim)
        explanations.append(f"resultado {fmt_value(value,row['Unit'])} vs meta {lim:g}%")

    red = re.search(r"(?:≥|>=)\s*(\d+(?:\.\d+)?)\s*%\s*de\s*reducci", target.lower())
    if red and not pd.isna(baseline) and float(baseline) != 0:
        required = float(red.group(1)) / 100
        actual = (float(baseline) - float(value)) / abs(float(baseline))
        checks.append(actual >= required)
        explanations.append(f"reducción vs baseline {actual:.1%} vs requerida {required:.1%}")

    if not checks:
        return None, f"Meta textual: {target or 'N/D'}"

    return all(checks), "; ".join(explanations)


# =========================
# DATA QUALITY
# =========================
def suspicious_units(row, value_dict):
    issues = []
    subset = value_dict[value_dict["KPI_ID"] == row["KPI_ID"]]

    for _, v in subset.iterrows():
        canonical = str(v.get("Canonical_Field", "")).lower()
        unit = str(v.get("Unit", "")).lower()
        role = str(v.get("Input_Role", ""))

        if ("count" in canonical or "request" in canonical) and ("día" in unit or "day" in unit):
            issues.append(
                f"{role}: '{v.get('Value_Name')}' está etiquetado como "
                f"'{v.get('Unit')}', aunque parece ser un conteo."
            )

        if row.get("Formula_Type") == "RATIO_PCT" and role in ("NUM", "DEN") and "moneda" in unit:
            issues.append(
                f"{role}: '{v.get('Value_Name')}' aparece con unidad monetaria "
                "dentro de un ratio porcentual; conviene validarla."
            )
    return issues


# =========================
# FORECAST
# =========================
def forecast_2027(row, rel):
    values = [row["2023"], row["2024"], row["2025"], row["2026"]]
    years = [2023, 2024, 2025, 2026]

    if any(pd.isna(v) for v in values):
        return None, "Not enough historical data"

    x_mean = sum(years) / len(years)
    y_mean = sum(values) / len(values)
    numerator = sum((x-x_mean)*(y-y_mean) for x,y in zip(years,values))
    denominator = sum((x-x_mean)**2 for x in years)
    slope = numerator / denominator if denominator else 0
    forecast = y_mean + slope * (2027-x_mean)

    if not rel.empty:
        related_momentum = rel["Related_Performance_Change"].dropna()
        if len(related_momentum):
            avg_related = related_momentum.mean()
            adjustment = max(min(avg_related * .20, .10), -.10)

            if str(row["Direction"]) == "↑":
                forecast *= (1 + adjustment)
            elif str(row["Direction"]) == "↓":
                forecast *= (1 - adjustment)

    unit = str(row["Unit"]).lower()

    if "%" in unit:
        forecast = max(0, min(100, forecast))
    elif any(x in unit for x in ["día", "day", "#", "proyecto", "alianza"]):
        forecast = max(0, forecast)

    return forecast, "Trend 2023–2026 + related KPI signal"


# =========================
# DECISION LOGIC
# =========================
def build_analysis(row, rel, value_dict):
    perf = float(row["Performance_Change"]) if not pd.isna(row["Performance_Change"]) else None
    target_met, target_detail = evaluate_target(row)

    deteriorating = rel[rel["Related_Performance_Change"] < 0].sort_values("Related_Performance_Change")
    improving = rel[rel["Related_Performance_Change"] >= 0].sort_values(
        "Related_Performance_Change", ascending=False
    )

    if target_met is True and (perf is None or perf >= 0):
        diagnosis = (
            f"{row['KPI_ID']} está cumpliendo la meta 2026 y su momentum ajustado por "
            "dirección es positivo. No hay evidencia, con estos datos, de que requiera "
            "una corrección inmediata."
        )
    elif target_met is False and perf is not None and perf >= 0:
        diagnosis = (
            f"{row['KPI_ID']} está mejorando frente al baseline, pero todavía no cumple "
            "completamente la meta 2026. La prioridad es cerrar el gap restante sin perder "
            "las mejoras ya obtenidas."
        )
    elif perf is not None and perf < 0:
        diagnosis = (
            f"{row['KPI_ID']} se está deteriorando frente al baseline ajustado por dirección. "
            "Conviene investigar la desviación antes de escalar o replicar el proceso actual."
        )
    else:
        diagnosis = (
            f"{row['KPI_ID']} necesita revisión adicional porque la información disponible "
            "no permite clasificar con suficiente claridad el cumplimiento de meta y la tendencia."
        )

    evidence = [
        f"2026: {fmt_value(row['2026'],row['Unit'])}.",
        f"Baseline 3Y: {fmt_value(row['Baseline_3Y'],row['Unit'])}.",
        f"Target 2026: {row['Target_2026']}.",
        f"Momentum ajustado por dirección: {perf:+.1%}." if perf is not None else "Momentum: N/D.",
        f"Evaluación de meta: {target_detail}.",
    ]

    related_signal = []

    if not deteriorating.empty:
        names = ", ".join(
            f"{r['Related_KPI_ID']} ({float(r['Related_Performance_Change']):+.1%})"
            for _, r in deteriorating.head(3).iterrows()
        )
        related_signal.append(
            f"Hay KPIs relacionados con deterioro: {names}. Esto merece atención, pero no prueba causalidad."
        )

    if not improving.empty:
        names = ", ".join(
            f"{r['Related_KPI_ID']} ({float(r['Related_Performance_Change']):+.1%})"
            for _, r in improving.head(3).iterrows()
        )
        related_signal.append(f"Las señales relacionadas más positivas son: {names}.")

    if not related_signal:
        related_signal.append("No se detectaron señales relacionadas suficientes en el paquete actual.")

    static_action = str(row.get("Static_Required_Action", "")).strip()

    if target_met is True:
        if not deteriorating.empty:
            action = (
                f"Mantener el proceso que sostiene el desempeño de {row['KPI_ID']} y validar "
                f"su calidad con evidencia. Después, investigar primero "
                f"{deteriorating.iloc[0]['Related_KPI_ID']} porque es la señal relacionada con "
                "mayor deterioro. No atribuirle causalidad hasta revisar datos operativos."
            )
        else:
            weakest = rel.sort_values("Related_Performance_Change").head(1)
            if not weakest.empty:
                action = (
                    "Mantener el desempeño actual y revisar periódicamente el control anti-gaming. "
                    f"Como siguiente foco, revisar {weakest.iloc[0]['Related_KPI_ID']} para encontrar "
                    "el punto más débil del contexto relacionado."
                )
            else:
                action = "Mantener el desempeño actual y realizar controles periódicos de calidad y evidencia."

    elif target_met is False and perf is not None and perf >= 0:
        action = (
            f"Cerrar el gap restante a meta. Acción base del catálogo: {static_action} "
            "Desglosar el proceso o los inputs que componen el KPI para localizar dónde "
            "se concentra el tiempo, volumen o pérdida restante."
        )
    else:
        if not deteriorating.empty:
            action = (
                f"Priorizar la revisión de la desviación y contrastarla con "
                f"{deteriorating.iloc[0]['Related_KPI_ID']}, la señal relacionada más débil. "
                "Validar primero los datos antes de atribuir una causa."
            )
        else:
            action = (
                f"Priorizar revisión de la desviación. Acción base del catálogo: {static_action} "
                "Comparar el KPI con sus señales relacionadas y validar primero los datos antes "
                "de atribuir una causa."
            )

    unit_issues = suspicious_units(row, value_dict)
    gaps = list(unit_issues)
    gaps.append(
        "El paquete contiene resultados agregados; para explicar causas reales hacen falta "
        "datos operativos más granulares por etapa, segmento, causa, responsable o periodo."
    )
    gaps.append(
        "Las relaciones entre KPIs son heurísticas de recuperación; no deben interpretarse como causalidad."
    )

    if target_met is not None and perf is not None and not unit_issues:
        confidence = "Moderada-alta para describir desempeño; moderada para explicar causas."
    elif unit_issues:
        confidence = "Moderada-baja hasta validar las unidades señaladas."
    else:
        confidence = "Moderada"

    return {
        "diagnosis": diagnosis,
        "evidence": evidence,
        "related": related_signal,
        "action": action,
        "gaps": gaps,
        "confidence": confidence,
        "target_met": target_met,
    }


# =========================
# APP
# =========================
kpis, related, value_dict = load_data()

st.markdown("""
<div class="fjv-header">
    <div class="fjv-kicker">Fundación Jorge Vergara · Impact Intelligence</div>
    <div class="fjv-title">FJV Decision Copilot</div>
    <div class="fjv-sub">KPI analysis · related signals · management recommendations · 2027 forecast</div>
</div>
""", unsafe_allow_html=True)

st.info(
    "Los datos actuales son sintéticos/dummy. Este prototipo utiliza reglas de decisión "
    "y relaciones entre KPIs para apoyar la interpretación gerencial."
)

labels = {
    f"{r.KPI_ID} — {r.KPI_Name}": r.KPI_ID
    for r in kpis.itertuples()
}

selected_label = st.selectbox("Selecciona un KPI", list(labels.keys()))
selected_id = labels[selected_label]

row = kpis[kpis["KPI_ID"] == selected_id].iloc[0]
rel = related[related["Focal_KPI_ID"] == selected_id].sort_values("Rank")

analysis = build_analysis(row, rel, value_dict)
forecast_2027_value, forecast_method = forecast_2027(row, rel)

st.markdown(f"""
<div class="kpi-box">
    <div class="kpi-id">{row['KPI_ID']} · {row['Department']}</div>
    <div class="kpi-name">{row['KPI_Name']}</div>
    <div class="kpi-meta">{row['Strategic_Objective']} · Owner: {row['Owner']}</div>
</div>
""", unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("2026 Result", fmt_value(row["2026"], row["Unit"]))
c2.metric("Target", str(row["Target_2026"]))
c3.metric("3Y Baseline", fmt_value(row["Baseline_3Y"], row["Unit"]))

momentum = row["Performance_Change"]
c4.metric(
    "Momentum",
    f"{float(momentum):+.1%}" if not pd.isna(momentum) else "N/D"
)

c5.metric(
    "2027 Forecast",
    fmt_value(forecast_2027_value, row["Unit"])
    if forecast_2027_value is not None
    else "N/D"
)

st.markdown("""
<div class="section-kicker">Performance trend</div>
<div class="section-title">Histórico 2023–2026 + proyección 2027</div>
""", unsafe_allow_html=True)

trend_data = pd.DataFrame({
    "Year": [2023, 2024, 2025, 2026, 2027],
    "Value": [
        row["2023"],
        row["2024"],
        row["2025"],
        row["2026"],
        forecast_2027_value
    ]
}).set_index("Year")

st.line_chart(trend_data)

st.caption(
    f"2027 Estimated Forecast — {forecast_method}. "
    "Projection for decision support; not an actual result."
)

tab1, tab2, tab3 = st.tabs(["Decision Brief", "Related KPIs", "KPI Context"])

with tab1:
    st.markdown("### Diagnosis")
    st.write(analysis["diagnosis"])

    st.markdown("### Evidence")
    for e in analysis["evidence"]:
        st.write("•", e)

    st.markdown("### Related KPI Signals")
    for r in analysis["related"]:
        st.write("•", r)

    st.markdown("### Recommended Action")
    st.success(analysis["action"])

    st.markdown("### Confidence & Data Gaps")
    st.write(f"**Confidence:** {analysis['confidence']}")
    for g in analysis["gaps"]:
        st.write("•", g)

with tab2:
    if rel.empty:
        st.write("No hay KPIs relacionados en el paquete actual.")
    else:
        show = rel[[
            "Rank",
            "Related_KPI_ID",
            "Related_KPI_Name",
            "Related_2026",
            "Related_Baseline_3Y",
            "Related_Performance_Change",
            "Related_Trend_Health",
            "Relation_Type"
        ]].copy()

        show["Related_Performance_Change"] = show["Related_Performance_Change"].map(
            lambda x: f"{x:+.1%}" if pd.notna(x) else "N/D"
        )

        st.dataframe(show, use_container_width=True, hide_index=True)

with tab3:
    st.write("**Purpose:**", row["Purpose"])
    st.write("**Formula:**", row["Formula"])
    st.write("**Direction:**", row["Direction"])
    st.write("**Frequency:**", row["Frequency"])
    st.write("**Data Source:**", row["Data_Source"])
    st.write("**Evidence:**", row["Evidence"])
    st.write("**Underlying 2026 values:**", row["Underlying_Values_2026"])
    st.write("**Anti-gaming control:**", row["Anti_Gaming_Control"])

st.divider()
st.markdown("### 💬 Ask the Decision Copilot")
st.caption(
    "Esta versión gratuita responde con reglas de gestión basadas en el KPI seleccionado "
    "y sus señales relacionadas."
)

question = st.chat_input(
    "Ej. ¿Qué debería hacer? / ¿Qué KPIs relacionados importan? / ¿Qué datos faltan?"
)

if question:
    q = question.lower()

    with st.chat_message("user"):
        st.write(question)

    with st.chat_message("assistant"):
        if any(x in q for x in [
            "qué hago","que hago","acción","accion","recomienda",
            "debería hacer","deberia hacer"
        ]):
            st.write(analysis["action"])

        elif any(x in q for x in [
            "relacion","otros kpi","otros indicadores","importan"
        ]):
            for r in analysis["related"]:
                st.write("•", r)

        elif any(x in q for x in [
            "dato","falta","confianza","confidence"
        ]):
            st.write(f"Confianza: {analysis['confidence']}")
            for g in analysis["gaps"]:
                st.write("•", g)

        elif any(x in q for x in [
            "forecast","2027","proyección","proyeccion","futuro"
        ]):
            st.write(
                f"La proyección estimada para 2027 es "
                f"{fmt_value(forecast_2027_value,row['Unit'])}. "
                f"Se construye con {forecast_method}. "
                "Es una estimación para apoyar decisiones, no un resultado garantizado."
            )

        elif any(x in q for x in [
            "por qué","porque","diagn","qué pasa","que pasa","cambio"
        ]):
            st.write(analysis["diagnosis"])
            for e in analysis["evidence"]:
                st.write("•", e)

        else:
            st.write(
                "En esta versión gratuita puedo ayudarte a interpretar qué está pasando, "
                "qué acción tomar, qué KPIs relacionados revisar, qué datos faltan y "
                "qué muestra la proyección 2027."
            )

st.markdown("""
<div class="fjv-footer">
    <div class="fjv-footer-left">Fundación Jorge Vergara · Decision Copilot</div>
    <div class="fjv-footer-right">Prototype · synthetic data · human review required</div>
</div>
""", unsafe_allow_html=True)
