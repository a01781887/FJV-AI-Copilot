
import re
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="FJV KPI Decision Copilot — Free Prototype",
    page_icon="🧭",
    layout="wide",
)

DATA_FILE = Path(__file__).parent / "FJV_AI_Copilot_Context.xlsx"

@st.cache_data
def load_data():
    kpis = pd.read_excel(DATA_FILE, sheet_name="01_KPI_Context")
    related = pd.read_excel(DATA_FILE, sheet_name="02_Related_KPIs")
    values = pd.read_excel(DATA_FILE, sheet_name="05_Value_Dictionary")
    return kpis, related, values

def fmt_value(value, unit):
    if pd.isna(value):
        return "N/D"
    unit = "" if pd.isna(unit) else str(unit)
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

def evaluate_target(row):
    target = "" if pd.isna(row["Target_2026"]) else str(row["Target_2026"])
    value = row["2026"]
    baseline = row["Baseline_3Y"]
    direction = str(row["Direction"])
    if pd.isna(value):
        return None, "No hay resultado 2026 suficiente para evaluar la meta."

    checks = []
    explanations = []

    # Absolute threshold based on inequality.
    le = re.search(r"(?:≤|<=)\s*(\d+(?:\.\d+)?)", target)
    ge = re.search(r"(?:≥|>=)\s*(\d+(?:\.\d+)?)", target)
    eq_pct = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%\s*", target)

    if le:
        lim = float(le.group(1))
        checks.append(float(value) <= lim)
        explanations.append(f"resultado {fmt_value(value, row['Unit'])} vs límite ≤{lim:g}")
    elif ge:
        lim = float(ge.group(1))
        # If target says reduction, don't use this as the absolute KPI threshold.
        if "reducción" not in target.lower() or direction == "↑":
            checks.append(float(value) >= lim)
            explanations.append(f"resultado {fmt_value(value, row['Unit'])} vs mínimo ≥{lim:g}")
    elif eq_pct:
        lim = float(eq_pct.group(1))
        checks.append(float(value) >= lim if direction == "↑" else float(value) <= lim)
        explanations.append(f"resultado {fmt_value(value, row['Unit'])} vs meta {lim:g}%")

    # Reduction-vs-baseline requirement such as ≥40% de reducción.
    red = re.search(r"(?:≥|>=)\s*(\d+(?:\.\d+)?)\s*%\s*de\s*reducci", target.lower())
    if red and not pd.isna(baseline) and float(baseline) != 0:
        required = float(red.group(1)) / 100.0
        actual = (float(baseline) - float(value)) / abs(float(baseline))
        checks.append(actual >= required)
        explanations.append(f"reducción vs baseline {actual:.1%} vs requerida {required:.1%}")

    if not checks:
        return None, f"Meta textual: {target or 'N/D'}"
    return all(checks), "; ".join(explanations)

def suspicious_units(row, value_dict):
    issues = []
    subset = value_dict[value_dict["KPI_ID"] == row["KPI_ID"]]
    if subset.empty:
        return issues
    for _, v in subset.iterrows():
        canonical = str(v.get("Canonical_Field", "")).lower()
        unit = str(v.get("Unit", "")).lower()
        role = str(v.get("Input_Role", ""))
        if ("count" in canonical or "request" in canonical) and ("día" in unit or "day" in unit):
            issues.append(
                f"{role}: '{v.get('Value_Name')}' está etiquetado como '{v.get('Unit')}', "
                "aunque el campo canónico sugiere que es un conteo."
            )
        if row.get("Formula_Type") == "RATIO_PCT" and role in ("NUM", "DEN") and "moneda" in unit:
            issues.append(
                f"{role}: '{v.get('Value_Name')}' aparece con unidad monetaria dentro de un ratio porcentual; conviene validarla."
            )
    return issues

def build_analysis(row, rel, value_dict):
    perf = float(row["Performance_Change"]) if not pd.isna(row["Performance_Change"]) else None
    target_met, target_detail = evaluate_target(row)

    deteriorating = rel[rel["Related_Performance_Change"] < 0].sort_values("Related_Performance_Change")
    improving = rel[rel["Related_Performance_Change"] >= 0].sort_values("Related_Performance_Change", ascending=False)

    if target_met is True and (perf is None or perf >= 0):
        diagnosis = (
            f"{row['KPI_ID']} está cumpliendo la meta 2026 y su momentum ajustado por dirección es positivo. "
            "No hay evidencia, con estos datos, de que requiera una corrección inmediata."
        )
    elif target_met is False and perf is not None and perf >= 0:
        diagnosis = (
            f"{row['KPI_ID']} está mejorando frente al baseline, pero todavía no cumple completamente la meta 2026. "
            "La prioridad es cerrar el gap restante sin perder las mejoras ya obtenidas."
        )
    elif perf is not None and perf < 0:
        diagnosis = (
            f"{row['KPI_ID']} se está deteriorando frente al baseline ajustado por dirección. "
            "Conviene investigar la desviación antes de escalar o replicar el proceso actual."
        )
    else:
        diagnosis = (
            f"{row['KPI_ID']} necesita revisión adicional porque la información disponible no permite clasificar con suficiente claridad "
            "el cumplimiento de meta y la tendencia."
        )

    evidence = [
        f"2026: {fmt_value(row['2026'], row['Unit'])}.",
        f"Baseline 3Y: {fmt_value(row['Baseline_3Y'], row['Unit'])}.",
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
                f"Mantener el proceso que sostiene el desempeño de {row['KPI_ID']} y validar su calidad con evidencia. "
                f"Después, investigar primero {deteriorating.iloc[0]['Related_KPI_ID']} porque es la señal relacionada con mayor deterioro. "
                "No atribuirle causalidad hasta revisar datos operativos."
            )
        else:
            weakest = rel.sort_values("Related_Performance_Change").head(1)
            if not weakest.empty:
                action = (
                    f"Mantener el desempeño actual y revisar periódicamente el control anti-gaming. "
                    f"Como siguiente foco, revisar {weakest.iloc[0]['Related_KPI_ID']} para encontrar el punto más débil del contexto relacionado."
                )
            else:
                action = "Mantener el desempeño actual y realizar controles periódicos de calidad y evidencia."
    elif target_met is False and perf is not None and perf >= 0:
        action = (
            f"Cerrar el gap restante a meta. Acción base del catálogo: {static_action} "
            "Desglosar el proceso o los inputs que componen el KPI para localizar dónde se concentra el tiempo, volumen o pérdida restante."
        )
    else:
        action = (
            f"Priorizar revisión de la desviación. Acción base del catálogo: {static_action} "
            "Comparar el KPI con sus señales relacionadas y validar primero los datos antes de atribuir una causa."
        )

    gaps = []
    unit_issues = suspicious_units(row, value_dict)
    gaps.extend(unit_issues)
    gaps.append(
        "El paquete contiene resultados agregados; para explicar causas reales hacen falta datos operativos más granulares "
        "(por etapa, segmento, causa, responsable o periodo, según el KPI)."
    )
    gaps.append(
        "Las relaciones entre KPIs son heurísticas de recuperación; no deben interpretarse como causalidad."
    )

    confidence = "Moderada"
    if target_met is not None and perf is not None and not unit_issues:
        confidence = "Moderada-alta para describir desempeño; moderada para explicar causas."
    elif unit_issues:
        confidence = "Moderada-baja hasta validar las unidades señaladas."

    return {
        "diagnosis": diagnosis,
        "evidence": evidence,
        "related": related_signal,
        "action": action,
        "gaps": gaps,
        "confidence": confidence,
        "target_met": target_met,
    }

kpis, related, value_dict = load_data()

st.title("🧭 FJV KPI Decision Copilot")
st.caption("Prototipo gratuito basado en reglas — sin API y sin costo de IA.")

st.info(
    "Los datos actuales son sintéticos/dummy. Este prototipo no usa IA generativa: "
    "usa reglas de decisión y el contexto del catálogo para validar la experiencia antes de pagar una API."
)

labels = {
    f"{row.KPI_ID} — {row.KPI_Name}": row.KPI_ID
    for row in kpis.itertuples()
}
selected_label = st.selectbox("Selecciona un KPI", list(labels.keys()))
selected_id = labels[selected_label]

row = kpis[kpis["KPI_ID"] == selected_id].iloc[0]
rel = related[related["Focal_KPI_ID"] == selected_id].sort_values("Rank")
analysis = build_analysis(row, rel, value_dict)

st.subheader(row["KPI_Name"])
st.caption(f"{row['Department']} · {row['Strategic_Objective']} · Owner: {row['Owner']}")

c1, c2, c3, c4 = st.columns(4)
c1.metric("2026 Result", fmt_value(row["2026"], row["Unit"]))
c2.metric("Target", str(row["Target_2026"]))
c3.metric("3Y Baseline", fmt_value(row["Baseline_3Y"], row["Unit"]))
momentum = row["Performance_Change"]
c4.metric("Momentum", f"{float(momentum):+.1%}" if not pd.isna(momentum) else "N/D")

years = pd.DataFrame({
    "Year": [2023, 2024, 2025, 2026],
    "Value": [row["2023"], row["2024"], row["2025"], row["2026"]],
}).set_index("Year")
st.line_chart(years)

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
            "Rank", "Related_KPI_ID", "Related_KPI_Name",
            "Related_2026", "Related_Baseline_3Y",
            "Related_Performance_Change", "Related_Trend_Health",
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
st.markdown("### 💬 Ask the free copilot")
st.caption("Esta versión responde con reglas, no con un modelo generativo.")

question = st.chat_input("Ej. ¿Qué debería hacer? / ¿Qué KPIs relacionados importan? / ¿Qué datos faltan?")
if question:
    q = question.lower()
    with st.chat_message("user"):
        st.write(question)
    with st.chat_message("assistant"):
        if any(x in q for x in ["qué hago", "que hago", "acción", "accion", "recomienda", "debería hacer", "deberia hacer"]):
            st.write(analysis["action"])
        elif any(x in q for x in ["relacion", "otros kpi", "otros indicadores", "importan"]):
            for r in analysis["related"]:
                st.write("•", r)
        elif any(x in q for x in ["dato", "falta", "confianza", "confidence"]):
            st.write(f"Confianza: {analysis['confidence']}")
            for g in analysis["gaps"]:
                st.write("•", g)
        elif any(x in q for x in ["por qué", "porque", "diagn", "qué pasa", "que pasa", "cambio"]):
            st.write(analysis["diagnosis"])
            for e in analysis["evidence"]:
                st.write("•", e)
        else:
            st.write(
                "En esta versión gratuita puedo responder cuatro tipos de preguntas: "
                "qué está pasando, qué acción tomar, qué KPIs relacionados importan y qué datos faltan."
            )
