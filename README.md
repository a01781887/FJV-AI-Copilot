# FJV KPI Decision Copilot — Free Prototype

Este prototipo NO usa una API de IA. Por eso no genera costos por consulta.

## Qué hace
- Lee `FJV_AI_Copilot_Context.xlsx`.
- Permite seleccionar cualquiera de los KPIs.
- Evalúa resultado 2026, baseline, momentum, target y dirección.
- Revisa KPIs relacionados.
- Genera un Decision Brief mediante reglas.
- Incluye un chat sencillo basado en intenciones.

## Importante
Los datos del archivo son sintéticos/dummy y las relaciones entre KPIs son heurísticas, no evidencia causal.

## Ejecutarlo en Mac

1. Abre Terminal.
2. Entra a esta carpeta, por ejemplo:
   cd ~/Downloads/FJV_AI_Copilot_FREE
3. Crea un entorno:
   python3 -m venv .venv
4. Actívalo:
   source .venv/bin/activate
5. Instala dependencias:
   pip install -r requirements.txt
6. Ejecuta:
   streamlit run app.py

Se abrirá en tu navegador.

## Siguiente fase
Cuando la lógica y UX estén aprobadas, se puede cambiar el motor de reglas por un modelo de IA y luego integrar el panel con Tableau.
