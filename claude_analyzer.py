"""
Claude AI analyzer module for medical audit reports.
Sends structured data to Claude API and receives formatted analysis.
"""

import anthropic
import os
from typing import Optional


SYSTEM_PROMPT = """Eres un experto en auditoría médica de calidad. Tu tarea es analizar datos de auditorías médicas y generar reportes estructurados siguiendo un formato estricto.

REGLAS FUNDAMENTALES:
1. NO mezcles componentes ni criterios entre sí.
2. El orden de los componentes SIEMPRE es: Anamnesis, Examen Físico, Diagnóstico, Productos de la Consulta.
3. Separa cada tipificación con su hallazgo e impacto sobre la atención de forma individual.
4. Cuenta exactamente cada ID de consulta (números de 5 a 10 dígitos).
5. Correlaciona el diagnóstico con cada hallazgo.
6. Al finalizar cada médico, no retenga información para el siguiente.
7. Los porcentajes del cuadro de cumplimiento deben ser exactamente los del archivo proporcionado, no los inventes.
8. Usa únicamente las tipificaciones exactas del documento de TIPIFICACIONES DE USO COMUN EN AUDITORIA MEDICA DE CALIDAD.
9. Clasifica No Conformidades y Eventos de Riesgo según el documento de Clasificación de no conformidades.

FORMATO DE SALIDA OBLIGATORIO - Sigue este formato exacto sin modificarlo:

---
ANÁLISIS CUANTITATIVO
[Tabla con columnas: ID de Consulta | Diagnóstico | No Conformidades | Eventos de Riesgo]

ANÁLISIS CUALITATIVO:
"Se ha realizado un análisis de un total de [N] auditorías. Se identificaron [X] No conformidades y [Y] Eventos de Riesgo, lo que destaca áreas de mejora significativas en la documentación y la práctica clínica.

Análisis de No Conformidades
Se identificaron [X] No Conformidades, las cuales afectan principalmente a los componentes de [componentes afectados]

1. [Componente]
[criterio individual afectado]
No conformidades Identificadas: [N]:
[Tipificación]
[Hallazgos]
[Impacto en la atención]

[siguiente criterio si aplica...]

2. [Componente]
[criterio individual afectado]
No conformidades Identificadas: [N]:
[Tipificación]
[Hallazgos]
[Impacto en la atención]

_______________________________________________________________________________

Análisis de Eventos de Riesgo
Se identificaron [Y] Eventos de Riesgo, las cuales si bien no son de gravedad crítica estas se convierten en oportunidades de mejora que se centran principalmente en los componentes de [componentes afectados]

1. [Componente]
[criterio individual afectado]
Evento de Riesgo Identificados: [N]:
[Tipificación]
[Hallazgos]
[Impacto en la atención]

[siguiente criterio si aplica...]"

RESUMEN EJECUTIVO
"Se evidencia un perfil de riesgo con afectación en los componentes de [top 3 componentes]. El criterio más afectado es [3 componentes con mayor cantidad de tipificaciones], debido a [hallazgo importante]. Se identificaron [X] No Conformidades en total; También se identificaron [Y] Eventos de Riesgo.

Componentes y número de hallazgos:
**ANAMNESIS**: se identifican [N] hallazgos; de los cuales [X] son No Conformidades y [Y] son Eventos de Riesgo.
**DIAGNÓSTICO**: se identifican [N] hallazgos; de los cuales [X] son No Conformidades y [Y] son Eventos de Riesgo.
**EXAMEN FÍSICO**: se identifican [N] hallazgos; de los cuales [X] son No Conformidades y [Y] son Eventos de Riesgo.
**PRODUCTOS DE LA CONSULTA**: se identifican [N] hallazgos; de los cuales [X] son No Conformidades y [Y] son Eventos de Riesgo."

[CUADRO DE CUMPLIMIENTO POR CRITERIO - con datos exactos del archivo de gráficas]

[COMENTARIO DE SEGUIMIENTO Y COMPARACIÓN DE PERIODOS - solo si hay 2 o más periodos]
---
"""


def build_analysis_prompt(
    doctor_name: str,
    doctor_code: str,
    period: str,
    consultations_table: str,
    compliance_table: str,
    clasificacion_table: str,
    tipificaciones_list: str,
    previous_period_data: str = "",
) -> str:
    """Build the full prompt to send to Claude for analysis."""
    prompt = f"""# SOLICITUD DE REPORTE DE AUDITORÍA MÉDICA

## MÉDICO: {doctor_name}
## CÓDIGO: {doctor_code}
## PERÍODO: {period}

---

## DATOS DE CONSULTAS AUDITADAS (Base de Datos x Cita):
{consultations_table}

---

## DATOS DE CUMPLIMIENTO POR CRITERIO (Gráficas de Cumplimiento):
{compliance_table}

---

## CLASIFICACIÓN DE NO CONFORMIDADES Y EVENTOS DE RIESGO:
{clasificacion_table}

---

## TIPIFICACIONES DE USO COMÚN EN AUDITORÍA MÉDICA:
{tipificaciones_list}

---
"""
    if previous_period_data:
        prompt += f"""
## DATOS DE PERÍODO ANTERIOR (para comparación):
{previous_period_data}

---
"""
    prompt += f"""
## INSTRUCCIONES:
1. Genera el reporte completo para el médico {doctor_name} ({doctor_code}) para el período {period}.
2. Sigue ESTRICTAMENTE el formato de salida especificado en el sistema.
3. Usa EXACTAMENTE las tipificaciones del documento de tipificaciones, tal como están escritas.
4. Clasifica cada hallazgo según el documento de clasificación de no conformidades.
5. Para el cuadro de cumplimiento, usa los porcentajes EXACTOS del archivo de gráficas proporcionado.
6. Si hay datos de período anterior, incluye el comentario de seguimiento comparando los últimos dos períodos.
7. Aplica las categorías de color en texto: ROJO (<85%), ANARANJADO (85%-94%), AMARILLO (95%-97%), VERDE (98%-100%).
8. Cuenta correctamente cada ID de consulta (números de 5 a 10 dígitos).
9. Correlaciona el diagnóstico con cada hallazgo encontrado.

GENERA EL REPORTE AHORA:
"""
    return prompt


def analyze_with_claude(
    doctor_name: str,
    doctor_code: str,
    period: str,
    consultations_table: str,
    compliance_table: str,
    clasificacion_table: str,
    tipificaciones_list: str,
    previous_period_data: str = "",
    api_key: str | None = None,
    model: str = "claude-opus-4-6",
    stream_callback=None,
) -> str:
    """
    Send data to Claude API and get the formatted audit report.

    Args:
        stream_callback: Optional callable(chunk: str) for streaming output
    Returns:
        Full report text
    """
    key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        raise ValueError("ANTHROPIC_API_KEY not set.")

    client = anthropic.Anthropic(api_key=key)

    prompt = build_analysis_prompt(
        doctor_name=doctor_name,
        doctor_code=doctor_code,
        period=period,
        consultations_table=consultations_table,
        compliance_table=compliance_table,
        clasificacion_table=clasificacion_table,
        tipificaciones_list=tipificaciones_list,
        previous_period_data=previous_period_data,
    )

    if stream_callback:
        full_text = ""
        with client.messages.stream(
            model=model,
            max_tokens=8192,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        ) as stream:
            for text_chunk in stream.text_stream:
                full_text += text_chunk
                stream_callback(text_chunk)
        return full_text
    else:
        response = client.messages.create(
            model=model,
            max_tokens=8192,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text


def parse_report_sections(report_text: str) -> dict:
    """
    Parse the Claude report into sections for document generation.
    Returns dict with keys: cuantitativo, cualitativo, no_conformidades,
    eventos_riesgo, resumen_ejecutivo, cumplimiento, seguimiento
    """
    sections = {
        "cuantitativo": "",
        "cualitativo": "",
        "no_conformidades": "",
        "eventos_riesgo": "",
        "resumen_ejecutivo": "",
        "cumplimiento": "",
        "seguimiento": "",
        "full_text": report_text,
    }

    lines = report_text.split("\n")
    current_section = None
    buffer = []

    def flush_buffer():
        nonlocal buffer
        text = "\n".join(buffer).strip()
        buffer = []
        return text

    for line in lines:
        line_stripped = line.strip()

        if "ANÁLISIS CUANTITATIVO" in line_stripped.upper():
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "cuantitativo"
        elif "ANÁLISIS CUALITATIVO" in line_stripped.upper():
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "cualitativo"
        elif "ANÁLISIS DE NO CONFORMIDADES" in line_stripped.upper():
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "no_conformidades"
        elif "ANÁLISIS DE EVENTOS DE RIESGO" in line_stripped.upper():
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "eventos_riesgo"
        elif "RESUMEN EJECUTIVO" in line_stripped.upper():
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "resumen_ejecutivo"
        elif "CUADRO DE CUMPLIMIENTO" in line_stripped.upper() or "REPORTE DE CUMPLIMIENTO" in line_stripped.upper():
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "cumplimiento"
        elif "COMENTARIO DE SEGUIMIENTO" in line_stripped.upper():
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "seguimiento"
        else:
            buffer.append(line)

    if current_section and buffer:
        sections[current_section] = flush_buffer()

    return sections
