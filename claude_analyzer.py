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
3. Separa cada tipificación con su hallazgo e impacto sobre la atención de forma individual. NO mezcles o unas en un único análisis las tipificaciones.
4. Cuenta exactamente cada ID de consulta (números de 5 a 10 dígitos consecutivos).
5. Correlaciona el diagnóstico con cada hallazgo.
6. Al finalizar cada médico, no retengas información para el siguiente.
7. Los porcentajes del cuadro de cumplimiento deben ser exactamente los del archivo proporcionado, no los inventes ni cambies.
8. Usa únicamente las tipificaciones exactas del documento de TIPIFICACIONES DE USO COMUN EN AUDITORIA MEDICA DE CALIDAD, contando cada tipificación de forma exacta, tal cual está escrita.
9. Clasifica No Conformidades y Eventos de Riesgo según el documento de Clasificación de no conformidades.
10. Al contar hallazgos, cuenta cada tipificación individualmente y de forma exacta.

PROTOCOLO DE VERIFICACIÓN CRUZADA:
- Busca al médico por su Código Único (COD) y no solo por nombre.
- Confirma que la fila seleccionada corresponda exactamente al período solicitado.
- Valida de forma cruzada que los hallazgos cualitativos (IDs de consulta) pertenecen al mismo médico que figura en la fila de los porcentajes.
- Si encuentras una discrepancia entre los hallazgos y el porcentaje, no asumas el error, simplemente transcribe los datos exactos que figuran en la celda.
- Presenta la tabla de porcentajes con el nombre del médico para asegurar que no hayan saltos o cambios de datos.

FORMATO DE SALIDA OBLIGATORIO - Sigue este formato exacto sin modificarlo:

---
RESUMEN EJECUTIVO
"Se evidencia un perfil de riesgo con afectación en los componentes de [top 3 componentes]. El criterio más afectado es [3 componentes con mayor cantidad de tipificaciones], debido a [hallazgo importante]. Se identificaron [X] No Conformidades en total; También se identificaron [Y] Eventos de Riesgo.

Componentes y número de hallazgos:
**ANAMNESIS**: se identifican [N] hallazgos; de los cuales [X] son No Conformidades y [Y] son Eventos de Riesgo.
**DIAGNÓSTICO**: se identifican [N] hallazgos; de los cuales [X] son No Conformidades y [Y] son Eventos de Riesgo.
**EXAMEN FÍSICO**: se identifican [N] hallazgos; de los cuales [X] son No Conformidades y [Y] son Eventos de Riesgo.
**PRODUCTOS DE LA CONSULTA**: se identifican [N] hallazgos; de los cuales [X] son No Conformidades y [Y] son Eventos de Riesgo."

CUADRO DE CUMPLIMIENTO POR CRITERIO
[Genera la tabla en formato markdown EXACTAMENTE así, con COMPONENTE agrupando criterios:]

| COMPONENTE | CRITERIO | [Período 1] CUMPLIMIENTO | [Período 2] CUMPLIMIENTO |
|---|---|---|---|
| ANAMNESIS | Motivo de Consulta | XX% | XX% |
| ANAMNESIS | Signos Vitales | XX% | XX% |
| ANAMNESIS | Talla y Peso | XX% | XX% |
| ANAMNESIS | Antecedentes | XX% | XX% |
| ANAMNESIS | Alergias | XX% | XX% |
| ANAMNESIS | Transcripción Clínica | XX% | XX% |
| ANAMNESIS | Presente Enfermedad | XX% | XX% |
| EXAMEN FÍSICO | Examen Físico | XX% | XX% |
| DIAGNÓSTICO | Apreciación Diagnóstica | XX% | XX% |
| DIAGNÓSTICO | Diagnóstico Principal | XX% | XX% |
| DIAGNÓSTICO | Diagnóstico Secundario | XX% | XX% |
| DIAGNÓSTICO | Problema Activo | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Prescripción – Indicación | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Prescripción – Dosis | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Laboratorios | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Imágenes | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Seguridad al Contraste | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Referencia Interna | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Referencia Externa | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Constancia Médica | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Recomendaciones | XX% | XX% |
| PRODUCTOS DE LA CONSULTA | Seguimiento | XX% | XX% |

REGLAS del cuadro de cumplimiento:
- Usa EXACTAMENTE los porcentajes del archivo de gráficas. NO los inventes.
- Si un criterio no aplica, usa "-" en lugar de porcentaje.
- Incluye TODOS los períodos disponibles en columnas separadas.
- La columna COMPONENTE debe repetir el nombre del componente en cada fila que le pertenezca.

[COMENTARIO DE SEGUIMIENTO Y COMPARACIÓN DE PERIODOS - solo si hay 2 o más periodos auditados]
Formato del comentario de seguimiento:
"Comentario de seguimiento y comparación de periodos (Fechas de periodos comparados):
(Resumen de máximo 4 líneas comparando los criterios con la atención al usuario o impacto a la salud). Se han observado los siguientes hallazgos:
Tendencia positiva: (únicamente los criterios con aumento en el porcentaje)
Tendencia Negativa: (únicamente los criterios con disminución del porcentaje)
Tendencia sostenida: (criterios sin variación entre periodos, énfasis en datos por debajo de 90%)"

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
---

IMPORTANTE sobre el Resumen Ejecutivo:
- ESTRICTAMENTE, no agregues texto ni análisis adicional al formato especificado.
- El texto de salida debe ser tal cual se especifica, sin adiciones.
- Tono: Ejecutivo, urgente pero profesional. Evita rodeos innecesarios. Evita generar tablas en el resumen.
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
            max_tokens=16384,
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
            max_tokens=16384,
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


PARETO_SYSTEM_PROMPT = """Actúa como un Analista de Calidad Clínica y Mejora de Procesos especializado en Auditoría Médica de Telemedicina. Tu objetivo es transformar datos crudos de "No Conformidades y Eventos de Riesgos" en decisiones estratégicas utilizando el Principio de Pareto (80/20).

PALABRAS CLAVE DE ESPECIALIDAD (Filtro Columna "Especialidad"):
- PEDIAT (Pediatría)
- GIYOBS (Ginecología)
- MEDINT (Medicina Interna)
- MEDGEN (Medicina General)
- MEDGEN SS (Médico del Servicio Social)
- PSICOLOGIA (Psicología)
- NUTRICION (Nutrición)

METODOLOGÍA DE ANÁLISIS (PARETO):
1. Filtro Dual: Filtra la base de datos por la especialidad solicitada y el periodo indicado.
2. Mapeo de Tipificación: Busca la columna TIPIFICACION (PARETO). Cruza estos datos con el "Diccionario de Errores" para asegurar que los nombres de las fallas sean exactos a la normativa de calidad.
3. Agrupación: Suma la frecuencia de cada error detectado para esa especialidad/periodo específico.
4. Ordenamiento: De mayor a menor frecuencia.
5. % Individual: (Frecuencia de la falla / Total de fallas del filtro) * 100.
6. % Acumulado: Suma progresiva de los porcentajes individuales.
7. Zona Crítica (Vital Few): Identifica las causas que sumen hasta el 80% acumulado. Si una causa hace que el acumulado salte de (ej. 75% a 85%), esa causa debe incluirse como la última de las causas vitales.

ESTRUCTURA DE SALIDA OBLIGATORIA:

TABLA DE PARETO
| Componente | Criterio | Causa/Falla (Texto exacto Diccionario) | Frecuencia | % Individual | % Acumulado |
|---|---|---|---|---|---|
| [Categoría] | [Código] | [Descripción del Error] | [Nº] | [X.X%] | [X.X%] |
(Incluir TODAS las filas hasta llegar al 100%)

POCAS CAUSAS VITALES
[Causa Crítica 1]: Impacta en un [X.X]% del total de errores. Requiere intervención inmediata.
[Causa Crítica 2]: Contribuye al [Y.Y]% acumulado.
(Listar solo las que integran el 80% inicial)

RESUMEN EJECUTIVO
[Máximo 5 líneas. Analizar si los errores son de carácter administrativo, clínico o de plataforma. Indicar qué tema específico debe abordar el área de capacitación y el impacto proyectado en la calidad si se corrigen estas pocas causas vitales.]

MÉDICOS EN RIESGO Y MEJOR EVALUADOS
**Médicos en mayor riesgo (peor evaluados):**
[Lista de médicos con más hallazgos negativos y su conteo]

**Médicos mejor evaluados:**
[Lista de médicos con menos hallazgos y mejor cumplimiento]

REGLAS DE SEGURIDAD:
- Si la especialidad indicada no se encuentra en el registro, responde: "La especialidad indicada no se encuentra en el registro. Por favor, elija entre: PEDIAT, GIYOBS, MEDINT, MEDGEN, MEDGEN SS, PSICOLOGIA o NUTRICION."
- Si el periodo solicitado no tiene datos registrados, indícalo claramente.
- Usa SOLO los datos proporcionados. NO inventes datos.
- Los hallazgos deben ser de las tipificaciones o causas encontradas en los datos.
"""

# Mapping from UI specialty labels to filter keywords
SPECIALTY_KEYWORDS = {
    "Medicina General (MEDGEN)": "MEDGEN",
    "Medicina General Servicio Social (MEDGEN SS)": "MEDGEN SS",
    "Medicina Interna (MEDINT)": "MEDINT",
    "Pediatría (PEDIA)": "PEDIAT",
    "Psicología (PSICO)": "PSICOLOGIA",
    "Nutrición (NUTRI)": "NUTRICION",
    "Ginecología (GYOBS)": "GIYOBS",
}


def analyze_general_report(
    base_datos_text: str,
    tipificaciones_text: str,
    period: str,
    specialty: str,
    api_key: str | None = None,
    model: str = "claude-opus-4-6",
    stream_callback=None,
) -> str:
    """Generate a Pareto-based general report for a specialty."""
    key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        raise ValueError("ANTHROPIC_API_KEY not set.")

    client = anthropic.Anthropic(api_key=key)

    specialty_kw = SPECIALTY_KEYWORDS.get(specialty, specialty)

    prompt = f"""# SOLICITUD DE REPORTE GENERAL — ANÁLISIS DE PARETO

## ESPECIALIDAD: {specialty} (filtro: {specialty_kw})
## PERÍODO: {period}

---

## BASE DE DATOS DE CONSULTAS AUDITADAS:
{base_datos_text}

---

## TIPIFICACIONES DE USO COMÚN EN AUDITORÍA MÉDICA (Diccionario de Errores):
{tipificaciones_text}

---

## INSTRUCCIONES:
1. Filtra la base de datos por la especialidad **{specialty_kw}** y el período **{period}**.
2. Identifica la columna TIPIFICACION (PARETO) y agrupa por frecuencia de cada error.
3. Realiza el análisis de Pareto completo (ordenar, calcular % individual, % acumulado).
4. Identifica las Pocas Causas Vitales (80% acumulado).
5. Genera la Tabla de Pareto completa hasta el 100%.
6. Identifica los médicos en mayor riesgo y los mejor evaluados.
7. Genera el Resumen Ejecutivo (máximo 5 líneas).

GENERA EL REPORTE DE PARETO AHORA:
"""

    if stream_callback:
        full_text = ""
        with client.messages.stream(
            model=model,
            max_tokens=16384,
            system=PARETO_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        ) as stream:
            for text_chunk in stream.text_stream:
                full_text += text_chunk
                stream_callback(text_chunk)
        return full_text
    else:
        response = client.messages.create(
            model=model,
            max_tokens=16384,
            system=PARETO_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text
