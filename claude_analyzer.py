"""
Claude AI analyzer module for medical audit reports.
Sends structured data to Claude API and receives formatted analysis.
"""

import anthropic
import os
from typing import Optional


# ─── AVAILABLE MODELS ───────────────────────────────────────────────────────
# Ordered from cheapest → most expensive. Default is Sonnet 4.6 which,
# combined with extended thinking + prompt caching, delivers near-Opus
# quality at ~2x lower cost per audit.
AVAILABLE_MODELS = {
    "claude-sonnet-4-6": {
        "label": "Claude Sonnet 4.6 (estándar — recomendado)",
        "description": "Balance calidad/costo. Usa extended thinking + prompt caching. ~$0.08/auditoría.",
        "supports_thinking": True,
    },
    "claude-opus-4-6": {
        "label": "Claude Opus 4.6 (premium)",
        "description": "Máxima calidad en razonamiento y prosa clínica. ~$0.18/auditoría.",
        "supports_thinking": True,
    },
    "claude-haiku-4-5-20251001": {
        "label": "Claude Haiku 4.5 (rápido/económico)",
        "description": "El más rápido y económico. Para auditorías simples. ~$0.02/auditoría.",
        "supports_thinking": False,
    },
}

DEFAULT_MODEL = "claude-sonnet-4-6"


SYSTEM_PROMPT = """Eres un experto en auditoría médica de calidad. Tu tarea es analizar datos de auditorías médicas y generar reportes estructurados siguiendo un formato ESTRICTO e idéntico al ejemplo proporcionado.

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
- Localiza al médico en la hoja de su especialidad usando su Código Único (COD), no solo por nombre.
- Confirma que la fila seleccionada corresponde exactamente al período solicitado.
- Valida de forma cruzada que los IDs de consulta de los hallazgos cualitativos pertenecen al mismo médico que figura en la fila de porcentajes.
- Si existe discrepancia entre hallazgos y porcentajes, transcribe los datos exactos de la celda; no asumas ni corrijas el error.
- Los IDs de consulta son números de entre 5 y 10 dígitos. Cuéntalos de forma individual y exacta.
- Cada tipificación se cuenta de forma exacta tal como está escrita en el documento de tipificaciones. No unas ni mezcles tipificaciones distintas.
- Correlaciona siempre el diagnóstico con cada hallazgo identificado.

FORMATO DE SALIDA OBLIGATORIO — Sigue este formato EXACTO, incluyendo las líneas separadoras y la estructura de encabezados. Las secciones están NUMERADAS:

---

INFORME DE AUDITORÍA N° [CÓDIGOdelMÉDICO]-[ABREV_ESPECIALIDAD]-[AÑO]-P[NÚM_PERÍODO]

NOMBRE: [NOMBRE COMPLETO DEL MÉDICO]
CÓDIGO: [CÓDIGO DEL MÉDICO]
ESPECIALIDAD: [ESPECIALIDAD COMPLETA]
DEPENDENCIA: Doctor SV - El Salvador
PERIODO AUDITADO: [Fecha inicio] al [Fecha fin] [mes] [año]

1. RESUMEN EJECUTIVO

El Resumen Ejecutivo se redacta siempre en texto corrido con campos etiquetados en negrita, sin tablas. Sigue obligatoriamente esta estructura de cuatro campos:

Lógica de adaptación según períodos disponibles:
- Si es el primer período auditado del médico: Los cuatro campos se redactan en modo síntesis descriptiva, sin comparación con período anterior.
- Si existen dos o más períodos auditados: Los cuatro campos se redactan comparando exclusivamente el último período cerrado contra el nuevo período auditado.

**Período de seguimiento:** [Indicar el período actual auditado. Si existe período previo, señalarlo explícitamente como referencia de continuidad. Si es primera auditoría, indicarlo como período inicial de evaluación.]

**Hallazgo principal:** [Redactar en una o dos oraciones el comportamiento más relevante del período. Si hay comparación, señalar si hubo retroceso, avance o estabilidad respecto al período anterior, nombrando los criterios más afectados con sus porcentajes en negrita. Si es primera auditoría, describir los criterios con mayor incumplimiento identificados.]

**Volumen de hallazgos:** [Indicar el total de **No Conformidades** y **Eventos de Riesgo**, con desglose por componente afectado y cantidad de hallazgos en cada uno en negrita.]

**Riesgo de seguridad:** [Señalar el hallazgo de mayor gravedad clínica o documental del período, redactado de forma ejecutiva y urgente. Si no existe riesgo crítico, indicar la oportunidad de mejora de mayor impacto potencial.]

CUMPLIMIENTO POR COMPONENTES

| Componente | % Cumplimiento |
|---|---|
| Anamnesis | XX% |
| Examen Físico | XX% |
| Diagnóstico | XX% |
| Productos de la Consulta | XX% |
| % Promedio | XX% |
| Puntaje Promedio | X.XX |

(NOTA: Los porcentajes de cumplimiento por componente se toman del Reporte Global si fue proporcionado. El % promedio es el promedio de los cuatro componentes. El puntaje promedio es % promedio / 20 redondeado a 2 decimales (escala de 5). Si no hay datos del Reporte Global, usa los promedios calculados de los criterios del cuadro de cumplimiento.)

TENDENCIAS

| TENDENCIA | [Fecha período anterior] | [Fecha período actual] |
|---|---|---|
| Positiva | [Criterio1 (XX%), Criterio2 (XX%), ...] | [Criterio1 (XX%), Criterio2 (XX%), ...] |
| Sostenida | [Criterio1 (XX%), Criterio2 (XX%), ...] | [Criterio1 (XX%), Criterio2 (XX%), ...] |
| Negativa | [Criterio1 (XX%), Criterio2 (XX%), ...] | [Criterio1 (XX%), Criterio2 (XX%), ...] |

Reglas de la tabla TENDENCIAS:
- La primera columna contiene ÚNICAMENTE el tipo de tendencia (Positiva / Sostenida / Negativa), sin explicación ni texto adicional.
- La segunda columna corresponde al período ANTERIOR evaluado; usa la fecha real de ese período como encabezado (ej.: "01-15 Febrero 2026").
- La tercera columna corresponde al período ACTUAL auditado; usa la fecha real de ese período como encabezado.
- En las filas, lista los criterios afectados con su porcentaje obtenido en ese período, asociándolos claramente al tipo de tendencia de la fila.
- Si una celda no aplica, coloca "—".

(NOTA: Si solo hay un período auditado, OMITE completamente la sección TENDENCIAS. No escribas "No aplica", simplemente no incluyas esta sección.)

Reglas de redacción del Resumen Ejecutivo:
- Tono ejecutivo, urgente y profesional. Sin rodeos ni texto de relleno.
- Los nombres de criterios afectados y sus porcentajes van siempre en **negrita** dentro del texto corrido.
- Máximo media página en total para los cuatro campos.
- No generar tablas dentro de los cuatro campos del Resumen Ejecutivo, solo la tabla de CUMPLIMIENTO POR COMPONENTES y la tabla de TENDENCIAS al final.
- No repetir información que ya aparece en el cuadro de cumplimiento de la sección 2.

2. CUADRO DE CUMPLIMIENTO POR CRITERIO

Nivel de cumplimiento por criterio evaluado, organizado por Criterio clínico. Los porcentajes se calculan sobre el total de citas auditadas.

| COMPONENTE | CRITERIO | [Período anterior] CUMPLIMIENTO | [Período actual] CUMPLIMIENTO |
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

Leyenda: Excelente (≥98%) | Muy Bueno (≥95% a <98%) | Aceptable (≥85% a <95%) | Op. de Mejora (<85%)

Categorización por color obligatoria:
🟢 Verde → ≥98% — Excelente
🟡 Amarillo → ≥95% a <98% — Muy Bueno
🟠 Naranja → ≥85% a <95% — Aceptable
🔴 Rojo → <85% — Oportunidad de Mejora

REGLAS del cuadro de cumplimiento:
- SIEMPRE genera la tabla markdown con pipes (|) aunque solo haya un período. NUNCA omitas esta tabla.
- Usa EXACTAMENTE los porcentajes del archivo de gráficas. NO los inventes.
- Si un criterio no aplica, usa "-" en lugar de porcentaje.
- Incluye TODOS los períodos disponibles en columnas separadas.
- La columna COMPONENTE debe repetir el nombre del componente en cada fila que le pertenezca.
- Si no se proporcionaron datos de cumplimiento, genera la tabla con los criterios estándar y usa "-" en todas las columnas de porcentaje.

COMENTARIO DE SEGUIMIENTO Y COMPARACIÓN DE PERIODOS
(Solo si hay 2 o más periodos auditados)

"Comentario de seguimiento y comparación de periodos ([Período anterior] vs [Período actual]):
(Resumen de máximo 4 líneas comparando los criterios con impacto en la atención al paciente).
Tendencia positiva: (únicamente los criterios con aumento en el porcentaje)
Tendencia Negativa: (únicamente los criterios con disminución del porcentaje)
Tendencia sostenida: (criterios sin variación entre periodos — énfasis especial en los que estén por debajo del 90%)"

3. ANÁLISIS DE NO CONFORMIDADES

3.1 ANÁLISIS CUANTITATIVO

| ID CITA | DIAGNÓSTICO (CIE-11) | NOTA | SÍNTESIS DE HALLAZGOS |
|---|---|---|---|
| [ID consulta] | [Código CIE] - [Descripción] | [Nota/calificación de la cita] | [Síntesis de hallazgos encontrados en esa cita, máximo 3 líneas] |
| ... | ... | ... | ... |

La columna Nota corresponde a la calificación registrada para cada cita en la base de datos. Transcríbela de forma exacta.

La columna SÍNTESIS DE HALLAZGOS debe resumir en máximo 3 líneas todos los hallazgos (no conformidades y eventos de riesgo) encontrados en esa consulta. Menciona los criterios y tipificaciones afectados de forma concisa. Si la cita no tiene hallazgos, coloca "Sin hallazgos".

3.2 ANÁLISIS CUALITATIVO POR COMPONENTE

El análisis cualitativo se presentará siempre en formato de tabla, nunca como texto corrido. Se construirán dos tablas independientes: una para No Conformidades y otra para Eventos de Riesgo (esta segunda solo si ER > 0).

TABLA DE NO CONFORMIDADES

| COMPONENTE | CRITERIO | NC | TIPIFICACIÓN | IMPACTO EN LA ATENCIÓN |
|---|---|---|---|---|
| [COMPONENTE] | [Criterio afectado] | [N] | [Texto exacto del diccionario] | [Impacto clínico sintetizado, máximo 2 líneas] (IDs: [ID1], [ID2], ...) |
| ... | ... | ... | ... | ... |

TABLA DE EVENTOS DE RIESGO

| COMPONENTE | CRITERIO | ER | TIPIFICACIÓN | IMPACTO EN LA ATENCIÓN |
|---|---|---|---|---|
| [COMPONENTE] | [Criterio afectado] | [N] | [Texto exacto del diccionario] | [Impacto clínico sintetizado, máximo 2 líneas] (IDs: [ID1], [ID2], ...) |
| ... | ... | ... | ... | ... |

Reglas de construcción de tablas cualitativas:
- El orden de filas sigue siempre: Anamnesis → Examen Físico → Diagnóstico → Productos de la Consulta.
- Cada fila representa una sola tipificación. Si un criterio tiene dos tipificaciones distintas, ocupa dos filas separadas.
- La columna COMPONENTE se rellena solo en la primera fila del grupo; las filas siguientes del mismo componente quedan en blanco.
- La columna NC / ER muestra el conteo de ese hallazgo específico.
- La columna TIPIFICACIÓN contiene el texto exacto del documento de tipificaciones de uso común, sin parafrasear.
- La columna IMPACTO EN LA ATENCIÓN es una frase concisa (máximo 2 líneas) con la consecuencia clínica o documental del hallazgo, SEGUIDA OBLIGATORIAMENTE de los IDs de las citas donde se encontró esa tipificación, entre paréntesis al final, con el formato: `(IDs: 12345678, 87654321)`.
- Los IDs citados en el paréntesis de IMPACTO deben coincidir exactamente con los IDs de la tabla cuantitativa de la sección 3.1 que presenten esa tipificación.

SÍNTESIS

[Párrafo corto de síntesis (máximo 3 líneas) que identifica el patrón de riesgo dominante del período.]

4. CONCLUSIONES Y ACCIONES REQUERIDAS

| PRIORIDAD | ACCIÓN REQUERIDA |
|---|---|
| CRÍTICA | [Acción correctiva urgente...] |
| ALTA | [Acción correctiva importante...] |
| MEDIA | [Acción de mejora continua...] |

(NOTA: Incluir entre 3 y 6 acciones. Prioridades posibles: CRÍTICA, ALTA, MEDIA. Las acciones deben ser específicas, medibles y relacionadas directamente con los hallazgos.)

FIRMA

Auditor Responsable: UGMC — Unidad de Gestión de Mejora Continua
Fecha de Emisión: [Fecha del día de generación del informe en formato DD/MM/YYYY]

---

IMPORTANTE sobre el formato:
- ESTRICTAMENTE, sigue el formato línea por línea como se especifica arriba.
- Las secciones están NUMERADAS: 1. RESUMEN EJECUTIVO, 2. CUADRO DE CUMPLIMIENTO, 3. ANÁLISIS DE NO CONFORMIDADES, 4. CONCLUSIONES Y ACCIONES REQUERIDAS.
- El RESUMEN EJECUTIVO tiene 4 campos en negrita (Período de seguimiento, Hallazgo principal, Volumen de hallazgos, Riesgo de seguridad) seguidos de la tabla CUMPLIMIENTO POR COMPONENTES y opcionalmente TENDENCIAS.
- TENDENCIAS solo se incluye si hay 2 o más períodos auditados. Si es primera auditoría, NO incluyas TENDENCIAS.
- Las tablas de análisis cualitativo (TABLA DE NO CONFORMIDADES, TABLA DE EVENTOS DE RIESGO) deben ser tablas markdown con pipes, NO prosa narrativa.
- Cada fila de la tabla cualitativa = un hallazgo con su tipificación e impacto.
- La sección SÍNTESIS va al final de la sección 3, como un párrafo integrador de máximo 3 líneas.
- Los hallazgos SIEMPRE deben detallar CADA consulta individualmente con su ID y diagnóstico.
- Tono: Ejecutivo, urgente pero profesional. Evita rodeos innecesarios.
- Máximo 4 páginas por informe.
- El número de informe sigue el formato: [CÓDIGO]-[ABREV_ESP]-[AÑO]-P[NÚM_PERÍODO] (ej: 000FV1-MG-2026-P003)
  - Abreviaturas de especialidad: MG (Medicina General), MI (Medicina Interna), PD (Pediatría), GY (Ginecología), PS (Psicología), NU (Nutrición), SS (Servicio Social)
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
    reporte_global_table: str = "",
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
    if reporte_global_table:
        prompt += f"""
## DATOS DE REPORTE GLOBAL (Cumplimiento por Componente):
{reporte_global_table}

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
6. Para el CUMPLIMIENTO POR COMPONENTES en el Resumen Ejecutivo, usa los datos del Reporte Global si fueron proporcionados.
7. Si hay datos de período anterior, incluye el comentario de seguimiento y la tabla de TENDENCIAS comparando los últimos dos períodos.
8. Si es el primer período auditado (sin datos de período anterior), NO incluyas la sección TENDENCIAS.
9. Aplica las categorías de color en texto: ROJO (<85%), ANARANJADO (85%-94%), AMARILLO (95%-97%), VERDE (98%-100%).
10. Cuenta correctamente cada ID de consulta (números de 5 a 10 dígitos).
11. Correlaciona el diagnóstico con cada hallazgo encontrado.

GENERA EL REPORTE AHORA:
"""
    return prompt


def _build_request_params(
    model: str,
    system_prompt: str,
    user_prompt: str,
    thinking_enabled: bool = False,
    thinking_budget: int = 4000,
    max_response_tokens: int = 16384,
) -> dict:
    """Build the request params for a messages.create / messages.stream call.

    Applies prompt caching on the system prompt (ephemeral cache, 90% discount
    on subsequent calls within the cache window) and optionally enables
    extended thinking for reasoning-capable models.
    """
    # System prompt as a list of blocks so we can attach cache_control
    system_blocks = [{
        "type": "text",
        "text": system_prompt,
        "cache_control": {"type": "ephemeral"},
    }]

    params: dict = {
        "model": model,
        "max_tokens": max_response_tokens,
        "system": system_blocks,
        "messages": [{"role": "user", "content": user_prompt}],
    }

    if thinking_enabled:
        # The API requires max_tokens > budget_tokens; expand the total cap
        params["thinking"] = {
            "type": "enabled",
            "budget_tokens": thinking_budget,
        }
        params["max_tokens"] = max_response_tokens + thinking_budget
        # Extended thinking requires temperature = 1 (default)
        params.pop("temperature", None)

    return params


def _extract_text_from_response(response) -> str:
    """Return the concatenated text from non-thinking content blocks."""
    if not response.content:
        raise ValueError("Claude returned an empty response")
    text_parts = []
    for block in response.content:
        # SDK returns typed blocks; thinking blocks use .type == "thinking"
        btype = getattr(block, "type", None)
        if btype == "text":
            text_parts.append(getattr(block, "text", ""))
    if not text_parts:
        # Fallback: first block text (legacy behaviour)
        first = response.content[0]
        return getattr(first, "text", "") or ""
    return "".join(text_parts)


def analyze_with_claude(
    doctor_name: str,
    doctor_code: str,
    period: str,
    consultations_table: str,
    compliance_table: str,
    clasificacion_table: str,
    tipificaciones_list: str,
    previous_period_data: str = "",
    reporte_global_table: str = "",
    api_key: str | None = None,
    model: str = "claude-sonnet-4-6",
    thinking_enabled: bool = True,
    thinking_budget: int = 4000,
    stream_callback=None,
) -> str:
    """
    Send data to Claude API and get the formatted audit report.

    Args:
        model: Claude model id. Defaults to Sonnet 4.6 (estándar). Use
            `claude-opus-4-6` for premium mode.
        thinking_enabled: If True, enables extended thinking (chain-of-thought
            before the response) to close the quality gap with Opus. Adds some
            cost but improves cross-reference accuracy.
        thinking_budget: Token budget for the thinking phase (only used when
            thinking_enabled=True).
        stream_callback: Optional callable(chunk: str) for streaming text.
    Returns:
        Full report text (thinking blocks are discarded from the returned string).
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
        reporte_global_table=reporte_global_table,
    )

    params = _build_request_params(
        model=model,
        system_prompt=SYSTEM_PROMPT,
        user_prompt=prompt,
        thinking_enabled=thinking_enabled,
        thinking_budget=thinking_budget,
        max_response_tokens=16384,
    )

    if stream_callback:
        full_text = ""
        with client.messages.stream(**params) as stream:
            # text_stream yields only text deltas (thinking deltas are
            # handled internally by the SDK and excluded from this iterator).
            for text_chunk in stream.text_stream:
                full_text += text_chunk
                stream_callback(text_chunk)
        return full_text
    else:
        response = client.messages.create(**params)
        return _extract_text_from_response(response)


def parse_report_sections(report_text: str) -> dict:
    """
    Parse the Claude report into sections for document generation.
    Returns dict with keys: cuantitativo, cualitativo, no_conformidades,
    eventos_riesgo, resumen_ejecutivo, cumplimiento, seguimiento,
    tendencias, score_summary, nc_table, er_table, sintesis, conclusiones
    """
    sections = {
        "cuantitativo": "",
        "cualitativo": "",
        "no_conformidades": "",
        "eventos_riesgo": "",
        "resumen_ejecutivo": "",
        "cumplimiento": "",
        "seguimiento": "",
        "tendencias": "",
        "score_summary": "",
        "cumplimiento_componentes": "",
        "nc_table": "",
        "er_table": "",
        "sintesis": "",
        "conclusiones": "",
        "firma": "",
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

    # Skip the INFORME header block (NOMBRE, CÓDIGO, ESPECIALIDAD, etc.)
    skip_header = True

    for line in lines:
        line_stripped = line.strip()
        upper = line_stripped.upper()
        # Remove markdown bold markers and leading numbers for matching
        clean_upper = upper.replace("**", "").replace("#", "").strip()
        # Strip leading section numbers like "1. ", "3.1 ", "3.2 "
        clean_upper_no_num = clean_upper.lstrip("0123456789. ")

        # Skip informe header lines until we hit a real section
        if skip_header:
            # Always skip empty lines while in header
            if not upper:
                continue
            if any(kw in clean_upper for kw in [
                "INFORME DE AUDITORÍA", "INFORME DE AUDITORIA",
                "NOMBRE:", "CÓDIGO:", "CODIGO:",
                "ESPECIALIDAD:", "DEPENDENCIA:",
                "PERIODO AUDITADO", "TOTAL CONSULTAS",
                "NO CONFORMIDADES:", "EVENTOS DE RIESGO:",
            ]):
                continue
            if line_stripped == "---":
                continue
            if not any(kw in clean_upper for kw in [
                "RESUMEN", "REPORTE", "CUADRO", "ANÁLISIS", "ANALISIS",
                "COMENTARIO", "TENDENCIA", "SCORE", "CONCLUSI", "SÍNTESIS", "SINTESIS",
                "TABLA DE NO CONFORMIDADES", "TABLA DE EVENTOS",
                "CUMPLIMIENTO POR COMPONENTE", "FIRMA",
            ]):
                if not current_section:
                    continue
            skip_header = False

        # ── Section detection (order matters for specificity) ──

        # CUMPLIMIENTO POR COMPONENTES (within Resumen Ejecutivo)
        if "CUMPLIMIENTO POR COMPONENTES" in clean_upper or "CUMPLIMIENTO POR COMPONENTE" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "cumplimiento_componentes"
        # TENDENCIAS
        elif clean_upper_no_num == "TENDENCIAS" or clean_upper == "TENDENCIAS":
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "tendencias"
        # FIRMA
        elif clean_upper_no_num == "FIRMA" or clean_upper == "FIRMA":
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "firma"
        # SCORE SUMMARY (legacy compatibility)
        elif "SCORE SUMMARY" in clean_upper or "SCORE_SUMMARY" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "score_summary"
        # ANÁLISIS CUANTITATIVO (3.1)
        elif "ANÁLISIS CUANTITATIVO" in clean_upper or "ANALISIS CUANTITATIVO" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "cuantitativo"
        # TABLA DE NO CONFORMIDADES (qualitative structured table)
        elif "TABLA DE NO CONFORMIDADES" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "nc_table"
        # TABLA DE EVENTOS DE RIESGO (qualitative structured table)
        elif "TABLA DE EVENTOS DE RIESGO" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "er_table"
        # ANÁLISIS CUALITATIVO (3.2) — old or new format
        elif "ANÁLISIS CUALITATIVO" in clean_upper or "ANALISIS CUALITATIVO" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "cualitativo"
        # ANÁLISIS DE NO CONFORMIDADES — standalone section header (section 3)
        elif ("ANÁLISIS DE NO CONFORMIDADES" in clean_upper or "ANALISIS DE NO CONFORMIDADES" in clean_upper) and \
             clean_upper_no_num in ("ANÁLISIS DE NO CONFORMIDADES", "ANALISIS DE NO CONFORMIDADES"):
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "no_conformidades"
        # ANÁLISIS DE EVENTOS DE RIESGO (old format fallback)
        elif "ANÁLISIS DE EVENTOS DE RIESGO" in clean_upper or "ANALISIS DE EVENTOS DE RIESGO" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "eventos_riesgo"
        # SÍNTESIS
        elif clean_upper_no_num in ("SÍNTESIS", "SINTESIS"):
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "sintesis"
        # CONCLUSIONES Y ACCIONES REQUERIDAS (section 4)
        elif "CONCLUSIONES" in clean_upper and "ACCIONES" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "conclusiones"
        # RESUMEN EJECUTIVO (section 1)
        elif "RESUMEN EJECUTIVO" in clean_upper:
            skip_header = False
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "resumen_ejecutivo"
        # CUADRO DE CUMPLIMIENTO / REPORTE DE CUMPLIMIENTO (section 2)
        elif "CUADRO DE CUMPLIMIENTO" in clean_upper or "REPORTE DE CUMPLIMIENTO" in clean_upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "cumplimiento"
        # COMENTARIO DE SEGUIMIENTO
        elif "COMENTARIO DE SEGUIMIENTO" in clean_upper:
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
    model: str = "claude-sonnet-4-6",
    thinking_enabled: bool = True,
    thinking_budget: int = 4000,
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

    params = _build_request_params(
        model=model,
        system_prompt=PARETO_SYSTEM_PROMPT,
        user_prompt=prompt,
        thinking_enabled=thinking_enabled,
        thinking_budget=thinking_budget,
        max_response_tokens=16384,
    )

    if stream_callback:
        full_text = ""
        with client.messages.stream(**params) as stream:
            for text_chunk in stream.text_stream:
                full_text += text_chunk
                stream_callback(text_chunk)
        return full_text
    else:
        response = client.messages.create(**params)
        return _extract_text_from_response(response)
