"""
informe_adapter.py
──────────────────────────────────────────────────────────────────────────
Adaptador que convierte la salida de Claude (texto markdown parseado en
secciones) al formato estructurado `data` dict que espera generate_informe.

Flujo:
  Claude output (markdown) → parse_report_sections() → sections dict
  sections dict → build_informe_data() → data dict → generate_informe()
"""

import re


def _strip_md(text: str) -> str:
    """Elimina marcadores markdown de un texto."""
    text = re.sub(r"\[([^\]]+)\]\{\.underline\}", r"\1", text)
    # Strip raw HTML <u>/<b> tags the AI may generate
    text = re.sub(r"</?[ub]>", "", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    text = re.sub(r"^#{1,4}\s+", "", text, flags=re.MULTILINE)
    text = text.replace("*", "")
    return text.strip()


def _get_specialty_abbrev(specialty: str) -> str:
    spec = specialty.lower()
    if "servicio social" in spec or "medgen ss" in spec:
        return "SS"
    if "general" in spec or "medgen" in spec:
        return "MG"
    if "interna" in spec or "medint" in spec:
        return "MI"
    if "pediat" in spec or "pedia" in spec:
        return "PD"
    if "ginec" in spec or "gyobs" in spec or "giyobs" in spec:
        return "GY"
    if "psic" in spec:
        return "PS"
    if "nutri" in spec:
        return "NU"
    if "metab" in spec or "medmeta" in spec:
        return "ME"
    return "MG"


def _compute_period_code(period: str) -> str:
    """Calcula código de período (P001..P024) a partir del texto."""
    month_map = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
        "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
        "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    }
    lower = period.lower()
    for name, num in month_map.items():
        if name in lower:
            q = (num - 1) * 2 + 1
            if "16" in period or "segunda" in lower:
                q += 1
            return f"P{q:03d}"
    return "P001"


# ─── PARSERS DE TABLAS MARKDOWN → DATOS ESTRUCTURADOS ───────────────────────

def parse_compliance_rows(compliance_text: str, period_headers: list | None = None) -> list:
    """
    Convierte la tabla de cumplimiento en markdown a la lista de dicts
    que espera add_compliance_table():
      [{"type": "header"|"subheader"|"data", "cols": [...]}, ...]

    Formato esperado:
      | COMPONENTE | CRITERIO | period1 | period2 | ... | TENDENCIA |

    La última columna siempre es TENDENCIA. Detecta automáticamente los
    headers de período de la primera fila.
    """
    lines = [l.strip() for l in compliance_text.strip().split("\n") if l.strip()]
    # Excluir líneas de leyenda que Claude pueda incluir por error
    table_lines = [
        l for l in lines
        if "|" in l and "---" not in l
        and not l.lower().lstrip("| ").startswith("leyenda")
        and "leyenda:" not in l.lower()
    ]

    if not table_lines:
        return []

    rows_raw = []
    for line in table_lines:
        parts = line.split("|")
        # Trim leading/trailing empty strings from outer pipes, keep inner empties
        if parts and not parts[0].strip():
            parts = parts[1:]
        if parts and not parts[-1].strip():
            parts = parts[:-1]
        cells = [_strip_md(c.strip()) for c in parts]
        rows_raw.append(cells)

    if not rows_raw:
        return []

    # Determinar headers desde la primera fila
    header_row = rows_raw[0]

    # Detectar si la última columna es TENDENCIA
    has_tendencia = False
    if len(header_row) >= 3 and "TENDENCIA" in header_row[-1].upper():
        has_tendencia = True

    # Separar periodos y tendencia
    if has_tendencia:
        period_from_header = header_row[2:-1] if len(header_row) > 3 else []
        tendencia_label = header_row[-1]
    else:
        period_from_header = header_row[2:] if len(header_row) > 2 else []
        tendencia_label = "TENDENCIA"

    if period_headers is None:
        period_headers = period_from_header

    num_period_cols = len(period_headers)
    # COMPONENTE + CRITERIO + N períodos + TENDENCIA
    total_cols = 2 + num_period_cols + 1

    result = []

    # Header row
    header_cols = ["COMPONENTE", "CRITERIO"] + period_headers + [tendencia_label or "TENDENCIA"]
    while len(header_cols) < total_cols:
        header_cols.append("")
    result.append({"type": "header", "cols": header_cols[:total_cols]})

    # Data rows — agrupar por componente
    current_component = ""
    for row in rows_raw[1:]:
        if not row:
            continue

        # Skip filas de leyenda / total defensivamente
        first_upper = (row[0] if row else "").upper()
        if first_upper.startswith("LEYENDA") or first_upper.startswith("TOTAL"):
            continue

        comp = row[0] if len(row) > 0 else ""
        criterio = row[1] if len(row) > 1 else ""

        # Separar los períodos de la columna tendencia
        if has_tendencia and len(row) >= 3:
            pct_values = row[2:-1] if len(row) > 3 else []
            tendencia_val = row[-1] if len(row) > 2 else ""
        else:
            pct_values = row[2:] if len(row) > 2 else []
            tendencia_val = ""

        # Si hay un componente nuevo, agregar subheader
        if comp and comp.upper() != current_component.upper():
            current_component = comp
            sub_cols = [comp.upper()] + [""] * (total_cols - 1)
            result.append({"type": "subheader", "cols": sub_cols})

        # Data row
        data_cols = ["", criterio] + pct_values
        # Ajustar al número de columnas de período esperado
        while len(data_cols) < 2 + num_period_cols:
            data_cols.append("")
        data_cols = data_cols[:2 + num_period_cols]
        data_cols.append(tendencia_val)
        result.append({"type": "data", "cols": data_cols[:total_cols]})

    return result


def parse_citas(cuantitativo_text: str) -> list:
    """
    Convierte la tabla cuantitativa en markdown a la lista de dicts
    que espera add_quantitative_table():
      [{"num_cita": "...", "diagnostico": "...", "nota": "...", "sintesis": "..."}, ...]

    Soporta el formato actual (4 columnas: ID | DIAGNÓSTICO | NOTA | SÍNTESIS)
    y los formatos legados (5 cols con NC/ER y 4 cols legacy sin NOTA).
    """
    lines = [l.strip() for l in cuantitativo_text.strip().split("\n") if l.strip()]
    table_lines = [l for l in lines if "|" in l and "---" not in l]

    if not table_lines:
        return []

    rows_raw = []
    for line in table_lines:
        parts = line.split("|")
        if parts and not parts[0].strip():
            parts = parts[1:]
        if parts and not parts[-1].strip():
            parts = parts[:-1]
        cells = [_strip_md(c.strip()) for c in parts]
        rows_raw.append(cells)

    if not rows_raw:
        return []

    # Detect format from the header row
    header = [c.upper() for c in rows_raw[0]]
    header_joined = " | ".join(header)
    has_sintesis = "SÍNTESIS" in header_joined or "SINTESIS" in header_joined
    has_nc_er = " NC" in f" {header_joined}" and " ER" in f" {header_joined}"

    citas = []
    for row in rows_raw[1:]:
        if not row or len(row) < 3:
            continue
        # Skip total row
        if row[0].upper().startswith("TOTAL"):
            continue

        num_cita = row[0].strip()
        diagnostico = row[1].strip() if len(row) > 1 else ""
        nota = ""
        sintesis = ""

        if has_sintesis and len(row) >= 4:
            # New format: ID | DIAGNÓSTICO | NOTA | SÍNTESIS
            nota = row[2].strip()
            sintesis = row[3].strip()
        elif has_nc_er and len(row) >= 5:
            # Legacy: ID | DIAGNÓSTICO | NOTA | NC | ER
            nota = row[2].strip()
            nc_str = row[3].strip()
            er_str = row[4].strip()
            nc_val = re.sub(r"[^\d]", "", nc_str) or "0"
            er_val = re.sub(r"[^\d]", "", er_str) or "0"
            sintesis = f"NC: {nc_val} | ER: {er_val}"
        elif len(row) >= 4:
            # Best-effort: treat as ID | DIAG | NOTA | SÍNTESIS
            nota = row[2].strip()
            sintesis = row[3].strip()
        else:
            # ID | DIAG | SÍNTESIS (3 cols)
            sintesis = row[2].strip() if len(row) > 2 else ""

        citas.append({
            "num_cita": num_cita,
            "diagnostico": diagnostico,
            "nota": nota,
            "sintesis": sintesis,
        })

    return citas


def parse_hallazgos_from_resumen(resumen_text: str) -> dict:
    """
    Extrae hallazgos por componente del texto de resumen ejecutivo.
    Busca patrones como:
      "ANAMNESIS: se identifican 11 hallazgos; de los cuales 7 son No Conformidades y 4 son Eventos"
    """
    hallazgos = {}
    clean = _strip_md(resumen_text)

    # Patrón: COMPONENTE: ... N son No Conformidades y M son Eventos de Riesgo
    pattern = re.compile(
        r"([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+?):\s*se\s+identifican?\s+\d+\s+hallazgos?.*?"
        r"(\d+)\s+son\s+No\s+Conformidades?\s+y\s+(\d+)\s+son\s+Eventos?\s+de\s+Riesgo",
        re.IGNORECASE
    )

    for match in pattern.finditer(clean):
        comp = match.group(1).strip().upper()
        nc = int(match.group(2))
        er = int(match.group(3))
        hallazgos[comp] = {"nc": nc, "er": er}

    return hallazgos


def parse_hallazgos_from_tables(nc_table_text: str, er_table_text: str) -> dict:
    """
    Alternativa: calcula hallazgos por componente desde las tablas
    cualitativas estructuradas de NC y ER (sección 3.2).

    Cada tabla tiene el formato:
      | COMPONENTE | CRITERIO | NC/ER | TIPIFICACIÓN | IMPACTO |
    donde COMPONENTE puede estar vacío en filas de continuación.

    Devuelve { "COMPONENTE": {"nc": N, "er": N}, ... }.
    """
    def _tally(text: str, key: str) -> dict:
        out: dict = {}
        if not text:
            return out
        lines = [l.strip() for l in text.strip().split("\n") if l.strip()]
        table_lines = [l for l in lines if "|" in l and "---" not in l]
        if len(table_lines) < 2:
            return out

        rows_raw = []
        for line in table_lines:
            parts = line.split("|")
            if parts and not parts[0].strip():
                parts = parts[1:]
            if parts and not parts[-1].strip():
                parts = parts[:-1]
            rows_raw.append([_strip_md(c.strip()) for c in parts])

        last_comp = ""
        for row in rows_raw[1:]:
            if len(row) < 3:
                continue
            comp = row[0].strip() or last_comp
            if not comp:
                continue
            last_comp = comp
            count_str = re.sub(r"[^\d]", "", row[2])
            try:
                count = int(count_str) if count_str else 0
            except ValueError:
                count = 0
            comp_key = comp.upper()
            if comp_key not in out:
                out[comp_key] = {"nc": 0, "er": 0}
            out[comp_key][key] += count
        return out

    nc_by_comp = _tally(nc_table_text, "nc")
    er_by_comp = _tally(er_table_text, "er")

    hallazgos: dict = {}
    for comp, v in nc_by_comp.items():
        hallazgos.setdefault(comp, {"nc": 0, "er": 0})["nc"] += v["nc"]
    for comp, v in er_by_comp.items():
        hallazgos.setdefault(comp, {"nc": 0, "er": 0})["er"] += v["er"]
    return hallazgos


# ─── FUNCIÓN PRINCIPAL DE CONVERSIÓN ────────────────────────────────────────

def build_informe_data(
    doctor_name: str,
    doctor_code: str,
    period: str,
    specialty: str,
    sections: dict,
) -> dict:
    """
    Convierte las secciones parseadas de Claude al formato `data` dict
    que espera generate_informe().

    sections es el dict retornado por parse_report_sections():
      {resumen_ejecutivo, cumplimiento, seguimiento, cuantitativo,
       cualitativo, no_conformidades, eventos_riesgo}
    """
    spec_abbrev = _get_specialty_abbrev(specialty)
    period_code = _compute_period_code(period)

    # Parsear tablas estructuradas
    compliance_rows = []
    if sections.get("cumplimiento"):
        compliance_rows = parse_compliance_rows(sections["cumplimiento"])

    citas = []
    if sections.get("cuantitativo"):
        citas = parse_citas(sections["cuantitativo"])

    # Parsear hallazgos del resumen; si no hay, derivarlos de las tablas
    # cualitativas estructuradas (NC/ER) por componente.
    hallazgos = {}
    if sections.get("resumen_ejecutivo"):
        hallazgos = parse_hallazgos_from_resumen(sections["resumen_ejecutivo"])
    if not hallazgos:
        hallazgos = parse_hallazgos_from_tables(
            sections.get("nc_table", ""),
            sections.get("er_table", ""),
        )

    # Limpiar textos de markdown
    resumen = _strip_md(sections.get("resumen_ejecutivo", ""))
    seguimiento = _strip_md(sections.get("comentario_seguimiento", "")
                            or sections.get("seguimiento", ""))

    # Construir análisis cualitativo combinando las secciones
    cualitativo_parts = []

    # Prefer structured tables (nc_table / er_table) over old prose sections
    if sections.get("nc_table"):
        cualitativo_parts.append("TABLA DE NO CONFORMIDADES\n")
        cualitativo_parts.append(_strip_md(sections["nc_table"]))
    if sections.get("er_table"):
        cualitativo_parts.append("\nTABLA DE EVENTOS DE RIESGO\n")
        cualitativo_parts.append(_strip_md(sections["er_table"]))

    # Fallback to old prose sections if no structured tables
    if not cualitativo_parts:
        if sections.get("cualitativo"):
            cualitativo_parts.append(_strip_md(sections["cualitativo"]))
        if sections.get("no_conformidades"):
            cualitativo_parts.append("\nAnálisis de No Conformidades\n")
            cualitativo_parts.append(_strip_md(sections["no_conformidades"]))
        if sections.get("eventos_riesgo"):
            cualitativo_parts.append("\n" + "─" * 40 + "\n")
            cualitativo_parts.append("\nAnálisis de Eventos de Riesgo\n")
            cualitativo_parts.append(_strip_md(sections["eventos_riesgo"]))

    # Add síntesis if present
    if sections.get("sintesis"):
        cualitativo_parts.append("\nSÍNTESIS\n")
        cualitativo_parts.append(_strip_md(sections["sintesis"]))

    analisis_cualitativo = "\n".join(cualitativo_parts)

    return {
        # Encabezado
        "nombre": doctor_name.upper(),
        "codigo": doctor_code,
        "especialidad": specialty.upper(),
        "dependencia": "Doctor SV - El Salvador",
        "periodo": period,
        "periodo_codigo": period_code,
        "spec_abbrev": spec_abbrev,
        # Secciones
        "resumen_ejecutivo": resumen,
        "hallazgos": hallazgos,
        "compliance_rows": compliance_rows,
        "comentario_seguimiento": seguimiento,
        "citas": citas,
        "analisis_cualitativo": analisis_cualitativo,
        # Structured qualitative tables (raw pipe text for DOCX table rendering)
        "nc_table": sections.get("nc_table", ""),
        "er_table": sections.get("er_table", ""),
        "sintesis": _strip_md(sections.get("sintesis", "")),
        # Component compliance table (from Reporte Global)
        "cumplimiento_componentes": sections.get("cumplimiento_componentes", ""),
        # Tendencias
        "tendencias": sections.get("tendencias", ""),
        # Conclusiones
        "conclusiones": sections.get("conclusiones", ""),
        # Firma
        "firma": sections.get("firma", ""),
    }
