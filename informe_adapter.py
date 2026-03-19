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

    Detecta automáticamente los headers de período de la primera fila.
    """
    lines = [l.strip() for l in compliance_text.strip().split("\n") if l.strip()]
    table_lines = [l for l in lines if "|" in l and "---" not in l]

    if not table_lines:
        return []

    rows_raw = []
    for line in table_lines:
        cells = [_strip_md(c.strip()) for c in line.split("|") if c.strip()]
        rows_raw.append(cells)

    if not rows_raw:
        return []

    # Determinar headers de período desde la primera fila
    header_row = rows_raw[0]
    # Columnas: [COMPONENTE, CRITERIO, period1, period2, ...]
    if period_headers is None:
        period_headers = header_row[2:] if len(header_row) > 2 else []

    num_period_cols = len(period_headers)
    total_cols = 2 + num_period_cols  # COMPONENTE + CRITERIO + N períodos

    # Ajustar col widths dinámicamente
    result = []

    # Header row
    header_cols = ["COMPONENTE", "CRITERIO"] + period_headers
    # Pad to total_cols
    while len(header_cols) < total_cols:
        header_cols.append("")
    result.append({"type": "header", "cols": header_cols[:total_cols]})

    # Data rows — agrupar por componente
    current_component = ""
    for row in rows_raw[1:]:
        if not row:
            continue

        comp = row[0] if len(row) > 0 else ""
        criterio = row[1] if len(row) > 1 else ""
        pct_values = row[2:] if len(row) > 2 else []

        # Si hay un componente nuevo, agregar subheader
        if comp and comp.upper() != current_component.upper():
            current_component = comp
            sub_cols = [comp.upper()] + [""] * (total_cols - 1)
            result.append({"type": "subheader", "cols": sub_cols})

        # Data row
        data_cols = ["", criterio] + pct_values
        while len(data_cols) < total_cols:
            data_cols.append("")
        result.append({"type": "data", "cols": data_cols[:total_cols]})

    return result


def parse_citas(cuantitativo_text: str) -> list:
    """
    Convierte la tabla cuantitativa en markdown a la lista de dicts
    que espera add_quantitative_table():
      [{"num_cita": "...", "diagnostico": "...", "nc": N, "er": N}, ...]
    """
    lines = [l.strip() for l in cuantitativo_text.strip().split("\n") if l.strip()]
    table_lines = [l for l in lines if "|" in l and "---" not in l]

    if not table_lines:
        return []

    rows_raw = []
    for line in table_lines:
        cells = [_strip_md(c.strip()) for c in line.split("|") if c.strip()]
        rows_raw.append(cells)

    # Skip header row (first row), parse data rows, skip TOTAL row
    citas = []
    for row in rows_raw[1:]:
        if not row or len(row) < 4:
            continue
        # Skip total row
        if row[0].upper() == "TOTAL":
            continue
        try:
            nc = int(re.sub(r"[^\d]", "", row[2])) if row[2].strip() else 0
        except (ValueError, IndexError):
            nc = 0
        try:
            er = int(re.sub(r"[^\d]", "", row[3])) if row[3].strip() else 0
        except (ValueError, IndexError):
            er = 0
        citas.append({
            "num_cita": row[0].strip(),
            "diagnostico": row[1].strip(),
            "nc": nc,
            "er": er,
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


def parse_hallazgos_from_citas(citas: list) -> dict:
    """
    Alternativa: calcula hallazgos totales desde la tabla cuantitativa.
    Agrupa NC y ER totales (sin desglose por componente).
    """
    total_nc = sum(c.get("nc", 0) for c in citas)
    total_er = sum(c.get("er", 0) for c in citas)
    if total_nc > 0 or total_er > 0:
        return {"TOTAL": {"nc": total_nc, "er": total_er}}
    return {}


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

    # Parsear hallazgos del resumen
    hallazgos = {}
    if sections.get("resumen_ejecutivo"):
        hallazgos = parse_hallazgos_from_resumen(sections["resumen_ejecutivo"])
    if not hallazgos and citas:
        hallazgos = parse_hallazgos_from_citas(citas)

    # Limpiar textos de markdown
    resumen = _strip_md(sections.get("resumen_ejecutivo", ""))
    seguimiento = _strip_md(sections.get("comentario_seguimiento", "")
                            or sections.get("seguimiento", ""))

    # Construir análisis cualitativo combinando las secciones
    cualitativo_parts = []
    if sections.get("cualitativo"):
        cualitativo_parts.append(_strip_md(sections["cualitativo"]))
    if sections.get("no_conformidades"):
        cualitativo_parts.append("\nAnálisis de No Conformidades\n")
        cualitativo_parts.append(_strip_md(sections["no_conformidades"]))
    if sections.get("eventos_riesgo"):
        cualitativo_parts.append("\n" + "─" * 40 + "\n")
        cualitativo_parts.append("\nAnálisis de Eventos de Riesgo\n")
        cualitativo_parts.append(_strip_md(sections["eventos_riesgo"]))

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
    }
