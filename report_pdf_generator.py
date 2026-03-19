"""
PDF Report Generator.
Generates audit reports as HTML with exact CSS styling, then converts to PDF.
Eliminates markdown character leaking issues inherent in python-docx approach.
"""

import re
import io
import base64
from datetime import date


# ─── COLOR PALETTE (matches the Word template) ──────────────────────────────
COLORS = {
    "header_bg": "#1F3964",
    "header_text": "#FFFFFF",
    "section_bg": "#2E74B5",
    "subsection_bg": "#BDD7EE",
    "body_text": "#000000",
    "table_header": "#1F3964",
    "table_alt": "#D6E4F7",
    "rojo": "#FF0000",
    "anaranjado": "#FF6600",
    "amarillo": "#FFC000",
    "verde": "#00B050",
}

FONT = "Calibri, Arial, sans-serif"

# ─── CSS ─────────────────────────────────────────────────────────────────────
BASE_CSS = f"""
@page {{
    size: Letter;
    margin: 2cm 2.5cm;
    @top-center {{
        content: "";
    }}
}}
body {{
    font-family: {FONT};
    font-size: 11pt;
    color: {COLORS["body_text"]};
    line-height: 1.4;
}}
h1 {{
    color: {COLORS["header_bg"]};
    font-size: 14pt;
    text-align: center;
    text-decoration: underline;
    margin-top: 20px;
    margin-bottom: 10px;
}}
h2 {{
    color: {COLORS["header_bg"]};
    font-size: 12pt;
    margin-top: 16px;
    margin-bottom: 8px;
}}
h3 {{
    color: {COLORS["header_bg"]};
    font-size: 11pt;
    margin-top: 12px;
    margin-bottom: 6px;
}}
p {{
    text-align: justify;
    margin: 4px 0;
}}
.cover-table {{
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
}}
.cover-banner {{
    background-color: {COLORS["header_bg"]};
    color: {COLORS["header_text"]};
    text-align: center;
    font-weight: bold;
    font-size: 14pt;
    padding: 10px;
}}
.cover-row td {{
    background-color: {COLORS["subsection_bg"]};
    padding: 6px 12px;
    border: 1px solid #999;
}}
.cover-label {{
    font-weight: bold;
    width: 180px;
}}
.section-banner {{
    background-color: {COLORS["section_bg"]};
    color: {COLORS["header_text"]};
    font-weight: bold;
    font-size: 12pt;
    padding: 8px 16px;
    margin-top: 20px;
    margin-bottom: 10px;
}}
.data-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0;
    font-size: 9pt;
}}
.data-table th {{
    background-color: {COLORS["table_header"]};
    color: {COLORS["header_text"]};
    padding: 6px 8px;
    text-align: center;
    font-weight: bold;
    border: 1px solid #666;
}}
.data-table td {{
    padding: 5px 8px;
    border: 1px solid #CCC;
}}
.data-table tr:nth-child(even) td {{
    background-color: {COLORS["table_alt"]};
}}
.comp-cell {{
    background-color: {COLORS["header_bg"]} !important;
    color: {COLORS["header_text"]};
    font-weight: bold;
    text-align: center;
    vertical-align: middle;
}}
.pct-cell {{
    text-align: center;
    font-weight: bold;
    color: white;
}}
.separator {{
    border: none;
    border-top: 2px solid {COLORS["section_bg"]};
    margin: 20px 0;
}}
.legend {{
    display: flex;
    gap: 16px;
    font-size: 8pt;
    margin-top: 8px;
    flex-wrap: wrap;
}}
.legend-item {{
    display: flex;
    align-items: center;
    gap: 4px;
}}
.legend-box {{
    width: 12px;
    height: 12px;
    display: inline-block;
}}
.footer {{
    text-align: center;
    font-size: 9pt;
    color: #7F7F7F;
    font-style: italic;
    margin-top: 30px;
}}
.page-break {{
    page-break-before: always;
}}
.chart-img {{
    display: block;
    margin: 10px auto;
    max-width: 100%;
}}
"""


def _get_specialty_abbrev(specialty: str) -> str:
    spec_lower = specialty.lower()
    if "servicio social" in spec_lower or "medgen ss" in spec_lower:
        return "SS"
    if "general" in spec_lower or "medgen" in spec_lower:
        return "MG"
    if "interna" in spec_lower or "medint" in spec_lower:
        return "MI"
    if "pediat" in spec_lower or "pedia" in spec_lower:
        return "PD"
    if "ginec" in spec_lower or "gyobs" in spec_lower or "giyobs" in spec_lower:
        return "GY"
    if "psic" in spec_lower:
        return "PS"
    if "nutri" in spec_lower:
        return "NU"
    return "MG"


def _build_informe_number(doctor_code: str, specialty: str, period: str) -> str:
    spec_abbrev = _get_specialty_abbrev(specialty)
    year_match = re.search(r"20\d{2}", period)
    year_str = year_match.group(0) if year_match else str(date.today().year)
    period_num = "001"
    month_map = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
        "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    }
    period_lower = period.lower()
    for m_name, m_num in month_map.items():
        if m_name in period_lower:
            q_num = (m_num - 1) * 2 + 1
            if "16" in period or "segunda" in period_lower:
                q_num += 1
            period_num = f"{q_num:03d}"
            break
    return f"{doctor_code}-{spec_abbrev}-{year_str}-P{period_num}"


def _esc(text: str) -> str:
    """Escape HTML special characters."""
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _get_compliance_color(pct_str: str) -> str:
    """Return hex color based on compliance percentage."""
    try:
        val = float(pct_str.replace("%", "").replace(",", ".").strip())
        if val < 85:
            return COLORS["rojo"]
        elif val < 95:
            return COLORS["anaranjado"]
        elif val < 98:
            return COLORS["amarillo"]
        else:
            return COLORS["verde"]
    except Exception:
        return "transparent"


def _strip_markdown(text: str) -> str:
    """Remove markdown formatting, returning clean text."""
    text = re.sub(r"\[([^\]]+)\]\{\.underline\}", r"\1", text)
    # Strip raw HTML <u>/<b> tags the AI may generate
    text = re.sub(r"</?[ub]>", "", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    text = text.replace("*", "")
    return text.strip()


def _md_to_html_inline(text: str) -> str:
    """Convert inline markdown (bold, underline) to HTML tags."""
    # [text]{.underline} → <u><b>text</b></u>
    text = re.sub(r"\[([^\]]+)\]\{\.underline\}", r"<u><b>\1</b></u>", text)
    # Restore escaped <u>/<b> tags that the AI may have generated directly
    text = re.sub(r"&lt;u&gt;(.*?)&lt;/u&gt;", r"<u>\1</u>", text)
    text = re.sub(r"&lt;b&gt;(.*?)&lt;/b&gt;", r"<b>\1</b>", text)
    # **bold** → <b>bold</b>
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    # Clean stray asterisks
    text = text.replace("*", "")
    return text


def _body_text_to_html(text: str) -> str:
    """Convert markdown body text to HTML paragraphs."""
    lines = text.split("\n")
    html_parts = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Separator
        if all(c in "─_═━" for c in stripped) and len(stripped) > 10:
            html_parts.append('<hr class="separator"/>')
            continue

        # Headings
        heading_match = re.match(r"^(#{1,4})\s+(.+)$", stripped)
        if heading_match:
            level = len(heading_match.group(1))
            h_text = _strip_markdown(heading_match.group(2))
            html_parts.append(f"<h{level}>{_esc(h_text)}</h{level}>")
            continue

        # Numbered heading (e.g. "1. ANAMNESIS")
        num_match = re.match(r"^(\d+\.)\s+(.+)$", stripped)
        if num_match and num_match.group(2).replace(" ", "").isupper():
            h_text = _strip_markdown(f"{num_match.group(1)} {num_match.group(2)}")
            html_parts.append(f'<h2>{_esc(h_text)}</h2>')
            continue

        # Bullet points
        bullet_match = re.match(r"^[-•]\s+(.+)$", stripped)
        if bullet_match:
            html_parts.append(f"<p>&bull; {_md_to_html_inline(_esc(bullet_match.group(1)))}</p>")
            continue

        # Regular paragraph
        html_parts.append(f"<p>{_md_to_html_inline(_esc(stripped))}</p>")

    return "\n".join(html_parts)


def _compliance_table_to_html(text: str) -> str:
    """Convert markdown compliance table to HTML with color-coded cells."""
    lines = [l.strip() for l in text.strip().split("\n") if l.strip()]
    table_lines = [l for l in lines if "|" in l and "---" not in l]
    if not table_lines:
        return f"<p>{_esc(text)}</p>"

    rows = []
    for line in table_lines:
        cells = [c.strip() for c in line.split("|") if c.strip()]
        rows.append(cells)

    if not rows:
        return ""

    html = ['<table class="data-table">']

    # Header
    html.append("<thead><tr>")
    for cell in rows[0]:
        html.append(f"<th>{_esc(cell)}</th>")
    html.append("</tr></thead>")

    # Group rows by component for rowspan
    html.append("<tbody>")
    i = 1
    while i < len(rows):
        row = rows[i]
        comp = row[0] if row else ""

        # Count consecutive rows with same component
        span = 1
        while i + span < len(rows) and rows[i + span][0] == comp:
            span += 1

        for s in range(span):
            html.append("<tr>")
            actual_row = rows[i + s]
            for j, cell in enumerate(actual_row):
                if j == 0:
                    if s == 0:
                        html.append(f'<td class="comp-cell" rowspan="{span}">{_esc(comp)}</td>')
                    # Skip for subsequent rows (covered by rowspan)
                elif j == 1:
                    html.append(f"<td>{_esc(cell)}</td>")
                else:
                    if "%" in cell:
                        color = _get_compliance_color(cell)
                        html.append(f'<td class="pct-cell" style="background-color:{color}">{_esc(cell)}</td>')
                    elif cell.strip() == "-":
                        html.append(f'<td style="text-align:center;font-style:italic">{_esc(cell)}</td>')
                    else:
                        html.append(f'<td style="text-align:center">{_esc(cell)}</td>')
            html.append("</tr>")
        i += span

    html.append("</tbody></table>")

    # Color legend
    html.append('<div class="legend">')
    legend_items = [
        (COLORS["verde"], "≥98% Óptimo"),
        (COLORS["amarillo"], "≥ 95% a < 98% Muy Bueno"),
        (COLORS["anaranjado"], "≥ 85% a < 95% Aceptable"),
        (COLORS["rojo"], "<85% Oportunidad de mejora"),
    ]
    for color, label in legend_items:
        html.append(f'<span class="legend-item"><span class="legend-box" style="background:{color}"></span>{label}</span>')
    html.append("</div>")

    return "\n".join(html)


def _quantitative_table_to_html(text: str) -> str:
    """Convert markdown quantitative table to HTML."""
    lines = [l.strip() for l in text.strip().split("\n") if l.strip()]
    table_lines = [l for l in lines if "|" in l and "---" not in l]
    if not table_lines:
        return f"<p>{_esc(text)}</p>"

    rows = []
    for line in table_lines:
        cells = [c.strip() for c in line.split("|") if c.strip()]
        rows.append(cells)

    if not rows:
        return ""

    html = ['<table class="data-table">']
    html.append("<thead><tr>")
    for cell in rows[0]:
        html.append(f"<th>{_esc(_strip_markdown(cell))}</th>")
    html.append("</tr></thead><tbody>")

    for row in rows[1:]:
        html.append("<tr>")
        for cell in row:
            clean = _strip_markdown(cell)
            html.append(f"<td>{_esc(clean)}</td>")
        html.append("</tr>")

    html.append("</tbody></table>")
    return "\n".join(html)


def _pareto_table_to_html(text: str) -> str:
    """Convert Pareto table markdown to HTML with bold support for vital-few rows."""
    lines = [l.strip() for l in text.strip().split("\n") if l.strip()]
    table_lines = [l for l in lines if "|" in l and "---" not in l]
    if not table_lines:
        return f"<p>{_esc(text)}</p>"

    rows = []
    for line in table_lines:
        cells = [c.strip() for c in line.split("|") if c.strip()]
        rows.append(cells)

    if not rows:
        return ""

    html = ['<table class="data-table">']
    html.append("<thead><tr>")
    for cell in rows[0]:
        html.append(f"<th>{_esc(_strip_markdown(cell))}</th>")
    html.append("</tr></thead><tbody>")

    for row in rows[1:]:
        # Check if row has bold markers (vital few zone)
        is_bold = any(c.startswith("**") and c.endswith("**") for c in row)
        html.append("<tr>")
        for cell in row:
            clean = _strip_markdown(cell)
            if is_bold:
                html.append(f'<td><b>{_esc(clean)}</b></td>')
            else:
                html.append(f"<td>{_esc(clean)}</td>")
        html.append("</tr>")

    html.append("</tbody></table>")
    return "\n".join(html)


# ─── INDIVIDUAL REPORT ───────────────────────────────────────────────────────

def generate_report_html(
    doctor_name: str,
    doctor_code: str,
    period: str,
    report_sections: dict,
    specialty: str = "Medicina General",
) -> str:
    """Generate individual doctor audit report as HTML string."""
    informe_num = _build_informe_number(doctor_code, specialty, period)

    parts = [f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><style>{BASE_CSS}</style></head>
<body>
"""]

    # ── Cover table ──
    parts.append(f"""
<table class="cover-table">
<tr><td colspan="2" class="cover-banner">INFORME DE AUDITORÍA N° {_esc(informe_num)}</td></tr>
<tr class="cover-row"><td class="cover-label">NOMBRE</td><td>{_esc(doctor_name.upper())}</td></tr>
<tr class="cover-row"><td class="cover-label">CÓDIGO</td><td>{_esc(doctor_code)}</td></tr>
<tr class="cover-row"><td class="cover-label">ESPECIALIDAD</td><td>{_esc(specialty.upper())}</td></tr>
<tr class="cover-row"><td class="cover-label">DEPENDENCIA</td><td>Doctor SV - El Salvador</td></tr>
<tr class="cover-row"><td class="cover-label">PERIODO AUDITADO</td><td>{_esc(period)}</td></tr>
</table>
""")

    # ── Resumen Ejecutivo ──
    if report_sections.get("resumen_ejecutivo"):
        parts.append("<h1>RESUMEN EJECUTIVO</h1>")
        parts.append(_body_text_to_html(report_sections["resumen_ejecutivo"]))

    # ── Reporte de Cumplimiento ──
    if report_sections.get("cumplimiento"):
        parts.append('<div class="section-banner">REPORTE DE CUMPLIMIENTO POR CRITERIO</div>')
        parts.append('<p style="font-size:10pt;font-style:italic">'
                     'Nivel de cumplimiento por criterio evaluado, organizado por Criterio clínico. '
                     'Los porcentajes se calculan sobre el total de citas auditadas.</p>')
        parts.append(_compliance_table_to_html(report_sections["cumplimiento"]))

    # ── Comentario de Seguimiento ──
    if report_sections.get("seguimiento"):
        parts.append("<h1>COMENTARIO DE SEGUIMIENTO Y COMPARACIÓN DE PERÍODOS</h1>")
        parts.append(_body_text_to_html(report_sections["seguimiento"]))

    # ── Análisis de No Conformidades ──
    parts.append('<div class="page-break"></div>')

    if report_sections.get("cuantitativo"):
        parts.append("<h2>ANÁLISIS DE NO CONFORMIDADES</h2>")
        parts.append("<h2>ANÁLISIS CUANTITATIVO</h2>")
        parts.append(_quantitative_table_to_html(report_sections["cuantitativo"]))

    if report_sections.get("cualitativo") or report_sections.get("no_conformidades"):
        parts.append("<h2>ANÁLISIS CUALITATIVO</h2>")
        if report_sections.get("cualitativo"):
            parts.append(_body_text_to_html(report_sections["cualitativo"]))
        if report_sections.get("no_conformidades"):
            parts.append("<h2>Análisis de No Conformidades</h2>")
            parts.append(_body_text_to_html(report_sections["no_conformidades"]))
        parts.append('<hr class="separator"/>')
        if report_sections.get("eventos_riesgo"):
            parts.append("<h2>Análisis de Eventos de Riesgo</h2>")
            parts.append(_body_text_to_html(report_sections["eventos_riesgo"]))

    # ── Footer ──
    parts.append('<p class="footer">Documento generado automáticamente por el Sistema de Auditoría Médica de Calidad</p>')
    parts.append("</body></html>")

    return "\n".join(parts)


# ─── GENERAL PARETO REPORT ───────────────────────────────────────────────────

def _parse_general_report_sections(report_text: str) -> dict:
    """Parse Pareto general report into sections."""
    sections = {
        "tabla_pareto": "",
        "causas_vitales": "",
        "resumen_ejecutivo": "",
        "medicos_riesgo": "",
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
        stripped = line.strip()
        upper = stripped.upper()

        if "TABLA DE PARETO" in upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "tabla_pareto"
        elif "POCAS CAUSAS VITALES" in upper or ("CAUSAS VITALES" in upper and "POCAS" in upper):
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "causas_vitales"
        elif "RESUMEN EJECUTIVO" in upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "resumen_ejecutivo"
        elif "MÉDICOS EN RIESGO" in upper or "MEDICOS EN RIESGO" in upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "medicos_riesgo"
        else:
            buffer.append(line)

    if current_section and buffer:
        sections[current_section] = flush_buffer()

    return sections


def generate_general_report_html(
    report_text: str,
    specialty: str,
    period: str,
    chart_pareto: bytes | None = None,
    chart_accumulated: bytes | None = None,
    chart_nc_vs_er: bytes | None = None,
) -> str:
    """Generate general Pareto report as HTML string."""
    sections = _parse_general_report_sections(report_text)
    spec_abbrev = _get_specialty_abbrev(specialty)

    parts = [f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><style>{BASE_CSS}</style></head>
<body>
"""]

    # ── Cover table ──
    year_match = re.search(r"20\d{2}", period)
    year_str = year_match.group(0) if year_match else str(date.today().year)
    informe_id = f"PARETO-{spec_abbrev}-{year_str}"

    parts.append(f"""
<table class="cover-table">
<tr><td colspan="2" class="cover-banner">REPORTE GENERAL DE AUDITORÍA — {_esc(specialty.upper())}</td></tr>
<tr class="cover-row"><td class="cover-label">INFORME</td><td>{_esc(informe_id)}</td></tr>
<tr class="cover-row"><td class="cover-label">ESPECIALIDAD</td><td>{_esc(specialty.upper())}</td></tr>
<tr class="cover-row"><td class="cover-label">DEPENDENCIA</td><td>Doctor SV - El Salvador</td></tr>
<tr class="cover-row"><td class="cover-label">PERIODO</td><td>{_esc(period)}</td></tr>
<tr class="cover-row"><td class="cover-label">FECHA</td><td>{date.today().strftime("%d/%m/%Y")}</td></tr>
</table>
""")

    # ── Tabla de Pareto ──
    if sections.get("tabla_pareto"):
        parts.append(f"<h1>TABLA DE PARETO — {_esc(specialty.upper())}</h1>")
        parts.append(_pareto_table_to_html(sections["tabla_pareto"]))

    # ── Causas Vitales ──
    if sections.get("causas_vitales"):
        parts.append("<h2>POCAS CAUSAS VITALES (PUNTOS CRÍTICOS DE INTERVENCIÓN)</h2>")
        parts.append(_body_text_to_html(sections["causas_vitales"]))

    # ── Resumen Ejecutivo ──
    if sections.get("resumen_ejecutivo"):
        parts.append("<h1>RESUMEN EJECUTIVO</h1>")
        parts.append(_body_text_to_html(sections["resumen_ejecutivo"]))

    # ── Médicos en Riesgo ──
    if sections.get("medicos_riesgo"):
        parts.append("<h2>MÉDICOS EN RIESGO Y MEJOR EVALUADOS</h2>")
        parts.append(_body_text_to_html(sections["medicos_riesgo"]))

    # ── Charts ──
    if chart_pareto or chart_accumulated or chart_nc_vs_er:
        parts.append('<div class="page-break"></div>')
        parts.append("<h1>GRÁFICAS DE ANÁLISIS</h1>")

        if chart_pareto:
            parts.append("<h2>Diagrama de Pareto — Hallazgos del Período</h2>")
            b64 = base64.b64encode(chart_pareto).decode()
            parts.append(f'<img class="chart-img" src="data:image/png;base64,{b64}" style="width:90%"/>')

        if chart_accumulated:
            parts.append("<h2>Hallazgos Acumulados por Período</h2>")
            b64 = base64.b64encode(chart_accumulated).decode()
            parts.append(f'<img class="chart-img" src="data:image/png;base64,{b64}" style="width:90%"/>')

        if chart_nc_vs_er:
            parts.append("<h2>Eventos de Riesgo vs No Conformidades</h2>")
            b64 = base64.b64encode(chart_nc_vs_er).decode()
            parts.append(f'<img class="chart-img" src="data:image/png;base64,{b64}" style="width:80%"/>')

    # ── Footer ──
    parts.append('<p class="footer">Documento generado automáticamente por el Sistema de Auditoría Médica de Calidad</p>')
    parts.append("</body></html>")

    return "\n".join(parts)


# ─── PDF CONVERSION ──────────────────────────────────────────────────────────

def html_to_pdf(html: str) -> bytes:
    """Convert HTML string to PDF bytes using weasyprint."""
    from weasyprint import HTML
    pdf_bytes = HTML(string=html).write_pdf()
    return pdf_bytes


def generate_report_pdf(
    doctor_name: str,
    doctor_code: str,
    period: str,
    report_sections: dict,
    specialty: str = "Medicina General",
) -> bytes:
    """Generate individual doctor audit report as PDF bytes."""
    html = generate_report_html(
        doctor_name=doctor_name,
        doctor_code=doctor_code,
        period=period,
        report_sections=report_sections,
        specialty=specialty,
    )
    return html_to_pdf(html)


def generate_full_report_pdf_from_text(
    doctor_name: str,
    doctor_code: str,
    period: str,
    full_report_text: str,
    specialty: str = "Medicina General",
) -> bytes:
    """Generate individual report PDF from raw Claude output text."""
    from claude_analyzer import parse_report_sections
    sections = parse_report_sections(full_report_text)
    return generate_report_pdf(
        doctor_name=doctor_name,
        doctor_code=doctor_code,
        period=period,
        report_sections=sections,
        specialty=specialty,
    )


def generate_general_report_pdf(
    report_text: str,
    specialty: str,
    period: str,
    chart_pareto: bytes | None = None,
    chart_accumulated: bytes | None = None,
    chart_nc_vs_er: bytes | None = None,
) -> bytes:
    """Generate general Pareto report as PDF bytes."""
    html = generate_general_report_html(
        report_text=report_text,
        specialty=specialty,
        period=period,
        chart_pareto=chart_pareto,
        chart_accumulated=chart_accumulated,
        chart_nc_vs_er=chart_nc_vs_er,
    )
    return html_to_pdf(html)
