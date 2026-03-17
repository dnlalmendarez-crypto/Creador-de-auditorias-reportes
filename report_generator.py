"""
Report generator module.
Creates formatted Word (.docx) documents from audit report text.
Color palette and formatting follow the 'formato de informe' template.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io
from datetime import date


# ─── COLOR PALETTE ────────────────────────────────────────────────────────────
COLORS = {
    "header_bg": RGBColor(0x1F, 0x39, 0x64),       # Dark navy blue
    "header_text": RGBColor(0xFF, 0xFF, 0xFF),       # White
    "section_bg": RGBColor(0x2E, 0x74, 0xB5),        # Medium blue
    "section_text": RGBColor(0xFF, 0xFF, 0xFF),       # White
    "subsection_bg": RGBColor(0xBD, 0xD7, 0xEE),     # Light blue
    "subsection_text": RGBColor(0x00, 0x00, 0x00),   # Black
    "body_text": RGBColor(0x00, 0x00, 0x00),          # Black
    "table_header": RGBColor(0x1F, 0x39, 0x64),      # Dark navy
    "table_alt": RGBColor(0xD6, 0xE4, 0xF7),         # Very light blue
    "accent": RGBColor(0xED, 0x7D, 0x31),            # Orange accent
    # Compliance thresholds
    "rojo": RGBColor(0xFF, 0x00, 0x00),              # Red  <85%
    "anaranjado": RGBColor(0xFF, 0x66, 0x00),        # Orange 85-94%
    "amarillo": RGBColor(0xFF, 0xC0, 0x00),          # Yellow 95-97%
    "verde": RGBColor(0x00, 0xB0, 0x50),             # Green 98-100%
}

FONT_NAME = "Calibri"


def _set_table_style(table, style_name="Table Grid"):
    """Safely set a table style, falling back to no style if not available."""
    try:
        table.style = style_name
    except KeyError:
        pass


def _set_cell_bg(cell, rgb: RGBColor):
    """Set background color of a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    hex_color = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_run_format(run, bold=False, size=11, color: RGBColor = None, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.name = FONT_NAME
    if color:
        run.font.color.rgb = color


def _add_heading(doc: Document, text: str, level: int = 1):
    """Add a styled heading paragraph."""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = para.add_run(text)

    if level == 1:
        _set_run_format(run, bold=True, size=14, color=COLORS["header_bg"])
    elif level == 2:
        _set_run_format(run, bold=True, size=12, color=COLORS["section_bg"])
    elif level == 3:
        _set_run_format(run, bold=True, size=11, color=COLORS["body_text"])

    return para


def _get_specialty_abbrev(specialty: str) -> str:
    """Get specialty abbreviation for informe number."""
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


def _add_cover_table(doc: Document, doctor_name: str, doctor_code: str, period: str, specialty: str = "Medicina General"):
    """Add the report header/cover table matching the informe format."""
    # Build informe number
    spec_abbrev = _get_specialty_abbrev(specialty)
    # Extract year and period number from period string
    import re as _re
    year_match = _re.search(r"20\d{2}", period)
    year_str = year_match.group(0) if year_match else str(date.today().year)
    # Try to determine period number from the period string
    period_num = "001"
    month_map = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
        "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
        "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
        "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
    }
    period_lower = period.lower()
    for m_name, m_num in month_map.items():
        if m_name in period_lower:
            # 2 quincenas per month
            q_num = (m_num - 1) * 2 + 1
            if "16" in period or "segunda" in period_lower:
                q_num += 1
            period_num = f"{q_num:03d}"
            break

    informe_num = f"{doctor_code}-{spec_abbrev}-{year_str}-P{period_num}"

    table = doc.add_table(rows=7, cols=2)
    _set_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: INFORME DE AUDITORÍA N°
    row0 = table.rows[0]
    row0.cells[0].merge(row0.cells[1])
    cell = row0.cells[0]
    _set_cell_bg(cell, COLORS["header_bg"])
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"INFORME DE AUDITORÍA N° {informe_num}")
    _set_run_format(run, bold=True, size=14, color=COLORS["header_text"])

    # Row 1: Separator
    row1 = table.rows[1]
    row1.cells[0].merge(row1.cells[1])
    _set_cell_bg(row1.cells[0], COLORS["subsection_bg"])

    # Row 2: NOMBRE
    row2 = table.rows[2]
    row2.cells[0].merge(row2.cells[1])
    _set_cell_bg(row2.cells[0], COLORS["subsection_bg"])
    p = row2.cells[0].paragraphs[0]
    r1 = p.add_run("NOMBRE           ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(doctor_name.upper())
    _set_run_format(r2, size=11)

    # Row 3: CÓDIGO / ESPECIALIDAD
    row3 = table.rows[3]
    _set_cell_bg(row3.cells[0], COLORS["subsection_bg"])
    p = row3.cells[0].paragraphs[0]
    r1 = p.add_run("CÓDIGO           ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(doctor_code)
    _set_run_format(r2, size=11)
    _set_cell_bg(row3.cells[1], COLORS["subsection_bg"])

    # Row 4: ESPECIALIDAD
    row4 = table.rows[4]
    row4.cells[0].merge(row4.cells[1])
    _set_cell_bg(row4.cells[0], COLORS["subsection_bg"])
    p = row4.cells[0].paragraphs[0]
    r1 = p.add_run("ESPECIALIDAD     ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(specialty.upper())
    _set_run_format(r2, size=11)

    # Row 5: DEPENDENCIA
    row5 = table.rows[5]
    row5.cells[0].merge(row5.cells[1])
    _set_cell_bg(row5.cells[0], COLORS["subsection_bg"])
    p = row5.cells[0].paragraphs[0]
    r1 = p.add_run("DEPENDENCIA      ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run("Doctor SV - El Salvador")
    _set_run_format(r2, size=11)

    # Row 6: PERIODO AUDITADO
    row6 = table.rows[6]
    row6.cells[0].merge(row6.cells[1])
    _set_cell_bg(row6.cells[0], COLORS["subsection_bg"])
    p = row6.cells[0].paragraphs[0]
    r1 = p.add_run("PERIODO AUDITADO ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(period)
    _set_run_format(r2, size=11)

    doc.add_paragraph()


def _add_section_banner(doc: Document, title: str):
    """Add a bold colored section banner paragraph."""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # Add shading via XML
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    hex_color = f"{COLORS['section_bg'][0]:02X}{COLORS['section_bg'][1]:02X}{COLORS['section_bg'][2]:02X}"
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)
    run = para.add_run(f"  {title}  ")
    _set_run_format(run, bold=True, size=12, color=COLORS["header_text"])
    return para


def _add_quantitative_table(doc: Document, table_text: str):
    """Parse and render the quantitative analysis table."""
    lines = [l.strip() for l in table_text.strip().split("\n") if l.strip()]
    # Filter markdown table lines
    table_lines = [l for l in lines if "|" in l and "---" not in l]

    if not table_lines:
        doc.add_paragraph(table_text)
        return

    rows_data = []
    for line in table_lines:
        cells = [c.strip() for c in line.split("|") if c.strip()]
        rows_data.append(cells)

    if not rows_data:
        return

    num_cols = max(len(r) for r in rows_data)
    table = doc.add_table(rows=len(rows_data), cols=num_cols)
    _set_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, row_data in enumerate(rows_data):
        row = table.rows[i]
        for j, cell_text in enumerate(row_data):
            if j >= num_cols:
                break
            cell = row.cells[j]
            if i == 0:
                _set_cell_bg(cell, COLORS["table_header"])
                p = cell.paragraphs[0]
                run = p.add_run(cell_text)
                _set_run_format(run, bold=True, size=10, color=COLORS["header_text"])
            else:
                if i % 2 == 0:
                    _set_cell_bg(cell, COLORS["table_alt"])
                p = cell.paragraphs[0]
                run = p.add_run(cell_text)
                _set_run_format(run, size=10)

    doc.add_paragraph()


def _get_compliance_color(percentage_str: str) -> RGBColor:
    """Return color based on compliance percentage string."""
    try:
        val_str = percentage_str.replace("%", "").replace(",", ".").strip()
        val = float(val_str)
        if val < 85:
            return COLORS["rojo"]
        elif val < 95:
            return COLORS["anaranjado"]
        elif val < 98:
            return COLORS["amarillo"]
        else:
            return COLORS["verde"]
    except Exception:
        return COLORS["body_text"]


def _add_compliance_table(doc: Document, compliance_text: str):
    """Render compliance table with COMPONENTE/CRITERIO grouped format and color-coded percentages."""
    lines = [l.strip() for l in compliance_text.strip().split("\n") if l.strip()]
    table_lines = [l for l in lines if "|" in l and "---" not in l]

    if not table_lines:
        doc.add_paragraph(compliance_text)
        return

    rows_data = []
    for line in table_lines:
        cells = [c.strip() for c in line.split("|") if c.strip()]
        rows_data.append(cells)

    if not rows_data:
        return

    num_cols = max(len(r) for r in rows_data)
    table = doc.add_table(rows=len(rows_data), cols=num_cols)
    _set_table_style(table)

    # ── Header row ────────────────────────────────────────────────────────────
    header_row = table.rows[0]
    for j, cell_text in enumerate(rows_data[0]):
        if j >= num_cols:
            break
        cell = header_row.cells[j]
        _set_cell_bg(cell, COLORS["table_header"])
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(cell_text)
        _set_run_format(run, bold=True, size=9, color=COLORS["header_text"])

    # ── Data rows ─────────────────────────────────────────────────────────────
    for i in range(1, len(rows_data)):
        row_data = rows_data[i]
        row = table.rows[i]
        for j, cell_text in enumerate(row_data):
            if j >= num_cols:
                break
            cell = row.cells[j]
            p = cell.paragraphs[0]

            if j == 0:
                # COMPONENTE column — dark background, bold white
                _set_cell_bg(cell, COLORS["header_bg"])
                run = p.add_run(cell_text)
                _set_run_format(run, bold=True, size=9, color=COLORS["header_text"])
            elif j == 1:
                # CRITERIO column — light alternating
                if i % 2 == 0:
                    _set_cell_bg(cell, COLORS["table_alt"])
                run = p.add_run(cell_text)
                _set_run_format(run, size=9)
            else:
                # Period CUMPLIMIENTO columns — color coded
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if "%" in cell_text:
                    color = _get_compliance_color(cell_text)
                    _set_cell_bg(cell, color)
                    run = p.add_run(cell_text)
                    _set_run_format(run, bold=True, size=9, color=COLORS["header_text"])
                elif cell_text.strip() == "-":
                    if i % 2 == 0:
                        _set_cell_bg(cell, COLORS["table_alt"])
                    run = p.add_run(cell_text)
                    _set_run_format(run, size=9, italic=True)
                else:
                    if i % 2 == 0:
                        _set_cell_bg(cell, COLORS["table_alt"])
                    run = p.add_run(cell_text)
                    _set_run_format(run, size=9)

    # ── Merge COMPONENTE cells vertically for same-component rows ─────────
    if len(rows_data) > 1 and num_cols > 0:
        comp_col = 0
        i = 1
        while i < len(rows_data):
            comp_name = rows_data[i][comp_col] if comp_col < len(rows_data[i]) else ""
            if not comp_name:
                i += 1
                continue
            # Find consecutive rows with same component
            j = i + 1
            while j < len(rows_data):
                next_comp = rows_data[j][comp_col] if comp_col < len(rows_data[j]) else ""
                if next_comp == comp_name:
                    j += 1
                else:
                    break
            # Merge if more than one row
            if j - i > 1:
                start_cell = table.cell(i, comp_col)
                end_cell = table.cell(j - 1, comp_col)
                start_cell.merge(end_cell)
                # Re-apply formatting to merged cell
                _set_cell_bg(start_cell, COLORS["header_bg"])
                p = start_cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # Set vertical alignment to center
                tc = start_cell._tc
                tcPr = tc.get_or_add_tcPr()
                vAlign = OxmlElement("w:vAlign")
                vAlign.set(qn("w:val"), "center")
                tcPr.append(vAlign)
                # Clear and re-add text
                for run in p.runs:
                    run.clear()
                run = p.add_run(comp_name)
                _set_run_format(run, bold=True, size=9, color=COLORS["header_text"])
            i = j

    doc.add_paragraph()


def _add_body_text(doc: Document, text: str):
    """Add body text with basic markdown-like formatting."""
    lines = text.split("\n")
    for line in lines:
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph()
            continue

        # Separator lines (─── or ___)
        if all(c in "─_═━" for c in stripped) and len(stripped) > 10:
            _add_separator(doc)
            continue

        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # First handle [text]{.underline} markers, then **bold**
        # Process: split by underline markers first, then bold within each part
        underline_parts = re.split(r"(\[[^\]]+\]\{\.underline\})", stripped)
        for u_part in underline_parts:
            underline_match = re.match(r"\[([^\]]+)\]\{\.underline\}", u_part)
            if underline_match:
                run = para.add_run(underline_match.group(1))
                _set_run_format(run, bold=True, size=11)
                run.underline = True
            else:
                # Handle **bold** markers within non-underline parts
                bold_parts = re.split(r"(\*\*[^*]+\*\*)", u_part)
                for part in bold_parts:
                    if part.startswith("**") and part.endswith("**"):
                        run = para.add_run(part[2:-2])
                        _set_run_format(run, bold=True, size=11)
                    else:
                        run = para.add_run(part)
                        _set_run_format(run, size=11)


def _add_separator(doc: Document):
    para = doc.add_paragraph("_" * 80)
    run = para.runs[0]
    _set_run_format(run, size=10, color=COLORS["section_bg"])


def generate_report_docx(
    doctor_name: str,
    doctor_code: str,
    period: str,
    report_sections: dict,
    specialty: str = "Medicina General",
    template_bytes: bytes | None = None,
) -> bytes:
    """
    Generate a formatted Word document from parsed report sections.
    If template_bytes is provided, uses the .docx template as a base
    (preserving its styles, headers, footers, etc.).
    Returns the document as bytes.
    """
    if template_bytes:
        doc = Document(io.BytesIO(template_bytes))
        # Remove all existing body content from the template, keeping only
        # styles, headers, footers, and page setup as format reference.
        body = doc.element.body
        for child in list(body):
            if child.tag.endswith("}sectPr"):
                continue  # preserve section properties (margins, headers, footers)
            body.remove(child)
    else:
        doc = Document()

    # Page margins (only set on blank docs to not override template)
    if not template_bytes:
        section = doc.sections[0]
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)

    # ── COVER TABLE ────────────────────────────────────────────────────────────
    _add_cover_table(doc, doctor_name, doctor_code, period, specialty)

    # ── 1. RESUMEN EJECUTIVO ─────────────────────────────────────────────────
    if report_sections.get("resumen_ejecutivo"):
        _add_section_banner(doc, "RESUMEN EJECUTIVO")
        _add_body_text(doc, report_sections["resumen_ejecutivo"])

    # ── 2. REPORTE DE CUMPLIMIENTO POR CRITERIO ─────────────────────────────
    if report_sections.get("cumplimiento"):
        _add_section_banner(doc, "REPORTE DE CUMPLIMIENTO POR CRITERIO")
        # Add description
        desc_para = doc.add_paragraph()
        desc_run = desc_para.add_run(
            "Nivel de cumplimiento por criterio evaluado, organizado por Criterio clínico. "
            "Los porcentajes se calculan sobre el total de citas auditadas."
        )
        _set_run_format(desc_run, size=10, italic=True)
        _add_compliance_table(doc, report_sections["cumplimiento"])
        _add_color_legend(doc)

    # ── 3. COMENTARIO DE SEGUIMIENTO ─────────────────────────────────────────
    if report_sections.get("seguimiento"):
        _add_section_banner(doc, "COMENTARIO DE SEGUIMIENTO Y COMPARACIÓN DE PERÍODOS")
        _add_body_text(doc, report_sections["seguimiento"])

    # ── 4. ANÁLISIS DE NO CONFORMIDADES ────────────────────────────────────
    doc.add_page_break()

    # ── 4a. ANÁLISIS CUANTITATIVO ──────────────────────────────────────────
    if report_sections.get("cuantitativo"):
        _add_section_banner(doc, "ANÁLISIS DE NO CONFORMIDADES")
        doc.add_paragraph()
        _add_section_banner(doc, "ANÁLISIS CUANTITATIVO")
        _add_quantitative_table(doc, report_sections["cuantitativo"])

    # ── 4b. ANÁLISIS CUALITATIVO ───────────────────────────────────────────
    if report_sections.get("cualitativo") or report_sections.get("no_conformidades"):
        _add_section_banner(doc, "ANÁLISIS CUALITATIVO")

        if report_sections.get("cualitativo"):
            _add_body_text(doc, report_sections["cualitativo"])

        if report_sections.get("no_conformidades"):
            _add_heading(doc, "Análisis de No Conformidades", level=2)
            _add_body_text(doc, report_sections["no_conformidades"])

        _add_separator(doc)

        if report_sections.get("eventos_riesgo"):
            _add_heading(doc, "Análisis de Eventos de Riesgo", level=2)
            _add_body_text(doc, report_sections["eventos_riesgo"])

    # ── FOOTER NOTE ───────────────────────────────────────────────────────────
    doc.add_paragraph()
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = note.add_run("Documento generado automáticamente por el Sistema de Auditoría Médica de Calidad")
    _set_run_format(run, size=9, italic=True, color=RGBColor(0x7F, 0x7F, 0x7F))

    # Save to bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _add_color_legend(doc: Document):
    """Add the color coding legend for compliance percentages."""
    doc.add_paragraph()
    legend_para = doc.add_paragraph()
    legend_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = legend_para.add_run("Leyenda de colores: ")
    _set_run_format(run, bold=True, size=10)

    legend_items = [
        ("ROJO: <85% (Oportunidad de Mejora)", COLORS["rojo"]),
        ("  |  ANARANJADO: 85-94% (Aceptable)", COLORS["anaranjado"]),
        ("  |  AMARILLO: 95-97% (Muy Bueno)", COLORS["amarillo"]),
        ("  |  VERDE: 98-100% (Excelente)", COLORS["verde"]),
    ]
    for text, color in legend_items:
        run = legend_para.add_run(text)
        _set_run_format(run, bold=True, size=10, color=color)

    doc.add_paragraph()


def generate_full_report_from_text(
    doctor_name: str,
    doctor_code: str,
    period: str,
    full_report_text: str,
    specialty: str = "Medicina General",
    template_bytes: bytes | None = None,
) -> bytes:
    """
    Generate a report docx from raw Claude output text.
    Parses sections automatically.
    If template_bytes is provided, uses the .docx as a base template.
    """
    from claude_analyzer import parse_report_sections
    sections = parse_report_sections(full_report_text)
    return generate_report_docx(
        doctor_name=doctor_name,
        doctor_code=doctor_code,
        period=period,
        report_sections=sections,
        specialty=specialty,
        template_bytes=template_bytes,
    )


def generate_general_report_docx(
    report_text: str,
    specialty: str,
    period: str,
    chart_pareto: bytes | None = None,
    chart_accumulated: bytes | None = None,
    chart_nc_vs_er: bytes | None = None,
    template_bytes: bytes | None = None,
) -> bytes:
    """
    Generate a general Pareto report Word document with embedded charts.
    """
    if template_bytes:
        doc = Document(io.BytesIO(template_bytes))
        body = doc.element.body
        for child in list(body):
            if child.tag.endswith("}sectPr"):
                continue
            body.remove(child)
    else:
        doc = Document()
        section = doc.sections[0]
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)

    # Cover table
    table = doc.add_table(rows=3, cols=2)
    _set_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    row0 = table.rows[0]
    row0.cells[0].merge(row0.cells[1])
    _set_cell_bg(row0.cells[0], COLORS["header_bg"])
    p = row0.cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("REPORTE GENERAL — ANÁLISIS DE PARETO")
    _set_run_format(run, bold=True, size=14, color=COLORS["header_text"])

    row1 = table.rows[1]
    _set_cell_bg(row1.cells[0], COLORS["subsection_bg"])
    p = row1.cells[0].paragraphs[0]
    r1 = p.add_run("ESPECIALIDAD: ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(specialty)
    _set_run_format(r2, size=11)
    _set_cell_bg(row1.cells[1], COLORS["subsection_bg"])
    p = row1.cells[1].paragraphs[0]
    r1 = p.add_run("PERÍODO: ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(period)
    _set_run_format(r2, size=11)

    row2 = table.rows[2]
    row2.cells[0].merge(row2.cells[1])
    _set_cell_bg(row2.cells[0], COLORS["subsection_bg"])
    p = row2.cells[0].paragraphs[0]
    from datetime import date as _date
    r1 = p.add_run("FECHA DE GENERACIÓN: ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(_date.today().strftime("%d/%m/%Y"))
    _set_run_format(r2, size=11)

    doc.add_paragraph()

    # Report body text
    _add_body_text(doc, report_text)

    # Charts section
    if chart_pareto or chart_accumulated or chart_nc_vs_er:
        doc.add_page_break()
        _add_section_banner(doc, "GRÁFICAS DE ANÁLISIS")

    if chart_pareto:
        doc.add_paragraph()
        _add_heading(doc, "Diagrama de Pareto — Hallazgos del Período", level=2)
        doc.add_picture(io.BytesIO(chart_pareto), width=Inches(6.0))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if chart_accumulated:
        doc.add_paragraph()
        _add_heading(doc, "Hallazgos Acumulados por Período", level=2)
        doc.add_picture(io.BytesIO(chart_accumulated), width=Inches(6.0))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if chart_nc_vs_er:
        doc.add_paragraph()
        _add_heading(doc, "Eventos de Riesgo vs No Conformidades", level=2)
        doc.add_picture(io.BytesIO(chart_nc_vs_er), width=Inches(5.0))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Footer
    doc.add_paragraph()
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = note.add_run("Documento generado automáticamente por el Sistema de Auditoría Médica de Calidad")
    _set_run_format(run, size=9, italic=True, color=RGBColor(0x7F, 0x7F, 0x7F))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
