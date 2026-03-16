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


def _add_cover_table(doc: Document, doctor_name: str, doctor_code: str, period: str, specialty: str = "Medicina General"):
    """Add the report header/cover table."""
    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: Title
    row0 = table.rows[0]
    row0.cells[0].merge(row0.cells[1])
    cell = row0.cells[0]
    _set_cell_bg(cell, COLORS["header_bg"])
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("REPORTE DE AUDITORÍA MÉDICA DE CALIDAD")
    _set_run_format(run, bold=True, size=14, color=COLORS["header_text"])

    # Row 1: Doctor / Code
    labels = [("MÉDICO:", doctor_name), ("CÓDIGO:", doctor_code)]
    row1 = table.rows[1]
    for i, (label, value) in enumerate(labels):
        _set_cell_bg(row1.cells[i], COLORS["subsection_bg"])
        p = row1.cells[i].paragraphs[0]
        r1 = p.add_run(f"{label} ")
        _set_run_format(r1, bold=True, size=11)
        r2 = p.add_run(value)
        _set_run_format(r2, bold=False, size=11)

    # Row 2: Period / Specialty
    row2 = table.rows[2]
    for i, (label, value) in enumerate([("PERÍODO:", period), ("ESPECIALIDAD:", specialty)]):
        _set_cell_bg(row2.cells[i], COLORS["subsection_bg"])
        p = row2.cells[i].paragraphs[0]
        r1 = p.add_run(f"{label} ")
        _set_run_format(r1, bold=True, size=11)
        r2 = p.add_run(value)
        _set_run_format(r2, bold=False, size=11)

    # Row 3: Date
    row3 = table.rows[3]
    row3.cells[0].merge(row3.cells[1])
    _set_cell_bg(row3.cells[0], COLORS["subsection_bg"])
    p = row3.cells[0].paragraphs[0]
    r1 = p.add_run("FECHA DE GENERACIÓN: ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(date.today().strftime("%d/%m/%Y"))
    _set_run_format(r2, size=11)

    doc.add_paragraph()


def _add_section_banner(doc: Document, title: str):
    """Add a bold colored section banner paragraph."""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # Add shading via XML
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    hex_color = f"{COLORS['section_bg'].red:02X}{COLORS['section_bg'].green:02X}{COLORS['section_bg'].blue:02X}"
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
    table.style = "Table Grid"
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
    """Render compliance table with color-coded percentages."""
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
    table.style = "Table Grid"

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
                p = cell.paragraphs[0]
                # Check if this cell contains a percentage
                if "%" in cell_text:
                    color = _get_compliance_color(cell_text)
                    bg = color
                    _set_cell_bg(cell, bg)
                    run = p.add_run(cell_text)
                    _set_run_format(run, bold=True, size=10, color=COLORS["header_text"])
                else:
                    if i % 2 == 0:
                        _set_cell_bg(cell, COLORS["table_alt"])
                    run = p.add_run(cell_text)
                    _set_run_format(run, size=10)

    doc.add_paragraph()


def _add_body_text(doc: Document, text: str):
    """Add body text with basic markdown-like formatting."""
    lines = text.split("\n")
    for line in lines:
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph()
            continue

        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Handle **bold** markers
        parts = re.split(r"(\*\*[^*]+\*\*)", stripped)
        for part in parts:
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
) -> bytes:
    """
    Generate a formatted Word document from parsed report sections.
    Returns the document as bytes.
    """
    doc = Document()

    # Page margins
    section = doc.sections[0]
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    # ── COVER TABLE ────────────────────────────────────────────────────────────
    _add_cover_table(doc, doctor_name, doctor_code, period, specialty)

    # ── ANÁLISIS CUANTITATIVO ─────────────────────────────────────────────────
    if report_sections.get("cuantitativo"):
        _add_section_banner(doc, "ANÁLISIS CUANTITATIVO")
        _add_quantitative_table(doc, report_sections["cuantitativo"])

    # ── ANÁLISIS CUALITATIVO ──────────────────────────────────────────────────
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

    # ── RESUMEN EJECUTIVO ─────────────────────────────────────────────────────
    if report_sections.get("resumen_ejecutivo"):
        doc.add_page_break()
        _add_section_banner(doc, "RESUMEN EJECUTIVO")
        _add_body_text(doc, report_sections["resumen_ejecutivo"])

    # ── CUADRO DE CUMPLIMIENTO ────────────────────────────────────────────────
    if report_sections.get("cumplimiento"):
        _add_section_banner(doc, "REPORTE DE CUMPLIMIENTO POR CRITERIO")
        _add_compliance_table(doc, report_sections["cumplimiento"])

    # ── COLOR LEGEND ──────────────────────────────────────────────────────────
    _add_color_legend(doc)

    # ── COMENTARIO DE SEGUIMIENTO ─────────────────────────────────────────────
    if report_sections.get("seguimiento"):
        _add_section_banner(doc, "COMENTARIO DE SEGUIMIENTO Y COMPARACIÓN DE PERÍODOS")
        _add_body_text(doc, report_sections["seguimiento"])

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
) -> bytes:
    """
    Generate a report docx from raw Claude output text.
    Parses sections automatically.
    """
    from claude_analyzer import parse_report_sections
    sections = parse_report_sections(full_report_text)
    return generate_report_docx(
        doctor_name=doctor_name,
        doctor_code=doctor_code,
        period=period,
        report_sections=sections,
        specialty=specialty,
    )
