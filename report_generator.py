"""
Report generator module.
Creates formatted Word (.docx) documents from audit report text.
Color palette and formatting are extracted from the uploaded template.
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


# ─── DEFAULT COLOR PALETTE (used when no template is provided) ───────────────
DEFAULT_COLORS = {
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

DEFAULT_FONT = "Calibri"

# Module-level active styles (overridden when template is loaded)
COLORS = dict(DEFAULT_COLORS)
FONT_NAME = DEFAULT_FONT


def _hex_to_rgb(hex_str: str) -> RGBColor | None:
    """Convert a 6-char hex string to RGBColor, or None if invalid."""
    hex_str = hex_str.strip().lstrip("#")
    if len(hex_str) != 6:
        return None
    try:
        r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
        return RGBColor(r, g, b)
    except ValueError:
        return None


def _get_cell_fill(cell) -> str | None:
    """Extract the fill hex color from a cell's XML shading."""
    tc = cell._tc
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is not None:
        shd = tcPr.find(qn("w:shd"))
        if shd is not None:
            fill = shd.get(qn("w:fill"))
            if fill and fill.lower() != "auto":
                return fill
    return None


def _get_run_color(run) -> str | None:
    """Extract font color hex from a run's XML."""
    if run.font.color and run.font.color.rgb:
        rgb = run.font.color.rgb
        return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    # Check XML directly
    rPr = run._r.find(qn("w:rPr"))
    if rPr is not None:
        color_el = rPr.find(qn("w:color"))
        if color_el is not None:
            val = color_el.get(qn("w:val"))
            if val and val.lower() != "auto":
                return val
    return None


def _get_paragraph_shading(para) -> str | None:
    """Extract paragraph background shading hex."""
    pPr = para._p.find(qn("w:pPr"))
    if pPr is not None:
        shd = pPr.find(qn("w:shd"))
        if shd is not None:
            fill = shd.get(qn("w:fill"))
            if fill and fill.lower() != "auto":
                return fill
    return None


def extract_template_styles(template_bytes: bytes) -> dict:
    """
    Extract colors and font from a template .docx file.
    Returns a dict with 'colors' (dict) and 'font_name' (str).

    Strategy:
    - Scan all tables: first row cell fills -> header_bg, text color -> header_text
    - Second/third rows -> subsection_bg
    - Scan paragraphs with shading -> section_bg, text color -> section_text
    - Scan body text runs -> body_text color, font name
    - Alternating row colors in tables -> table_alt
    """
    doc = Document(io.BytesIO(template_bytes))

    extracted = {}
    font_name = DEFAULT_FONT

    # ── Extract from tables ──────────────────────────────────────────────
    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue

        # First row = header
        for cell in rows[0].cells:
            fill = _get_cell_fill(cell)
            if fill:
                rgb = _hex_to_rgb(fill)
                if rgb and "header_bg" not in extracted:
                    extracted["header_bg"] = rgb
                    extracted["table_header"] = rgb

            for para in cell.paragraphs:
                for run in para.runs:
                    rc = _get_run_color(run)
                    if rc:
                        rgb = _hex_to_rgb(rc)
                        if rgb and "header_text" not in extracted:
                            extracted["header_text"] = rgb
                            extracted["section_text"] = rgb
                    if run.font.name and run.font.name.strip():
                        font_name = run.font.name.strip()

        # Rows 1-3 = subsection area
        for row_idx in range(1, min(4, len(rows))):
            for cell in rows[row_idx].cells:
                fill = _get_cell_fill(cell)
                if fill:
                    rgb = _hex_to_rgb(fill)
                    if rgb and "subsection_bg" not in extracted:
                        extracted["subsection_bg"] = rgb

        # Later rows = look for alternating colors
        for row_idx in range(2, len(rows)):
            for cell in rows[row_idx].cells:
                fill = _get_cell_fill(cell)
                if fill:
                    rgb = _hex_to_rgb(fill)
                    if rgb and rgb != extracted.get("header_bg") and rgb != extracted.get("subsection_bg"):
                        if "table_alt" not in extracted:
                            extracted["table_alt"] = rgb

    # ── Extract from paragraphs ──────────────────────────────────────────
    for para in doc.paragraphs:
        shd = _get_paragraph_shading(para)
        if shd:
            rgb = _hex_to_rgb(shd)
            if rgb and "section_bg" not in extracted:
                # Section banners have shaded backgrounds
                if rgb != extracted.get("header_bg") and rgb != extracted.get("subsection_bg"):
                    extracted["section_bg"] = rgb

        for run in para.runs:
            if run.font.name and run.font.name.strip():
                font_name = run.font.name.strip()

            rc = _get_run_color(run)
            if rc:
                rgb = _hex_to_rgb(rc)
                if rgb:
                    # Body text color (non-white, non-header)
                    if rgb != extracted.get("header_text") and "body_text" not in extracted:
                        extracted["body_text"] = rgb

    # Build final colors dict: extracted values override defaults
    colors = dict(DEFAULT_COLORS)
    colors.update(extracted)

    return {"colors": colors, "font_name": font_name}


def apply_template_styles(template_bytes: bytes | None):
    """
    Extract styles from template and set them as module-level active styles.
    Call this ONCE when the template is loaded.
    """
    global COLORS, FONT_NAME
    if template_bytes:
        styles = extract_template_styles(template_bytes)
        COLORS = styles["colors"]
        FONT_NAME = styles["font_name"]
    else:
        COLORS = dict(DEFAULT_COLORS)
        FONT_NAME = DEFAULT_FONT


def _fix_sectpr_margins(doc: Document):
    """
    Sanitize section margin values in the template XML.
    Some templates produce float twip values (e.g. '1115.669...')
    which python-docx cannot parse as int. Convert them to int.
    """
    for sectPr in doc.element.body.iter(qn("w:sectPr")):
        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            for attr in ("left", "right", "top", "bottom",
                         "header", "footer", "gutter"):
                full_attr = qn("w:" + attr)
                val = pgMar.get(full_attr)
                if val is not None:
                    try:
                        int(val)
                    except ValueError:
                        pgMar.set(full_attr, str(int(float(val))))


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
    """Add a bold colored section banner with background shading (e.g. REPORTE DE CUMPLIMIENTO)."""
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


def _add_section_title(doc: Document, title: str, centered: bool = True, underline: bool = True, size: int = 14):
    """Add a section title without background — bold, dark blue, optionally centered and underlined.
    Used for RESUMEN EJECUTIVO, ANÁLISIS DE NO CONFORMIDADES, etc."""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT
    run = para.add_run(title)
    _set_run_format(run, bold=True, size=size, color=COLORS["header_bg"])
    if underline:
        run.underline = True
    return para


def _add_subsection_title(doc: Document, title: str, size: int = 12):
    """Add a subsection title — bold, dark blue, left-aligned, no background.
    Used for ANÁLISIS CUANTITATIVO, ANÁLISIS CUALITATIVO, etc."""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = para.add_run(title)
    _set_run_format(run, bold=True, size=size, color=COLORS["header_bg"])
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
    """
    Add body text converting markdown formatting to proper Word formatting.
    Handles: headings (# ## ###), bold (**text**), underline ([text]{.underline}),
    numbered lists, bullet points, separators, and plain text.
    """
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

        # ── Markdown headings → Word headings ────────────────────────────
        heading_match = re.match(r"^(#{1,4})\s+(.+)$", stripped)
        if heading_match:
            level = len(heading_match.group(1))
            heading_text = _strip_markdown(heading_match.group(2))
            _add_heading(doc, heading_text, level=min(level, 3))
            continue

        # ── Numbered list items (e.g. "1. ANAMNESIS") ────────────────────
        # Keep the number but render as bold heading if ALL CAPS
        num_match = re.match(r"^(\d+\.)\s+(.+)$", stripped)
        if num_match and num_match.group(2).replace(" ", "").isupper():
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            heading_text = _strip_markdown(f"{num_match.group(1)} {num_match.group(2)}")
            run = para.add_run(heading_text)
            _set_run_format(run, bold=True, size=12, color=COLORS["header_bg"])
            continue

        # ── Bullet points (- or •) ───────────────────────────────────────
        bullet_match = re.match(r"^[-•]\s+(.+)$", stripped)
        if bullet_match:
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _add_formatted_runs(para, "• " + bullet_match.group(1))
            continue

        # ── Regular paragraph ─────────────────────────────────────────────
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _add_formatted_runs(para, stripped)


def _strip_markdown(text: str) -> str:
    """Remove all markdown formatting characters, returning clean text."""
    # Remove [text]{.underline} → text
    text = re.sub(r"\[([^\]]+)\]\{\.underline\}", r"\1", text)
    # Remove **bold** → bold
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    # Remove *italic* → italic
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    # Remove remaining stray * or #
    text = text.replace("*", "").strip()
    return text


def _add_formatted_runs(para, text: str):
    """
    Parse inline markdown in text and add properly formatted runs to paragraph.
    Handles **bold**, [text]{.underline}, and plain text.
    """
    # Split by underline markers and bold markers
    # Pattern order: underline first, then bold
    pattern = re.compile(r"(\[[^\]]+\]\{\.underline\}|\*\*[^*]+\*\*)")
    parts = pattern.split(text)

    for part in parts:
        if not part:
            continue
        underline_match = re.match(r"\[([^\]]+)\]\{\.underline\}", part)
        if underline_match:
            run = para.add_run(underline_match.group(1))
            _set_run_format(run, bold=True, size=11)
            run.underline = True
            continue
        if part.startswith("**") and part.endswith("**"):
            run = para.add_run(part[2:-2])
            _set_run_format(run, bold=True, size=11)
            continue
        # Clean any remaining stray markdown
        clean = part.replace("*", "")
        if clean:
            run = para.add_run(clean)
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
        _fix_sectpr_margins(doc)
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
        _add_section_title(doc, "RESUMEN EJECUTIVO", centered=True, underline=True)
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
        _add_section_title(doc, "COMENTARIO DE SEGUIMIENTO Y COMPARACIÓN DE PERÍODOS",
                           centered=True, underline=True)
        _add_body_text(doc, report_sections["seguimiento"])

    # ── 4. ANÁLISIS DE NO CONFORMIDADES ────────────────────────────────────
    doc.add_page_break()

    # ── 4a. ANÁLISIS CUANTITATIVO ──────────────────────────────────────────
    if report_sections.get("cuantitativo"):
        _add_section_title(doc, "ANÁLISIS DE NO CONFORMIDADES",
                           centered=False, underline=False, size=14)
        doc.add_paragraph()
        _add_subsection_title(doc, "ANÁLISIS CUANTITATIVO", size=12)
        _add_quantitative_table(doc, report_sections["cuantitativo"])

    # ── 4b. ANÁLISIS CUALITATIVO ───────────────────────────────────────────
    if report_sections.get("cualitativo") or report_sections.get("no_conformidades"):
        _add_subsection_title(doc, "ANÁLISIS CUALITATIVO", size=12)

        if report_sections.get("cualitativo"):
            _add_body_text(doc, report_sections["cualitativo"])

        if report_sections.get("no_conformidades"):
            _add_subsection_title(doc, "Análisis de No Conformidades", size=12)
            _add_body_text(doc, report_sections["no_conformidades"])

        _add_separator(doc)

        if report_sections.get("eventos_riesgo"):
            _add_subsection_title(doc, "Análisis de Eventos de Riesgo", size=12)
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


def _add_general_cover_table(doc: Document, specialty: str, period: str):
    """Add the general report header/cover table — same style as individual reports."""
    table = doc.add_table(rows=5, cols=2)
    _set_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: Title banner
    row0 = table.rows[0]
    row0.cells[0].merge(row0.cells[1])
    _set_cell_bg(row0.cells[0], COLORS["header_bg"])
    p = row0.cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("REPORTE GENERAL — ANÁLISIS DE PARETO")
    _set_run_format(run, bold=True, size=14, color=COLORS["header_text"])

    # Row 1: Separator
    row1 = table.rows[1]
    row1.cells[0].merge(row1.cells[1])
    _set_cell_bg(row1.cells[0], COLORS["subsection_bg"])

    # Row 2: ESPECIALIDAD
    row2 = table.rows[2]
    row2.cells[0].merge(row2.cells[1])
    _set_cell_bg(row2.cells[0], COLORS["subsection_bg"])
    p = row2.cells[0].paragraphs[0]
    r1 = p.add_run("ESPECIALIDAD     ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(specialty.upper())
    _set_run_format(r2, size=11)

    # Row 3: PERIODO AUDITADO
    row3 = table.rows[3]
    row3.cells[0].merge(row3.cells[1])
    _set_cell_bg(row3.cells[0], COLORS["subsection_bg"])
    p = row3.cells[0].paragraphs[0]
    r1 = p.add_run("PERIODO AUDITADO ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(period)
    _set_run_format(r2, size=11)

    # Row 4: DEPENDENCIA + FECHA
    row4 = table.rows[4]
    _set_cell_bg(row4.cells[0], COLORS["subsection_bg"])
    p = row4.cells[0].paragraphs[0]
    r1 = p.add_run("DEPENDENCIA      ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run("Doctor SV - El Salvador")
    _set_run_format(r2, size=11)
    _set_cell_bg(row4.cells[1], COLORS["subsection_bg"])
    p = row4.cells[1].paragraphs[0]
    r1 = p.add_run("FECHA  ")
    _set_run_format(r1, bold=True, size=11)
    r2 = p.add_run(date.today().strftime("%d/%m/%Y"))
    _set_run_format(r2, size=11)

    doc.add_paragraph()


def _parse_general_report_sections(report_text: str) -> dict:
    """
    Parse the Pareto general report into sections for structured document generation.
    Returns dict with keys: tabla_pareto, causas_vitales, resumen_ejecutivo,
    medicos_riesgo, full_text.
    """
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
        elif "POCAS CAUSAS VITALES" in upper or "CAUSAS VITALES" in upper and "POCAS" in upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "causas_vitales"
        elif "RESUMEN EJECUTIVO" in upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "resumen_ejecutivo"
        elif "MÉDICOS EN RIESGO" in upper or "MEDICOS EN RIESGO" in upper or "MEJOR EVALUADOS" in upper:
            if current_section:
                sections[current_section] = flush_buffer()
            current_section = "medicos_riesgo"
        else:
            buffer.append(line)

    if current_section and buffer:
        sections[current_section] = flush_buffer()

    return sections


def _add_pareto_table(doc: Document, table_text: str):
    """Parse and render the Pareto analysis table with same style as compliance/quantitative tables."""
    lines = [l.strip() for l in table_text.strip().split("\n") if l.strip()]
    table_lines = [l for l in lines if l.startswith("|") and "---" not in l]
    if not table_lines:
        _add_body_text(doc, table_text)
        return

    # Parse header + rows
    header_cells = [c.strip() for c in table_lines[0].split("|") if c.strip()]
    data_rows = []
    for tl in table_lines[1:]:
        cells = [c.strip() for c in tl.split("|") if c.strip()]
        if cells:
            data_rows.append(cells)

    ncols = len(header_cells)
    table = doc.add_table(rows=1 + len(data_rows), cols=ncols)
    _set_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for j, hdr in enumerate(header_cells):
        cell = table.cell(0, j)
        _set_cell_bg(cell, COLORS["table_header"])
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(hdr)
        _set_run_format(run, bold=True, size=9, color=COLORS["header_text"])

    # Data rows
    for i, row_data in enumerate(data_rows):
        for j in range(min(len(row_data), ncols)):
            cell = table.cell(i + 1, j)
            if i % 2 == 1:
                _set_cell_bg(cell, COLORS["table_alt"])
            p = cell.paragraphs[0]
            # Check if this row contains accumulated % > 80 (vital few zone)
            is_bold = row_data[j].startswith("**") and row_data[j].endswith("**")
            text = row_data[j].strip("*")
            run = p.add_run(text)
            _set_run_format(run, bold=is_bold, size=9)

    doc.add_paragraph()


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
    Uses the same structural pattern and template styles as individual reports.
    """
    if template_bytes:
        doc = Document(io.BytesIO(template_bytes))
        _fix_sectpr_margins(doc)
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

    # Parse report into sections
    sections = _parse_general_report_sections(report_text)

    # ── COVER TABLE (same style as individual reports) ────────────────────
    _add_general_cover_table(doc, specialty, period)

    # ── TABLA DE PARETO ──────────────────────────────────────────────────
    if sections.get("tabla_pareto"):
        _add_section_title(doc, "TABLA DE PARETO — " + specialty.upper(),
                           centered=True, underline=True)
        _add_pareto_table(doc, sections["tabla_pareto"])

    # ── POCAS CAUSAS VITALES ─────────────────────────────────────────────
    if sections.get("causas_vitales"):
        _add_section_title(doc, "POCAS CAUSAS VITALES (PUNTOS CRÍTICOS DE INTERVENCIÓN)",
                           centered=False, underline=False, size=12)
        _add_body_text(doc, sections["causas_vitales"])

    # ── RESUMEN EJECUTIVO ────────────────────────────────────────────────
    if sections.get("resumen_ejecutivo"):
        _add_section_title(doc, "RESUMEN EJECUTIVO", centered=True, underline=True)
        _add_body_text(doc, sections["resumen_ejecutivo"])

    # ── MÉDICOS EN RIESGO Y MEJOR EVALUADOS ──────────────────────────────
    if sections.get("medicos_riesgo"):
        _add_section_title(doc, "MÉDICOS EN RIESGO Y MEJOR EVALUADOS",
                           centered=False, underline=False, size=12)
        _add_body_text(doc, sections["medicos_riesgo"])

    # ── GRÁFICAS DE ANÁLISIS ─────────────────────────────────────────────
    if chart_pareto or chart_accumulated or chart_nc_vs_er:
        doc.add_page_break()
        _add_section_title(doc, "GRÁFICAS DE ANÁLISIS",
                           centered=True, underline=True)

    if chart_pareto:
        doc.add_paragraph()
        _add_subsection_title(doc, "Diagrama de Pareto — Hallazgos del Período")
        doc.add_picture(io.BytesIO(chart_pareto), width=Inches(6.0))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if chart_accumulated:
        doc.add_paragraph()
        _add_subsection_title(doc, "Hallazgos Acumulados por Período")
        doc.add_picture(io.BytesIO(chart_accumulated), width=Inches(6.0))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if chart_nc_vs_er:
        doc.add_paragraph()
        _add_subsection_title(doc, "Eventos de Riesgo vs No Conformidades")
        doc.add_picture(io.BytesIO(chart_nc_vs_er), width=Inches(5.0))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── FOOTER ───────────────────────────────────────────────────────────
    doc.add_paragraph()
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = note.add_run("Documento generado automáticamente por el Sistema de Auditoría Médica de Calidad")
    _set_run_format(run, size=9, italic=True, color=RGBColor(0x7F, 0x7F, 0x7F))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
