"""
generate_informe.py
────────────────────────────────────────────────────────────────────────────
Módulo para generar el INFORME DE AUDITORÍA en formato DOCX exactamente igual
al documento original. Usar desde Streamlit con st.download_button().

Especificaciones extraídas del DOCX original:
  - Página  : US Letter  12240 × 15840 DXA, márgenes 1080 DXA (≈0.75 in)
  - Fuente  : Arial, tamaño base 10 pt (20 half-points)
  - Colores :
      DARK_BLUE  = "1F3864"  → encabezados de tabla
      MED_BLUE   = "2E74B5"  → filas de subgrupo
      LIGHT_BLUE = "D6E4F0"  → filas alternadas claras
      PALE_BLUE  = "EBF3F9"  → filas de datos
      WHITE      = "FFFFFF"
  - Bordes  : single, color CCCCCC, tamaño 1 (todos los lados)
  - Tablas  :
      Tabla 1 – encabezado del doctor   col widths [2000, 7360]
      Tabla 2 – cumplimiento por criterio col widths [2000, 2200, 2000, 2000]
      Tabla 3 – análisis cuantitativo    col widths [1800, 4500, 1700, 1360]
"""

import io
import base64
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Twips, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement
from lxml import etree


# ─── CONSTANTES DE COLOR ────────────────────────────────────────────────────
DARK_BLUE  = "1F3864"   # encabezado principal de tablas
MED_BLUE   = "2E74B5"   # fila de subgrupo / totales
LIGHT_BLUE = "D6E4F0"   # fila alternada clara
PALE_BLUE  = "EBF3F9"   # fila de datos normal
WHITE      = "FFFFFF"


# ─── HELPERS DE BAJO NIVEL ──────────────────────────────────────────────────

def _set_cell_shading(cell, fill: str):
    """Aplica color de fondo a una celda."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),  "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill)
    existing = tcPr.find(qn("w:shd"))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(shd)


def _set_cell_borders(cell, color: str = "CCCCCC", size: str = "1"):
    """Aplica bordes a todos los lados de una celda."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"),   "single")
        border.set(qn("w:sz"),    size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        tcBorders.append(border)
    existing = tcPr.find(qn("w:tcBorders"))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(tcBorders)


def _set_cell_vertical_align(cell, align=WD_ALIGN_VERTICAL.CENTER):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    existing = tcPr.find(qn("w:vAlign"))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(vAlign)


def _set_col_width(table, col_widths_dxa: list):
    """Fija el ancho de cada columna (DXA)."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"),    str(sum(col_widths_dxa)))
    tblW.set(qn("w:type"), "dxa")
    existing = tblPr.find(qn("w:tblW"))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(tblW)
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        tbl.remove(tblGrid)
    tblGrid = OxmlElement("w:tblGrid")
    for w in col_widths_dxa:
        gridCol = OxmlElement("w:gridCol")
        gridCol.set(qn("w:w"), str(w))
        tblGrid.append(gridCol)
    tbl.insert(list(tbl).index(tblPr) + 1, tblGrid)


def _fix_cell_width(cell, width_dxa: int):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"),    str(width_dxa))
    tcW.set(qn("w:type"), "dxa")
    existing = tcPr.find(qn("w:tcW"))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(tcW)


def _style_cell(cell, fill: str, width_dxa: int | None = None,
                border_color: str = "CCCCCC"):
    """Aplica fondo + bordes + ancho de columna a una celda."""
    _set_cell_shading(cell, fill)
    _set_cell_borders(cell, color=border_color)
    _set_cell_vertical_align(cell)
    if width_dxa:
        _fix_cell_width(cell, width_dxa)


def _para_text(cell, text: str, bold: bool = False, italic: bool = False,
               font_size_pt: float = 9, color_hex: str | None = None,
               align=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    """
    Limpia párrafos existentes de la celda y escribe texto con formato.
    Soporta segmentos **negrita** inline si bold=True en el texto completo,
    o pasa una lista de tuplas [(texto, bold), ...] en `text`.
    """
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""
    para = cell.paragraphs[0]
    para.alignment = align
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.space_after  = Pt(1)
    segments = text if isinstance(text, list) else [(text, bold)]
    for seg_text, seg_bold in segments:
        run = para.add_run(seg_text)
        run.bold = seg_bold
        run.italic = italic
        run.font.name = "Arial"
        run.font.size = Pt(font_size_pt)
        if color_hex:
            r, g, b = (int(color_hex[i:i+2], 16) for i in (0, 2, 4))
            run.font.color.rgb = RGBColor(r, g, b)


# ─── CONFIGURACIÓN DEL DOCUMENTO ────────────────────────────────────────────

def _configure_document(doc: Document):
    """Establece tamaño de página, márgenes y fuente por defecto."""
    section = doc.sections[0]
    section.page_width   = Twips(12240)
    section.page_height  = Twips(15840)
    section.top_margin   = Twips(1080)
    section.bottom_margin = Twips(1080)
    section.left_margin  = Twips(1080)
    section.right_margin = Twips(1080)
    # Fuente por defecto: Arial 10pt
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10)


def _add_page_header(doc: Document):
    """Add DoctorSV logo (left) + UNIDAD DE GESTIÓN DE MEJORA CONTINUA (right)
    to the page header, matching the institutional template."""
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False

    # Create a two-column table for left/right alignment in header
    tbl = header.add_table(rows=1, cols=2, width=Twips(10080))
    tbl.autofit = True
    # Remove table borders
    tbl_xml = tbl._tbl
    tblPr = tbl_xml.tblPr if tbl_xml.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border_el = OxmlElement(f"w:{border_name}")
        border_el.set(qn("w:val"), "none")
        border_el.set(qn("w:sz"), "0")
        border_el.set(qn("w:space"), "0")
        border_el.set(qn("w:color"), "auto")
        borders.append(border_el)
    tblPr.append(borders)

    # Left cell: DoctorSV logo text
    left_cell = tbl.cell(0, 0)
    left_p = left_cell.paragraphs[0]
    left_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_doctor = left_p.add_run("Doctor")
    run_doctor.bold = True
    run_doctor.font.name = "Arial"
    run_doctor.font.size = Pt(14)
    run_doctor.font.color.rgb = RGBColor(0x1F, 0x39, 0x64)
    run_sv = left_p.add_run("SV")
    run_sv.bold = True
    run_sv.font.name = "Arial"
    run_sv.font.size = Pt(14)
    run_sv.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

    # Right cell: Institutional text
    right_cell = tbl.cell(0, 1)
    right_p = right_cell.paragraphs[0]
    right_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_ugmc = right_p.add_run("UNIDAD DE GESTIÓN DE MEJORA CONTINUA")
    run_ugmc.bold = True
    run_ugmc.font.name = "Arial"
    run_ugmc.font.size = Pt(7)
    run_ugmc.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)


def _add_watermark(doc: Document):
    """Add a faint diagonal text watermark 'DoctorSV' using Word's standard
    VML watermark mechanism embedded in the header."""
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False

    watermark_p = header.add_paragraph()
    watermark_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    r_elem = OxmlElement("w:r")
    pict = OxmlElement("w:pict")

    # VML shapetype for text watermark
    ns = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
    shapetype_xml = (
        f'<v:shapetype {ns} id="_x0000_t136" coordsize="21600,21600" '
        'o:spt="136" adj="10800" '
        'path="m@7,l@8,m@5,21600l@6,21600e">'
        '<v:formulas><v:f eqn="sum #0 0 10800"/>'
        '<v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>'
        '<v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/>'
        '<v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/>'
        '<v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>'
        '</v:formulas>'
        '<v:path textpathok="t" o:connecttype="custom" '
        'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" '
        'o:connectangles="270,180,90,0"/>'
        '<v:textpath on="t" fitshape="t"/>'
        '<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/>'
        '</v:handles><o:lock v:ext="edit" text="t" shapetype="t"/>'
        '</v:shapetype>'
    )
    pict.append(etree.fromstring(shapetype_xml))

    # VML shape for the actual watermark text
    shape_xml = (
        f'<v:shape {ns} id="PowerPlusWaterMarkObject" '
        'o:spid="_x0000_s2049" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:0;'
        'width:500pt;height:120pt;rotation:315;z-index:-251657216;'
        'mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
        'mso-position-vertical:center;mso-position-vertical-relative:margin" '
        'o:allowincell="f" fillcolor="#D0D0D0" stroked="f">'
        '<v:fill opacity=".15"/>'
        '<v:textpath style=\'font-family:"Arial";font-size:1pt\' '
        'string="DoctorSV"/>'
        '</v:shape>'
    )
    pict.append(etree.fromstring(shape_xml))

    r_elem.append(pict)
    watermark_p._p.append(r_elem)


# ─── SECCIÓN 1: TABLA DE ENCABEZADO ─────────────────────────────────────────

def add_header_table(doc: Document, data: dict):
    """
    Tabla de identificación del doctor.
    Col widths originales: [2000, 7360]
    """
    COL_W = [2000, 7360]
    ROWS = [
        ("NOMBRE",           data.get("nombre",       "")),
        ("CÓDIGO",           data.get("codigo",        "")),
        ("ESPECIALIDAD",     data.get("especialidad", "")),
        ("DEPENDENCIA",      data.get("dependencia",  "")),
        ("PERIODO AUDITADO", data.get("periodo",      "")),
    ]
    table = doc.add_table(rows=len(ROWS), cols=2)
    _set_col_width(table, COL_W)
    for i, (label, value) in enumerate(ROWS):
        row = table.rows[i]
        c0 = row.cells[0]
        _style_cell(c0, DARK_BLUE, COL_W[0])
        _para_text(c0, label, bold=True, font_size_pt=9, color_hex="FFFFFF",
                   align=WD_ALIGN_PARAGRAPH.LEFT)
        c1 = row.cells[1]
        _style_cell(c1, PALE_BLUE if i % 2 == 0 else WHITE, COL_W[1])
        _para_text(c1, value, bold=True, font_size_pt=9,
                   align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_paragraph()


# ─── SECCIÓN 2: RESUMEN EJECUTIVO (párrafos de texto) ───────────────────────

def add_section_title(doc: Document, title: str):
    """Agrega un título de sección en negrita (simula heading sin estilo)."""
    p = doc.add_paragraph()
    run = p.add_run(title)
    run.bold = True
    run.font.name  = "Arial"
    run.font.size  = Pt(11)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(3)


def add_bold_body_paragraph(doc: Document, segments: list):
    """
    Párrafo de cuerpo con segmentos de texto mezclados bold/normal.
    segments = [("texto normal ", False), ("texto bold", True), ...]
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    for text, bold in segments:
        run = p.add_run(text)
        run.bold = bold
        run.font.name = "Arial"
        run.font.size = Pt(10)


# ─── SECCIÓN 3: TABLA DE CUMPLIMIENTO ───────────────────────────────────────

def add_compliance_table(doc: Document, rows_data: list):
    """
    Tabla de cumplimiento por criterio.
    Col widths originales: [2000, 2200, 2000, 2000]  total=8200
    """
    COL_W = [2000, 2200, 2000, 2000]
    table = doc.add_table(rows=len(rows_data), cols=4)
    _set_col_width(table, COL_W)
    data_row_index = 0
    for i, row_info in enumerate(rows_data):
        row = table.rows[i]
        rtype = row_info.get("type", "data")
        cols  = row_info.get("cols", ["", "", "", ""])
        if rtype == "header":
            fill = DARK_BLUE; text_color = "FFFFFF"; bold = True; fsize = 9
        elif rtype == "subheader":
            fill = MED_BLUE;  text_color = "FFFFFF"; bold = True; fsize = 9
        elif rtype == "total":
            fill = MED_BLUE;  text_color = "FFFFFF"; bold = True; fsize = 9
        else:   # data
            fill = PALE_BLUE if data_row_index % 2 == 0 else WHITE
            data_row_index += 1
            text_color = "000000"; bold = False; fsize = 9
        for j, (cell, val) in enumerate(zip(row.cells, cols)):
            _style_cell(cell, fill, COL_W[j])
            align = (WD_ALIGN_PARAGRAPH.CENTER
                     if j >= 2 else WD_ALIGN_PARAGRAPH.LEFT)
            _para_text(cell, val, bold=bold, font_size_pt=fsize,
                       color_hex=text_color, align=align)
    doc.add_paragraph()


# ─── SECCIÓN 4: TABLA ANÁLISIS CUANTITATIVO ─────────────────────────────────

def add_quantitative_table(doc: Document, citas: list):
    """
    Tabla de análisis cuantitativo.
    Formato: 4 columnas → ID CITA | DIAGNÓSTICO | NOTA | SÍNTESIS DE HALLAZGOS.
    """
    COL_W = [1300, 2600, 900, 4560]
    HEADERS = ["ID Cita", "Diagnóstico (CIE-11)", "Nota", "Síntesis de Hallazgos"]
    all_rows = [{"type": "header", "cols": HEADERS}]
    for c in citas:
        all_rows.append({
            "type": "data",
            "cols": [
                str(c.get("num_cita", "")),
                c.get("diagnostico", ""),
                str(c.get("nota", "")),
                str(c.get("sintesis", "")),
            ]
        })
    num_cols = len(COL_W)
    table = doc.add_table(rows=len(all_rows), cols=num_cols)
    _set_col_width(table, COL_W)
    data_row_index = 0
    for i, row_info in enumerate(all_rows):
        row = table.rows[i]
        rtype = row_info["type"]
        cols  = row_info["cols"]
        if rtype == "header":
            fill = DARK_BLUE; text_color = "FFFFFF"; bold = True; fsize = 9
        else:
            fill = PALE_BLUE if data_row_index % 2 == 0 else WHITE
            data_row_index += 1
            text_color = "000000"; bold = False; fsize = 9
        for j, (cell, val) in enumerate(zip(row.cells, cols)):
            _style_cell(cell, fill, COL_W[j])
            # Center ID and Nota, left-align diagnóstico and síntesis
            align = (WD_ALIGN_PARAGRAPH.CENTER
                     if j in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT)
            _para_text(cell, val, bold=bold, font_size_pt=fsize,
                       color_hex=text_color, align=align)
    doc.add_paragraph()


def _parse_pipe_table(text: str) -> list[list[str]]:
    """Parse a markdown pipe-delimited table into rows of cells.
    Preserves empty cells to maintain column alignment."""
    import re as _re
    lines = [l.strip() for l in text.strip().split("\n") if l.strip()]
    table_lines = [
        l for l in lines
        if "|" in l and "---" not in l
        and not l.lower().startswith("leyenda")
        and "leyenda:" not in l.lower()
    ]
    rows = []
    for line in table_lines:
        # Split by | and strip leading/trailing empty segments from outer pipes
        parts = line.split("|")
        # Remove first and last empty strings from leading/trailing pipes
        if parts and not parts[0].strip():
            parts = parts[1:]
        if parts and not parts[-1].strip():
            parts = parts[:-1]
        # Strip markdown bold and whitespace, but KEEP empty cells
        cells = [_re.sub(r"\*\*([^*]+)\*\*", r"\1", c.strip()) for c in parts]
        if cells:
            rows.append(cells)
    return rows


def add_qualitative_table(doc: Document, table_text: str, title: str):
    """Add a qualitative analysis table (NC or ER) to the document.
    Groups rows by COMPONENTE and merges the first column vertically."""
    rows = _parse_pipe_table(table_text)
    if not rows:
        return

    # Add title
    p = doc.add_paragraph()
    run = p.add_run(title)
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(10)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)

    header_row = rows[0]
    data_rows = rows[1:]
    num_cols = len(header_row)

    # Column widths: COMPONENTE(1800) | CRITERIO(1600) | NC/ER(600) | TIPIFICACIÓN(2800) | IMPACTO(2560)
    COL_W = [1800, 1600, 600, 2800, 2560]
    while len(COL_W) < num_cols:
        COL_W.append(1800)
    COL_W = COL_W[:num_cols]

    # Fill in empty COMPONENTE cells with the last non-empty value
    # (Claude sometimes puts component only in first row of group)
    last_comp = ""
    for row in data_rows:
        if row[0].strip():
            last_comp = row[0].strip()
        else:
            row[0] = last_comp

    # Identify component groups for vertical merging
    groups = []  # [(start_index, count, component_name), ...]
    i = 0
    while i < len(data_rows):
        comp = data_rows[i][0].strip()
        count = 1
        while i + count < len(data_rows) and data_rows[i + count][0].strip() == comp:
            count += 1
        groups.append((i, count, comp))
        i += count

    total_rows = 1 + len(data_rows)  # header + data
    table = doc.add_table(rows=total_rows, cols=num_cols)
    _set_col_width(table, COL_W)

    # Header row
    for j in range(num_cols):
        cell = table.rows[0].cells[j]
        val = header_row[j] if j < len(header_row) else ""
        _style_cell(cell, DARK_BLUE, COL_W[j])
        _para_text(cell, val, bold=True, font_size_pt=8,
                   color_hex="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)

    # Data rows
    for row_idx, row_data in enumerate(data_rows):
        word_row = table.rows[1 + row_idx]
        alt_fill = PALE_BLUE if row_idx % 2 == 0 else WHITE
        for j in range(num_cols):
            cell = word_row.cells[j]
            val = row_data[j] if j < len(row_data) else ""
            if j == 0:
                # Component column — styled but text set later via merge
                _style_cell(cell, LIGHT_BLUE, COL_W[j])
            else:
                _style_cell(cell, alt_fill, COL_W[j])
                align = WD_ALIGN_PARAGRAPH.CENTER if j == 2 else WD_ALIGN_PARAGRAPH.LEFT
                _para_text(cell, val, bold=False, font_size_pt=8,
                           color_hex="000000", align=align)

    # Merge component cells vertically for each group
    for start_idx, count, comp_name in groups:
        if count > 1:
            top_cell = table.cell(1 + start_idx, 0)
            bottom_cell = table.cell(1 + start_idx + count - 1, 0)
            merged = top_cell.merge(bottom_cell)
            _style_cell(merged, LIGHT_BLUE, COL_W[0])
            _para_text(merged, comp_name, bold=True, font_size_pt=8,
                       color_hex="000000", align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            cell = table.cell(1 + start_idx, 0)
            _style_cell(cell, LIGHT_BLUE, COL_W[0])
            _para_text(cell, comp_name, bold=True, font_size_pt=8,
                       color_hex="000000", align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()


def _add_component_compliance_table(doc: Document, table_text: str):
    """Add the Cumplimiento por Componentes table to the DOCX."""
    rows = _parse_pipe_table(table_text)
    if not rows:
        return

    p = doc.add_paragraph()
    run = p.add_run("Cumplimiento por Componentes")
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(10)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)

    num_cols = len(rows[0])
    # Two columns: Componente | % Cumplimiento
    COL_W = [4600, 4760]
    while len(COL_W) < num_cols:
        COL_W.append(2000)
    COL_W = COL_W[:num_cols]

    table = doc.add_table(rows=len(rows), cols=num_cols)
    _set_col_width(table, COL_W)

    # Color scale for compliance percentages
    def _compliance_fill(val_str):
        import re as _re
        m = _re.search(r"[\d.]+", val_str)
        if not m:
            return WHITE, "000000"
        val = float(m.group())
        if val >= 98:
            return "C6EFCE", "006100"
        elif val >= 95:
            return "FFEB9C", "9C6500"
        elif val >= 85:
            return "FBE5D6", "BF4D00"
        else:
            return "FFC7CE", "9C0006"

    for i, row_cells in enumerate(rows):
        row = table.rows[i]
        if i == 0:
            # Header
            for j in range(num_cols):
                cell = row.cells[j]
                val = row_cells[j] if j < len(row_cells) else ""
                _style_cell(cell, DARK_BLUE, COL_W[j])
                _para_text(cell, val, bold=True, font_size_pt=9,
                           color_hex="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            for j in range(num_cols):
                cell = row.cells[j]
                val = row_cells[j] if j < len(row_cells) else ""
                if j == 0:
                    # Component name
                    is_summary = any(kw in val.lower() for kw in ["promedio", "puntaje"])
                    fill = LIGHT_BLUE if is_summary else WHITE
                    _style_cell(cell, fill, COL_W[j])
                    _para_text(cell, val, bold=is_summary, font_size_pt=9,
                               color_hex="000000", align=WD_ALIGN_PARAGRAPH.LEFT)
                else:
                    # Percentage — apply color
                    if "%" in val:
                        bg, fg = _compliance_fill(val)
                        _style_cell(cell, bg, COL_W[j])
                        _para_text(cell, val, bold=True, font_size_pt=9,
                                   color_hex=fg, align=WD_ALIGN_PARAGRAPH.CENTER)
                    else:
                        _style_cell(cell, WHITE, COL_W[j])
                        _para_text(cell, val, bold=True, font_size_pt=9,
                                   color_hex="000000", align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()


def _add_tendencias_section(doc: Document, tendencias_text: str):
    """Add the Tendencias table to DOCX.

    New format: 3 columns → TENDENCIA | [Periodo anterior] | [Periodo actual].
    Prefers a markdown pipe table if provided; falls back to the legacy
    "Positiva:/Sostenida:/Negativa:" list format.
    """
    import re as _re

    TREND_COLORS = {
        "POSITIVA": "38761D",
        "SOSTENIDA": "FF9900",
        "NEGATIVA": "980000",
    }

    # Attempt structured pipe-table parse first
    table_rows = _parse_pipe_table(tendencias_text)

    header = None
    data_rows = []
    if table_rows and len(table_rows) >= 2:
        header = table_rows[0]
        data_rows = table_rows[1:]
        # Only accept if it looks like our 3-column format
        if len(header) < 3:
            header = None
            data_rows = []

    if header:
        num_cols = min(len(header), 3)
        # Allocate widths: first column narrower, rest share remainder
        COL_W = [1800, 3780, 3780][:num_cols]
        while len(COL_W) < num_cols:
            COL_W.append(2000)

        table = doc.add_table(rows=1 + len(data_rows), cols=num_cols)
        _set_col_width(table, COL_W)

        # Header row
        for j in range(num_cols):
            cell = table.rows[0].cells[j]
            _style_cell(cell, DARK_BLUE, COL_W[j])
            label = header[j] if j < len(header) else ""
            _para_text(cell, label, bold=True, font_size_pt=9,
                       color_hex="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)

        # Data rows — color the trend label cell by its type
        for r, row_data in enumerate(data_rows):
            row = table.rows[1 + r]
            trend_label = row_data[0].strip() if row_data else ""
            trend_upper = trend_label.upper()
            trend_color = None
            for key, color in TREND_COLORS.items():
                if key in trend_upper:
                    trend_color = color
                    break

            for j in range(num_cols):
                cell = row.cells[j]
                val = row_data[j] if j < len(row_data) else ""
                if j == 0 and trend_color:
                    _style_cell(cell, trend_color, COL_W[j])
                    _para_text(cell, val, bold=True, font_size_pt=9,
                               color_hex="FFFFFF",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
                else:
                    fill = PALE_BLUE if r % 2 == 0 else WHITE
                    _style_cell(cell, fill, COL_W[j])
                    _para_text(cell, val, bold=False, font_size_pt=8,
                               color_hex="000000",
                               align=WD_ALIGN_PARAGRAPH.LEFT)

        doc.add_paragraph()
        return

    # ── Fallback: legacy list format (Positiva:/Sostenida:/Negativa:) ──
    positiva = []
    sostenida = []
    negativa = []
    current = None
    for line in tendencias_text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        upper = stripped.upper()
        if "POSITIVA" in upper:
            current = positiva
            after = _re.sub(r".*positiva\s*:?\s*", "", stripped, flags=_re.IGNORECASE).strip()
            if after and after != "-":
                positiva.append(after)
        elif "SOSTENIDA" in upper:
            current = sostenida
            after = _re.sub(r".*sostenida\s*:?\s*", "", stripped, flags=_re.IGNORECASE).strip()
            if after and after != "-":
                sostenida.append(after)
        elif "NEGATIVA" in upper:
            current = negativa
            after = _re.sub(r".*negativa\s*:?\s*", "", stripped, flags=_re.IGNORECASE).strip()
            if after and after != "-":
                negativa.append(after)
        elif current is not None:
            clean = _re.sub(r"^[-•]\s*", "", stripped)
            if clean:
                current.append(clean)

    max_rows = max(len(positiva), len(sostenida), len(negativa), 1)

    COL_W = [3120, 3120, 3120]
    table = doc.add_table(rows=1 + max_rows, cols=3)
    _set_col_width(table, COL_W)

    headers = [("POSITIVA", "38761D"), ("SOSTENIDA", "FF9900"), ("NEGATIVA", "980000")]
    for j, (label, color) in enumerate(headers):
        cell = table.rows[0].cells[j]
        _style_cell(cell, color, COL_W[j])
        _para_text(cell, label, bold=True, font_size_pt=9,
                   color_hex="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)

    for r in range(max_rows):
        row = table.rows[1 + r]
        for j, items in enumerate([positiva, sostenida, negativa]):
            cell = row.cells[j]
            val = items[r] if r < len(items) else ""
            _style_cell(cell, WHITE, COL_W[j])
            _para_text(cell, val, bold=False, font_size_pt=8,
                       color_hex="000000", align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.add_paragraph()


def _add_conclusiones_table(doc: Document, conclusiones_text: str):
    """Add the Conclusiones y Acciones Requeridas table to DOCX."""
    rows = _parse_pipe_table(conclusiones_text)
    if not rows:
        # Render as plain text if not a table
        p = doc.add_paragraph(conclusiones_text)
        if p.runs:
            p.runs[0].font.name = "Arial"
            p.runs[0].font.size = Pt(10)
        return

    num_cols = len(rows[0])
    COL_W = [1800, 7560]
    while len(COL_W) < num_cols:
        COL_W.append(3000)
    COL_W = COL_W[:num_cols]

    priority_colors = {
        "critica": "FFC7CE", "crítica": "FFC7CE",
        "alta": "FBE5D6",
        "media": "FFEB9C",
    }
    priority_text_colors = {
        "critica": "9C0006", "crítica": "9C0006",
        "alta": "BF4D00",
        "media": "9C6500",
    }

    table = doc.add_table(rows=len(rows), cols=num_cols)
    _set_col_width(table, COL_W)

    for i, row_cells in enumerate(rows):
        row = table.rows[i]
        if i == 0:
            for j in range(num_cols):
                cell = row.cells[j]
                val = row_cells[j] if j < len(row_cells) else ""
                _style_cell(cell, DARK_BLUE, COL_W[j])
                _para_text(cell, val, bold=True, font_size_pt=9,
                           color_hex="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            for j in range(num_cols):
                cell = row.cells[j]
                val = row_cells[j] if j < len(row_cells) else ""
                if j == 0:
                    # Priority cell with color
                    prio_key = val.lower().strip()
                    fill = priority_colors.get(prio_key, WHITE)
                    text_color = priority_text_colors.get(prio_key, "000000")
                    _style_cell(cell, fill, COL_W[j])
                    _para_text(cell, val, bold=True, font_size_pt=9,
                               color_hex=text_color, align=WD_ALIGN_PARAGRAPH.CENTER)
                else:
                    _style_cell(cell, WHITE, COL_W[j])
                    _para_text(cell, val, bold=False, font_size_pt=9,
                               color_hex="000000", align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.add_paragraph()


# ─── FUNCIÓN PRINCIPAL ───────────────────────────────────────────────────────

def generate_informe(data: dict) -> bytes:
    """
    Genera el DOCX completo y devuelve bytes para st.download_button().

    data = {
        "nombre", "codigo", "especialidad", "dependencia", "periodo",
        "periodo_codigo",
        "resumen_ejecutivo",
        "hallazgos": { "COMPONENTE": {"nc": N, "er": N}, ... },
        "compliance_rows": [ {"type": ..., "cols": [...]}, ... ],
        "comentario_seguimiento",
        "citas": [ {"num_cita", "diagnostico", "nc", "er"}, ... ],
        "analisis_cualitativo",
    }
    """
    doc = Document()
    _configure_document(doc)

    # ── Encabezado de página + marca de agua ──
    _add_page_header(doc)
    _add_watermark(doc)

    # ── Título del informe ───────────────────────────────────────────────
    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    import re as _re
    _year_match = _re.search(r"20\d{2}", data.get("periodo", ""))
    _year = _year_match.group(0) if _year_match else str(__import__("datetime").date.today().year)
    _spec = data.get("spec_abbrev", "MG")
    run = titulo.add_run(
        f"INFORME DE AUDITORÍA N° {data.get('codigo', '')}"
        f"-{_spec}-{_year}-{data.get('periodo_codigo', 'P003')}"
    )
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(12)
    titulo.paragraph_format.space_after = Pt(8)

    # ── Tabla encabezado ─────────────────────────────────────────────────
    add_header_table(doc, data)

    # ── Resumen Ejecutivo ────────────────────────────────────────────────
    add_section_title(doc, "1. RESUMEN EJECUTIVO")
    doc.add_paragraph(data.get("resumen_ejecutivo", ""))

    # Cumplimiento por componentes table (from Reporte Global)
    cumpl_comp = data.get("cumplimiento_componentes", "")
    if cumpl_comp:
        _add_component_compliance_table(doc, cumpl_comp)

    # Tendencias table
    tendencias = data.get("tendencias", "")
    if tendencias:
        _add_tendencias_section(doc, tendencias)

    doc.add_paragraph()

    # ── Reporte de cumplimiento ──────────────────────────────────────────
    add_section_title(doc, "2. CUADRO DE CUMPLIMIENTO POR CRITERIO")
    add_compliance_table(doc, data.get("compliance_rows", []))

    # Leyenda
    leyenda = doc.add_paragraph()
    leyenda.paragraph_format.space_after = Pt(6)
    for text, bold in [
        ("■ ", True), ("≥98% Óptimo  ", False),
        ("■ ", True), ("≥95% a <98% Muy Bueno  ", False),
        ("■ ", True), ("≥85% a <95% Aceptable  ", False),
        ("■ ", True), ("<85% Oportunidad de mejora", False),
    ]:
        r = leyenda.add_run(text)
        r.bold = bold
        r.font.name = "Arial"
        r.font.size = Pt(8)

    # ── Comentario de seguimiento ────────────────────────────────────────
    add_section_title(doc, "COMENTARIO DE SEGUIMIENTO Y COMPARACIÓN DE PERIODOS")
    p = doc.add_paragraph(data.get("comentario_seguimiento", ""))
    if p.runs:
        p.runs[0].font.name = "Arial"
        p.runs[0].font.size = Pt(10)
    doc.add_paragraph()

    # ── Análisis de No Conformidades ─────────────────────────────────────
    add_section_title(doc, "3. ANÁLISIS DE NO CONFORMIDADES")
    add_section_title(doc, "3.1 ANÁLISIS CUANTITATIVO")
    add_quantitative_table(doc, data.get("citas", []))

    add_section_title(doc, "3.2 ANÁLISIS CUALITATIVO POR COMPONENTE")

    # Render structured qualitative tables if available
    nc_table_text = data.get("nc_table", "")
    er_table_text = data.get("er_table", "")

    if nc_table_text:
        add_qualitative_table(doc, nc_table_text, "Tabla de No Conformidades")
    if er_table_text:
        add_qualitative_table(doc, er_table_text, "Tabla de Eventos de Riesgo")

    # Fallback: render as plain text if no structured tables
    if not nc_table_text and not er_table_text:
        cualitativo_text = data.get("analisis_cualitativo", "")
        if cualitativo_text:
            p = doc.add_paragraph(cualitativo_text)
            if p.runs:
                p.runs[0].font.name = "Arial"
                p.runs[0].font.size = Pt(10)

    # Síntesis paragraph
    sintesis_text = data.get("sintesis", "")
    if sintesis_text:
        p = doc.add_paragraph()
        run = p.add_run("SÍNTESIS: ")
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run = p.add_run(sintesis_text)
        run.font.name = "Arial"
        run.font.size = Pt(10)

    # ── Conclusiones y Acciones Requeridas ──────────────────────────────
    conclusiones = data.get("conclusiones", "")
    if conclusiones:
        add_section_title(doc, "4. CONCLUSIONES Y ACCIONES REQUERIDAS")
        _add_conclusiones_table(doc, conclusiones)

    # ── Firma ────────────────────────────────────────────────────────────
    firma = data.get("firma", "")
    if firma:
        doc.add_paragraph()
        for line in firma.strip().split("\n"):
            line = line.strip()
            if not line:
                continue
            p = doc.add_paragraph()
            # Bold the label part before ":"
            if ":" in line:
                label, value = line.split(":", 1)
                run = p.add_run(f"{label}: ")
                run.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(10)
                run = p.add_run(value.strip())
                run.font.name = "Arial"
                run.font.size = Pt(10)
            else:
                run = p.add_run(line)
                run.font.name = "Arial"
                run.font.size = Pt(10)
    else:
        # Default firma
        import datetime as _dt
        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run("Auditor Responsable: ")
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run = p.add_run("UGMC — Unidad de Gestión de Mejora Continua")
        run.font.name = "Arial"
        run.font.size = Pt(10)
        p = doc.add_paragraph()
        run = p.add_run("Fecha de Emisión: ")
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run = p.add_run(_dt.date.today().strftime("%d/%m/%Y"))
        run.font.name = "Arial"
        run.font.size = Pt(10)

    # ── Serializar a bytes ───────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()
