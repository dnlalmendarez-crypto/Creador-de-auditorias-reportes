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


# ─── WATERMARK SVG (El Salvador coat of arms + stars) ────────────────────────

def _build_watermark_svg() -> str:
    """Build an SVG watermark with scattered stars and a simplified
    El Salvador coat of arms silhouette, matching the institutional template."""
    return '''<svg xmlns="http://www.w3.org/2000/svg" width="816" height="1056" viewBox="0 0 816 1056">
  <defs>
    <style>
      .wm-star { fill: #D5D5D5; opacity: 0.32; }
      .esc { fill: #D0D0D0; opacity: 0.20; }
      .esc-line { stroke: #CFCFCF; fill: none; opacity: 0.18; }
    </style>
  </defs>
  <!-- Scattered stars -->
  <polygon class="wm-star" points="580,80 592,116 630,116 600,138 612,174 580,152 548,174 560,138 530,116 568,116"/>
  <polygon class="wm-star" points="350,155 360,183 390,183 365,200 375,228 350,211 325,228 335,200 310,183 340,183"/>
  <polygon class="wm-star" points="300,380 310,408 340,408 315,425 325,453 300,436 275,453 285,425 260,408 290,408"/>
  <polygon class="wm-star" points="560,330 570,358 600,358 575,375 585,403 560,386 535,403 545,375 520,358 550,358"/>
  <polygon class="wm-star" points="340,590 350,618 380,618 355,635 365,663 340,646 315,663 325,635 300,618 330,618"/>
  <polygon class="wm-star" points="290,800 300,828 330,828 305,845 315,873 290,856 265,873 275,845 250,828 280,828"/>
  <polygon class="wm-star" points="550,720 560,748 590,748 565,765 575,793 550,776 525,793 535,765 510,748 540,748"/>
  <polygon class="wm-star" points="400,930 410,958 440,958 415,975 425,1003 400,986 375,1003 385,975 360,958 390,958"/>

  <!-- El Salvador Coat of Arms — simplified silhouette (bottom-right) -->
  <g transform="translate(480, 440) scale(2.0)">
    <!-- Outer circle (laurel wreath) -->
    <circle cx="85" cy="85" r="82" class="esc-line" stroke-width="3"/>
    <circle cx="85" cy="85" r="75" class="esc-line" stroke-width="1.5"/>

    <!-- Triangle -->
    <polygon points="85,22 155,130 15,130" class="esc" style="opacity:0.18"/>

    <!-- Five volcanoes -->
    <polygon points="42,130 55,100 68,130" class="esc"/>
    <polygon points="55,130 70,95 85,130" class="esc"/>
    <polygon points="72,130 85,90 98,130" class="esc"/>
    <polygon points="88,130 101,95 114,130" class="esc"/>
    <polygon points="102,130 115,100 128,130" class="esc"/>

    <!-- Sun above volcanoes -->
    <circle cx="85" cy="55" r="10" class="esc"/>
    <!-- Sun rays -->
    <line x1="85" y1="38" x2="85" y2="30" class="esc-line" stroke-width="2"/>
    <line x1="71" y1="42" x2="65" y2="35" class="esc-line" stroke-width="2"/>
    <line x1="99" y1="42" x2="105" y2="35" class="esc-line" stroke-width="2"/>
    <line x1="68" y1="55" x2="60" y2="55" class="esc-line" stroke-width="2"/>
    <line x1="102" y1="55" x2="110" y2="55" class="esc-line" stroke-width="2"/>

    <!-- Phrygian cap -->
    <path d="M 78,68 Q 85,58 92,68" class="esc" style="opacity:0.15"/>

    <!-- Laurel leaves left -->
    <path d="M 10,50 Q 18,42 22,58 Q 14,53 10,50" class="esc"/>
    <path d="M 5,68 Q 13,60 17,76 Q 9,71 5,68" class="esc"/>
    <path d="M 3,88 Q 11,80 16,96 Q 8,91 3,88" class="esc"/>
    <path d="M 5,108 Q 13,100 20,113 Q 12,110 5,108" class="esc"/>
    <path d="M 12,125 Q 20,118 27,130 Q 19,128 12,125" class="esc"/>

    <!-- Laurel leaves right -->
    <path d="M 160,50 Q 152,42 148,58 Q 156,53 160,50" class="esc"/>
    <path d="M 165,68 Q 157,60 153,76 Q 161,71 165,68" class="esc"/>
    <path d="M 167,88 Q 159,80 154,96 Q 162,91 167,88" class="esc"/>
    <path d="M 165,108 Q 157,100 150,113 Q 158,110 165,108" class="esc"/>
    <path d="M 158,125 Q 150,118 143,130 Q 151,128 158,125" class="esc"/>

    <!-- Ribbon at bottom -->
    <path d="M 25,145 Q 85,162 145,145" class="esc-line" stroke-width="3"/>
    <path d="M 30,150 Q 85,166 140,150" class="esc-line" stroke-width="2"/>

    <!-- "REPUBLICA DE EL SALVADOR" text arc (decorative) -->
    <text x="85" y="178" text-anchor="middle" font-size="6" fill="#D0D0D0" opacity="0.18"
          font-family="serif" letter-spacing="1">REPÚBLICA DE EL SALVADOR</text>

    <!-- Five flags behind triangle -->
    <line x1="45" y1="20" x2="45" y2="5" class="esc-line" stroke-width="1.5"/>
    <line x1="65" y1="15" x2="65" y2="2" class="esc-line" stroke-width="1.5"/>
    <line x1="85" y1="12" x2="85" y2="0" class="esc-line" stroke-width="1.5"/>
    <line x1="105" y1="15" x2="105" y2="2" class="esc-line" stroke-width="1.5"/>
    <line x1="125" y1="20" x2="125" y2="5" class="esc-line" stroke-width="1.5"/>
    <!-- Small flag rectangles -->
    <rect x="42" y="2" width="8" height="5" rx="1" class="esc" style="opacity:0.15"/>
    <rect x="62" y="0" width="8" height="5" rx="1" class="esc" style="opacity:0.15"/>
    <rect x="82" y="-2" width="8" height="5" rx="1" class="esc" style="opacity:0.15"/>
    <rect x="102" y="0" width="8" height="5" rx="1" class="esc" style="opacity:0.15"/>
    <rect x="122" y="2" width="8" height="5" rx="1" class="esc" style="opacity:0.15"/>
  </g>
</svg>'''


def _watermark_data_uri() -> str:
    """Return data URI for the watermark SVG."""
    svg = _build_watermark_svg()
    b64 = base64.b64encode(svg.encode()).decode()
    return f"data:image/svg+xml;base64,{b64}"


# ─── CSS ─────────────────────────────────────────────────────────────────────
BASE_CSS = f"""
@page {{
    size: Letter;
    margin: 2.5cm 2.5cm 2cm 2.5cm;
    @top-left {{
        content: element(page-header-left);
    }}
    @top-right {{
        content: element(page-header-right);
    }}
}}
body {{
    font-family: {FONT};
    font-size: 11pt;
    color: {COLORS["body_text"]};
    line-height: 1.4;
}}
/* ── Page header (running elements) ── */
.page-header-left {{
    position: running(page-header-left);
    font-family: {FONT};
}}
.page-header-right {{
    position: running(page-header-right);
    font-family: {FONT};
    font-size: 7.5pt;
    color: {COLORS["section_bg"]};
    font-weight: bold;
    letter-spacing: 0.5px;
    text-align: right;
    white-space: nowrap;
}}
.logo-doctor {{
    font-size: 16pt;
    font-weight: bold;
    color: {COLORS["header_bg"]};
    letter-spacing: -0.5px;
}}
.logo-sv {{
    font-size: 16pt;
    font-weight: bold;
    color: {COLORS["section_bg"]};
    letter-spacing: -0.5px;
}}
/* ── Watermark background ── */
.watermark {{
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    z-index: -1;
    opacity: 1;
    pointer-events: none;
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


def _header_and_watermark_html() -> str:
    """Return HTML for running page header and fixed watermark background."""
    wm_uri = _watermark_data_uri()
    return f'''<!-- Running header elements (repeat on every page via CSS @page) -->
<div class="page-header-left">
  <span class="logo-doctor">Doctor</span><span class="logo-sv">SV</span>
</div>
<div class="page-header-right">
  UNIDAD DE GESTIÓN DE MEJORA CONTINUA
</div>
<!-- Watermark background (fixed, behind all content) -->
<img class="watermark" src="{wm_uri}" alt=""/>
'''


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
    # Filter: keep only real table rows (contain |), exclude separator rows (---)
    # and the "Leyenda:" line which contains | but is not a table row
    table_lines = [
        l for l in lines
        if "|" in l
        and "---" not in l
        and not l.lower().startswith("leyenda")
        and "leyenda:" not in l.lower()
    ]
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

    # ── Header + Watermark ──
    parts.append(_header_and_watermark_html())

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

    # ── Header + Watermark ──
    parts.append(_header_and_watermark_html())

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
