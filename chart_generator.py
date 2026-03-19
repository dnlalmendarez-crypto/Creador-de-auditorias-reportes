"""
Chart generator for Pareto analysis and general reports.
Generates matplotlib charts as PNG bytes for embedding in Word documents.
"""

import io
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import re


def _parse_pareto_table(report_text: str) -> list[dict]:
    """Extract rows from the Pareto table in the report text."""
    rows = []
    in_table = False
    for line in report_text.split("\n"):
        stripped = line.strip()
        if "Componente" in stripped and "Frecuencia" in stripped and "|" in stripped:
            in_table = True
            continue
        if in_table and stripped.startswith("|") and "---" not in stripped:
            cells = [c.strip() for c in stripped.split("|") if c.strip()]
            if len(cells) >= 6:
                try:
                    freq = int(re.sub(r"[^\d]", "", cells[3]))
                    pct_ind = float(re.sub(r"[^\d.]", "", cells[4]))
                    pct_acc = float(re.sub(r"[^\d.]", "", cells[5]))
                    rows.append({
                        "componente": cells[0],
                        "criterio": cells[1],
                        "causa": cells[2],
                        "frecuencia": freq,
                        "pct_individual": pct_ind,
                        "pct_acumulado": pct_acc,
                    })
                except (ValueError, IndexError):
                    continue
        elif in_table and not stripped.startswith("|") and stripped:
            in_table = False
    return rows


def _extract_nc_er_counts(report_text: str) -> tuple[int, int]:
    """Extract total No Conformidades and Eventos de Riesgo counts from report."""
    nc_total = 0
    er_total = 0
    text_lower = report_text.lower()

    nc_match = re.search(r"(\d+)\s*no\s*conformidad", text_lower)
    if nc_match:
        nc_total = int(nc_match.group(1))

    er_match = re.search(r"(\d+)\s*evento[s]?\s*de\s*riesgo", text_lower)
    if er_match:
        er_total = int(er_match.group(1))

    return nc_total, er_total


def generate_pareto_chart(report_text: str, specialty: str, period: str) -> bytes | None:
    """Generate a Pareto chart (bar + line) from the report text. Returns PNG bytes."""
    rows = _parse_pareto_table(report_text)
    if not rows:
        return None

    labels = [r["causa"][:35] + "..." if len(r["causa"]) > 35 else r["causa"] for r in rows]
    freqs = [r["frecuencia"] for r in rows]
    acumulados = [r["pct_acumulado"] for r in rows]

    fig, ax1 = plt.subplots(figsize=(12, 6))

    color_bars = []
    for acc in acumulados:
        if acc <= 80:
            color_bars.append("#FF4444")
        elif acc <= 95:
            color_bars.append("#FF9800")
        else:
            color_bars.append("#4CAF50")

    bars = ax1.bar(range(len(labels)), freqs, color=color_bars, edgecolor="white", linewidth=0.5)
    ax1.set_ylabel("Frecuencia", fontsize=10, fontweight="bold")
    ax1.set_xticks(range(len(labels)))
    ax1.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)

    ax2 = ax1.twinx()
    ax2.plot(range(len(labels)), acumulados, color="#1F3964", marker="o", linewidth=2, markersize=5)
    ax2.axhline(y=80, color="#FF0000", linestyle="--", linewidth=1, alpha=0.7, label="80% Pareto")
    ax2.set_ylabel("% Acumulado", fontsize=10, fontweight="bold")
    ax2.yaxis.set_major_formatter(mticker.PercentFormatter())
    ax2.set_ylim(0, 105)
    ax2.legend(loc="center right")

    plt.title(f"Diagrama de Pareto — {specialty} — {period}", fontsize=12, fontweight="bold", pad=15)
    fig.tight_layout()

    buf = io.BytesIO()
    try:
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        buf.seek(0)
        return buf.getvalue()
    finally:
        plt.close(fig)


def generate_accumulated_chart(
    individual_reports: dict, specialty: str
) -> bytes | None:
    """
    Generate a bar chart comparing findings across periods from individual reports.
    individual_reports: {code: {period, report_text, name, ...}}
    Returns PNG bytes.
    """
    if not individual_reports:
        return None

    period_findings = {}
    for code, rdata in individual_reports.items():
        period = rdata.get("period", "Desconocido")
        text = rdata.get("report_text", "")

        nc_count = 0
        er_count = 0
        nc_match = re.search(r"(\d+)\s*[Nn]o\s*[Cc]onformidad", text)
        er_match = re.search(r"(\d+)\s*[Ee]vento[s]?\s*de\s*[Rr]iesgo", text)
        if nc_match:
            nc_count = int(nc_match.group(1))
        if er_match:
            er_count = int(er_match.group(1))

        if period not in period_findings:
            period_findings[period] = {"nc": 0, "er": 0}
        period_findings[period]["nc"] += nc_count
        period_findings[period]["er"] += er_count

    if not period_findings:
        return None

    periods = list(period_findings.keys())
    nc_vals = [period_findings[p]["nc"] for p in periods]
    er_vals = [period_findings[p]["er"] for p in periods]

    fig, ax = plt.subplots(figsize=(10, 5))
    x = range(len(periods))
    width = 0.35

    bars1 = ax.bar([i - width/2 for i in x], nc_vals, width, label="No Conformidades", color="#FF4444", edgecolor="white")
    bars2 = ax.bar([i + width/2 for i in x], er_vals, width, label="Eventos de Riesgo", color="#FF9800", edgecolor="white")

    ax.set_ylabel("Cantidad de Hallazgos", fontsize=10, fontweight="bold")
    ax.set_xticks(list(x))
    ax.set_xticklabels(periods, rotation=30, ha="right", fontsize=9)
    ax.legend()
    ax.set_title(f"Hallazgos Acumulados por Período — {specialty}", fontsize=12, fontweight="bold")

    for bar in bars1:
        if bar.get_height() > 0:
            ax.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.3, str(int(bar.get_height())), ha="center", va="bottom", fontsize=8, fontweight="bold")
    for bar in bars2:
        if bar.get_height() > 0:
            ax.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.3, str(int(bar.get_height())), ha="center", va="bottom", fontsize=8, fontweight="bold")

    fig.tight_layout()
    buf = io.BytesIO()
    try:
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        buf.seek(0)
        return buf.getvalue()
    finally:
        plt.close(fig)


def generate_nc_vs_er_chart(report_text: str, specialty: str, period: str) -> bytes | None:
    """Generate a pie/donut chart comparing No Conformidades vs Eventos de Riesgo."""
    nc_total, er_total = _extract_nc_er_counts(report_text)

    if nc_total == 0 and er_total == 0:
        return None

    labels = ["No Conformidades", "Eventos de Riesgo"]
    values = [nc_total, er_total]
    colors = ["#FF4444", "#FF9800"]
    explode = (0.05, 0.05)

    fig, ax = plt.subplots(figsize=(7, 5))
    wedges, texts, autotexts = ax.pie(
        values, labels=labels, colors=colors, explode=explode,
        autopct=lambda pct: f"{pct:.1f}%\n({int(round(pct/100.*sum(values)))})",
        startangle=90, textprops={"fontsize": 10},
        wedgeprops={"edgecolor": "white", "linewidth": 2},
    )
    for autotext in autotexts:
        autotext.set_fontweight("bold")

    ax.set_title(f"Eventos de Riesgo vs No Conformidades — {specialty} — {period}",
                 fontsize=11, fontweight="bold", pad=15)

    fig.tight_layout()
    buf = io.BytesIO()
    try:
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        buf.seek(0)
        return buf.getvalue()
    finally:
        plt.close(fig)
