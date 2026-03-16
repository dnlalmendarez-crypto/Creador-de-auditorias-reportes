"""
Creador de Auditorías y Reportes Médicos
Main Streamlit application.
"""

import streamlit as st
import os
import io
import zipfile
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

import data_loader as dl
import claude_analyzer as ca
import report_generator as rg

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Auditorías Médicas de Calidad",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CUSTOM CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1F3964 0%, #2E74B5 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.8rem; }
    .main-header p  { color: #BDD7EE; margin: 0.3rem 0 0; }
    .section-card {
        background: #f8f9fa;
        border-left: 4px solid #2E74B5;
        padding: 1rem 1.2rem;
        border-radius: 4px;
        margin-bottom: 1rem;
    }
    .status-red    { color: #FF0000; font-weight: bold; }
    .status-orange { color: #FF6600; font-weight: bold; }
    .status-yellow { color: #FFC000; font-weight: bold; }
    .status-green  { color: #00B050; font-weight: bold; }
    .stButton>button {
        background-color: #1F3964;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.5rem 1.5rem;
    }
    .stButton>button:hover { background-color: #2E74B5; }
</style>
""", unsafe_allow_html=True)

# ─── HEADER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🏥 Creador de Auditorías y Reportes Médicos</h1>
    <p>Sistema de Análisis de No Conformidades y Generación de Informes de Calidad</p>
</div>
""", unsafe_allow_html=True)

# ─── SESSION STATE ────────────────────────────────────────────────────────────
if "generated_reports" not in st.session_state:
    st.session_state.generated_reports = {}
if "file_data" not in st.session_state:
    st.session_state.file_data = {}

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")

    api_key = st.text_input(
        "API Key de Anthropic",
        value=os.environ.get("ANTHROPIC_API_KEY", ""),
        type="password",
        help="Ingresa tu API Key de Anthropic Claude",
    )

    model_options = {
        "Claude Opus 4.6 (Más potente)": "claude-opus-4-6",
        "Claude Sonnet 4.6 (Balanceado)": "claude-sonnet-4-6",
    }
    selected_model_label = st.selectbox("Modelo Claude", list(model_options.keys()))
    selected_model = model_options[selected_model_label]

    st.markdown("---")
    st.markdown("### 📁 Archivos de Referencia")
    st.caption("Carga los archivos base para el análisis")

    clasificacion_file = st.file_uploader(
        "Clasificación de No Conformidades (.xlsx)",
        type=["xlsx", "xls"],
        key="clasificacion",
        help="Archivo: Clasificación de no conformidades.XLSX",
    )

    tipificaciones_file = st.file_uploader(
        "Tipificaciones de Auditoría (.xlsx)",
        type=["xlsx", "xls"],
        key="tipificaciones",
        help="Archivo: TIPIFICACIONES DE USO COMUN EN AUDITORIA MEDICA DE CALIDAD.XLSX",
    )

    st.markdown("---")
    st.markdown("### 📊 Archivos de Datos")

    graficas_file = st.file_uploader(
        "Gráficas de Cumplimiento (.xlsx)",
        type=["xlsx", "xls"],
        key="graficas",
        help="Archivo: 🕸️ Graficas de Cumplimiento 2026 || MEDGEN || DoctorSV",
    )

    base_datos_file = st.file_uploader(
        "Base de Datos x Cita (.xlsx)",
        type=["xlsx", "xls"],
        key="base_datos",
        help="Archivo: 🛢️ Base de Datos x Cita 2026 || DoctorSv",
    )

    st.markdown("---")
    st.markdown("### 🎨 Leyenda de Colores")
    st.markdown('<span class="status-green">● 98-100%: Excelente</span>', unsafe_allow_html=True)
    st.markdown('<span class="status-yellow">● 95-97%: Muy Bueno</span>', unsafe_allow_html=True)
    st.markdown('<span class="status-orange">● 85-94%: Aceptable</span>', unsafe_allow_html=True)
    st.markdown('<span class="status-red">● <85%: Oportunidad de Mejora</span>', unsafe_allow_html=True)


# ─── LOAD & CACHE DATA ────────────────────────────────────────────────────────
def load_uploaded_excel(uploaded_file, loader_fn):
    """Load an uploaded Streamlit file with a loader function."""
    if uploaded_file is None:
        return None
    import tempfile, os
    suffix = Path(uploaded_file.name).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name
    result = loader_fn(tmp_path)
    os.unlink(tmp_path)
    uploaded_file.seek(0)
    return result


@st.cache_data(show_spinner=False)
def get_file_data(_clasificacion_key, _tipificaciones_key, _graficas_key, _base_datos_key,
                  clasificacion_bytes, tipificaciones_bytes, graficas_bytes, base_datos_bytes):
    """Cache loaded Excel data keyed by file content hashes."""
    import tempfile, os

    def bytes_to_df(b, loader_fn):
        if b is None:
            return None
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(b)
            tmp_path = tmp.name
        result = loader_fn(tmp_path)
        os.unlink(tmp_path)
        return result

    return {
        "clasificacion": bytes_to_df(clasificacion_bytes, dl.load_clasificacion),
        "tipificaciones": bytes_to_df(tipificaciones_bytes, dl.load_tipificaciones),
        "graficas": bytes_to_df(graficas_bytes, dl.load_graficas_cumplimiento),
        "base_datos": bytes_to_df(base_datos_bytes, dl.load_base_datos_cita),
    }


def get_bytes(f):
    if f is None:
        return None
    b = f.read()
    f.seek(0)
    return b


# ─── MAIN CONTENT ─────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📋 Generar Reportes", "📄 Ver Reportes Generados", "ℹ️ Ayuda"])

with tab1:
    st.markdown("### 👨‍⚕️ Configuración de Médicos y Período")

    col1, col2 = st.columns([2, 1])
    with col1:
        period_input = st.text_input(
            "Período de Auditoría",
            placeholder="Ej: 01 al 15 de febrero 2026",
            help="Ingresa el período exactamente como aparece en los archivos",
        )
    with col2:
        specialty_input = st.text_input(
            "Especialidad",
            value="Medicina General",
            help="Especialidad médica para buscar en la hoja correspondiente",
        )

    st.markdown("#### Médicos a Auditar")
    st.caption("Agrega uno o más médicos. El código debe coincidir exactamente con el archivo de cumplimiento.")

    if "doctors" not in st.session_state:
        st.session_state.doctors = [{"name": "", "code": ""}]

    def add_doctor():
        st.session_state.doctors.append({"name": "", "code": ""})

    def remove_doctor(idx):
        if len(st.session_state.doctors) > 1:
            st.session_state.doctors.pop(idx)

    for i, doctor in enumerate(st.session_state.doctors):
        c1, c2, c3 = st.columns([3, 2, 1])
        with c1:
            st.session_state.doctors[i]["name"] = st.text_input(
                f"Nombre del Médico #{i+1}",
                value=doctor["name"],
                key=f"doc_name_{i}",
                placeholder="Ej: Dr. Juan Pérez",
            )
        with c2:
            st.session_state.doctors[i]["code"] = st.text_input(
                f"Código (COD) #{i+1}",
                value=doctor["code"],
                key=f"doc_code_{i}",
                placeholder="Ej: MED001",
            )
        with c3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️", key=f"remove_{i}", help="Eliminar médico"):
                remove_doctor(i)
                st.rerun()

    st.button("➕ Agregar Médico", on_click=add_doctor)

    st.markdown("---")

    # Validate inputs
    valid_doctors = [d for d in st.session_state.doctors if d["name"].strip() and d["code"].strip()]
    files_loaded = any([graficas_file, base_datos_file])

    if not api_key:
        st.warning("⚠️ Ingresa tu API Key de Anthropic en el panel lateral para continuar.")
    elif not period_input:
        st.info("ℹ️ Ingresa el período de auditoría para continuar.")
    elif not valid_doctors:
        st.info("ℹ️ Agrega al menos un médico con nombre y código.")
    elif not files_loaded:
        st.warning("⚠️ Carga al menos los archivos de Gráficas de Cumplimiento y Base de Datos.")
    else:
        st.success(f"✅ Listo para generar {len(valid_doctors)} reporte(s) para el período: **{period_input}**")

        if st.button("🚀 Generar Reportes", type="primary", use_container_width=True):
            # Load data
            with st.spinner("Cargando archivos..."):
                file_cache = get_file_data(
                    id(clasificacion_file), id(tipificaciones_file),
                    id(graficas_file), id(base_datos_file),
                    get_bytes(clasificacion_file), get_bytes(tipificaciones_file),
                    get_bytes(graficas_file), get_bytes(base_datos_file),
                )

            graficas_data = file_cache.get("graficas")
            base_datos_data = file_cache.get("base_datos")
            clasificacion_df = file_cache.get("clasificacion")
            tipificaciones_df = file_cache.get("tipificaciones")

            # Build reference text
            clasificacion_text = (
                dl.dataframe_to_markdown_table(clasificacion_df)
                if clasificacion_df is not None
                else "No se proporcionó archivo de clasificación."
            )
            tipificaciones_text = (
                dl.dataframe_to_markdown_table(tipificaciones_df)
                if tipificaciones_df is not None
                else "No se proporcionó archivo de tipificaciones."
            )

            # Process each doctor
            for doctor in valid_doctors:
                doc_name = doctor["name"].strip()
                doc_code = doctor["code"].strip()

                st.markdown(f"---\n#### Procesando: {doc_name} ({doc_code})")
                progress = st.progress(0)
                status_placeholder = st.empty()
                report_placeholder = st.empty()

                # Step 1: Find consultations
                status_placeholder.info("🔍 Buscando datos de consultas...")
                progress.progress(10)

                consultations_df = None
                if base_datos_data:
                    consultations_df = dl.find_doctor_consultations(
                        base_datos_data, doc_name, doc_code, period_input
                    )

                consultations_text = (
                    dl.dataframe_to_markdown_table(consultations_df)
                    if consultations_df is not None and not consultations_df.empty
                    else "No se encontraron consultas para este médico en el período indicado."
                )

                # Step 2: Find compliance data
                status_placeholder.info("📊 Buscando datos de cumplimiento...")
                progress.progress(25)

                compliance_text = "No se encontraron datos de cumplimiento para este médico."
                if graficas_data:
                    compliance_results = dl.find_doctor_compliance(
                        graficas_data, doc_name, doc_code, period_input
                    )
                    if compliance_results:
                        import pandas as pd
                        parts = []
                        for sheet, result in compliance_results.items():
                            rows = result.get("rows")
                            if rows is not None and not rows.empty:
                                parts.append(f"**Hoja: {sheet}**\n{dl.dataframe_to_markdown_table(rows)}")
                        if parts:
                            compliance_text = "\n\n".join(parts)

                # Step 3: Generate AI analysis
                status_placeholder.info("🤖 Generando análisis con Claude AI...")
                progress.progress(40)

                report_text_parts = []

                def stream_cb(chunk):
                    report_text_parts.append(chunk)
                    current = "".join(report_text_parts)
                    # Show live preview using markdown (avoids key conflicts)
                    report_placeholder.markdown(
                        f"**Vista previa del análisis:**\n\n```\n{current[-2000:]}\n```"
                    )

                try:
                    full_report = ca.analyze_with_claude(
                        doctor_name=doc_name,
                        doctor_code=doc_code,
                        period=period_input,
                        consultations_table=consultations_text,
                        compliance_table=compliance_text,
                        clasificacion_table=clasificacion_text,
                        tipificaciones_list=tipificaciones_text,
                        api_key=api_key,
                        model=selected_model,
                        stream_callback=stream_cb,
                    )
                    progress.progress(80)
                    status_placeholder.info("📝 Generando documento Word...")

                    # Step 4: Generate Word document
                    docx_bytes = rg.generate_full_report_from_text(
                        doctor_name=doc_name,
                        doctor_code=doc_code,
                        period=period_input,
                        full_report_text=full_report,
                        specialty=specialty_input,
                    )

                    # Store in session
                    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in doc_name)
                    filename = f"Reporte_Auditoria_{safe_name}_{doc_code}_{period_input[:10].replace(' ', '_')}.docx"
                    st.session_state.generated_reports[doc_code] = {
                        "name": doc_name,
                        "code": doc_code,
                        "period": period_input,
                        "filename": filename,
                        "docx_bytes": docx_bytes,
                        "report_text": full_report,
                    }

                    progress.progress(100)
                    status_placeholder.success(f"✅ Reporte generado exitosamente para {doc_name}")

                    st.download_button(
                        label=f"⬇️ Descargar Reporte - {doc_name}",
                        data=docx_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{doc_code}",
                    )

                except Exception as e:
                    progress.progress(0)
                    status_placeholder.error(f"❌ Error generando reporte para {doc_name}: {str(e)}")
                    st.exception(e)

            # Offer ZIP download if multiple doctors
            if len(valid_doctors) > 1 and len(st.session_state.generated_reports) > 0:
                st.markdown("---")
                st.markdown("### 📦 Descargar Todos los Reportes")

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for code, rdata in st.session_state.generated_reports.items():
                        zf.writestr(rdata["filename"], rdata["docx_bytes"])
                zip_buffer.seek(0)

                st.download_button(
                    label="⬇️ Descargar ZIP con Todos los Reportes",
                    data=zip_buffer.getvalue(),
                    file_name=f"Reportes_Auditoria_{period_input[:10].replace(' ', '_')}.zip",
                    mime="application/zip",
                    use_container_width=True,
                )


with tab2:
    st.markdown("### 📄 Reportes Generados en esta Sesión")

    if not st.session_state.generated_reports:
        st.info("No hay reportes generados aún. Ve a la pestaña 'Generar Reportes' para crear uno.")
    else:
        for code, rdata in st.session_state.generated_reports.items():
            with st.expander(f"📋 {rdata['name']} ({code}) - {rdata['period']}", expanded=False):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.text_area(
                        "Contenido del Reporte",
                        value=rdata["report_text"],
                        height=400,
                        disabled=True,
                        key=f"view_{code}",
                    )
                with col2:
                    st.download_button(
                        label="⬇️ Descargar .docx",
                        data=rdata["docx_bytes"],
                        file_name=rdata["filename"],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl2_{code}",
                    )
                    st.markdown(f"**Médico:** {rdata['name']}")
                    st.markdown(f"**Código:** {code}")
                    st.markdown(f"**Período:** {rdata['period']}")

        if st.button("🗑️ Limpiar Reportes de Sesión", type="secondary"):
            st.session_state.generated_reports = {}
            st.rerun()


with tab3:
    st.markdown("""
    ### ℹ️ Guía de Uso

    #### Archivos Requeridos

    | Archivo | Descripción |
    |---------|-------------|
    | **Clasificación de No Conformidades** | Define si cada criterio es No Conformidad o Evento de Riesgo |
    | **Tipificaciones de Auditoría** | Lista oficial de tipificaciones permitidas |
    | **Gráficas de Cumplimiento** | Porcentajes de cumplimiento por criterio por médico y período |
    | **Base de Datos x Cita** | Registro de consultas auditadas con IDs, diagnósticos y hallazgos |

    #### Flujo de Trabajo

    1. **Carga los archivos** en el panel lateral izquierdo
    2. **Configura el período** y la especialidad
    3. **Agrega los médicos** con su nombre y código (COD) exacto
    4. **Haz clic en "Generar Reportes"**
    5. **Descarga** los archivos `.docx` generados individualmente o en ZIP

    #### Estructura del Reporte Generado

    El reporte incluye:
    - **Análisis Cuantitativo**: Tabla de IDs con diagnósticos y conteo de hallazgos
    - **Análisis Cualitativo**: No Conformidades y Eventos de Riesgo por componente y criterio
    - **Resumen Ejecutivo**: Para Alta Gerencia (máximo media página)
    - **Cuadro de Cumplimiento**: Porcentajes exactos del archivo de gráficas
    - **Comentario de Seguimiento**: Comparación de períodos (si aplica)

    #### Categorías de Color

    | Color | Rango | Categoría |
    |-------|-------|-----------|
    | 🟢 Verde | 98% - 100% | Excelente |
    | 🟡 Amarillo | 95% - 97% | Muy Bueno |
    | 🟠 Anaranjado | 85% - 94% | Aceptable |
    | 🔴 Rojo | < 85% | Oportunidad de Mejora |

    #### Notas Importantes

    - Los IDs de consulta son números de **5 a 10 dígitos**
    - El sistema busca al médico por **Código (COD)** primero, luego por nombre
    - Los porcentajes en el cuadro de cumplimiento se toman **exactamente** del archivo cargado
    - Cada reporte es **independiente** por médico (no se mezclan datos)
    - Los componentes siempre van en orden: **Anamnesis → Examen Físico → Diagnóstico → Productos de la Consulta**
    """)
