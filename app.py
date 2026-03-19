"""
Creador de Auditorías y Reportes Médicos
Main Streamlit application.
Supports both local Excel upload and Google Sheets connection.
"""

import streamlit as st
import os
import io
import json
import zipfile
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

import data_loader as dl
import claude_analyzer as ca
import report_generator as rg
import report_pdf_generator as rpdf
from generate_informe import generate_informe
from informe_adapter import build_informe_data
from period_utils import generate_periods

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
if "gsheets_connected" not in st.session_state:
    st.session_state.gsheets_connected = False
if "gsheets_data" not in st.session_state:
    st.session_state.gsheets_data = {}
if "template_bytes" not in st.session_state:
    st.session_state.template_bytes = None

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")

    # Read API key from: st.secrets (Streamlit Cloud) > .env > manual input
    default_key = ""
    try:
        default_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        pass
    if not default_key:
        default_key = os.environ.get("ANTHROPIC_API_KEY", "")

    api_key = st.text_input(
        "API Key de Anthropic",
        value=default_key,
        type="password",
        help="Ingresa tu API Key de Anthropic Claude. En Streamlit Cloud se lee automáticamente de Secrets.",
    )

    model_options = {
        "Claude Opus 4.6 (Más potente)": "claude-opus-4-6",
        "Claude Sonnet 4.6 (Balanceado)": "claude-sonnet-4-6",
    }
    selected_model_label = st.selectbox("Modelo Claude", list(model_options.keys()))
    selected_model = model_options[selected_model_label]

    st.markdown("---")

    # ─── DATA SOURCE SELECTOR ────────────────────────────────────────────────
    st.markdown("### 📂 Origen de Datos")
    data_source = st.radio(
        "Selecciona el origen de los archivos:",
        ["📤 Subir archivos Excel", "📊 Google Sheets"],
        key="data_source",
        help="Elige si deseas cargar archivos locales o conectar directamente a Google Sheets.",
    )

    clasificacion_file = None
    tipificaciones_file = None
    graficas_file = None
    base_datos_file = None
    gs_clasificacion_id = ""
    gs_tipificaciones_id = ""
    gs_graficas_id = ""
    gs_base_datos_id = ""

    if data_source == "📤 Subir archivos Excel":
        # ─── EXCEL UPLOAD MODE ───────────────────────────────────────────────
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
            help="Archivo: Gráficas de Cumplimiento 2026",
        )

        base_datos_file = st.file_uploader(
            "Base de Datos x Cita (.xlsx)",
            type=["xlsx", "xls"],
            key="base_datos",
            help="Archivo: Base de Datos x Cita 2026",
        )

    else:
        # ─── GOOGLE SHEETS MODE ─────────────────────────────────────────────
        st.markdown("### 🔑 Credenciales de Google")
        st.caption("Sube el archivo JSON de Service Account o configúralo en Secrets.")

        # Try to load from st.secrets first
        gcp_creds = None
        try:
            if "gcp_service_account" in st.secrets:
                gcp_creds = dict(st.secrets["gcp_service_account"])
        except Exception:
            pass

        if not gcp_creds:
            creds_file = st.file_uploader(
                "Service Account JSON",
                type=["json"],
                key="gcp_creds_file",
                help="Archivo JSON de credenciales de Google Cloud Service Account.",
            )
            if creds_file:
                try:
                    gcp_creds = json.loads(creds_file.read())
                    creds_file.seek(0)
                except Exception as e:
                    st.error(f"Error leyendo credenciales: {e}")
        else:
            st.success("Credenciales cargadas desde Secrets")

        if gcp_creds:
            st.session_state["gcp_creds"] = gcp_creds

        st.markdown("---")
        st.markdown("### 📋 URLs o IDs de Google Sheets")
        st.caption("Pega el enlace completo o solo el ID de cada spreadsheet.")

        gs_clasificacion_id = st.text_input(
            "Clasificación de No Conformidades",
            key="gs_clasificacion",
            placeholder="https://docs.google.com/spreadsheets/d/... o ID",
        )
        gs_tipificaciones_id = st.text_input(
            "Tipificaciones de Auditoría",
            key="gs_tipificaciones",
            placeholder="https://docs.google.com/spreadsheets/d/... o ID",
        )
        gs_graficas_id = st.text_input(
            "Gráficas de Cumplimiento",
            key="gs_graficas",
            placeholder="https://docs.google.com/spreadsheets/d/... o ID",
        )
        gs_base_datos_id = st.text_input(
            "Base de Datos x Cita",
            key="gs_base_datos",
            placeholder="https://docs.google.com/spreadsheets/d/... o ID",
        )

        if st.button("🔗 Conectar a Google Sheets"):
            if not st.session_state.get("gcp_creds"):
                st.error("Primero carga las credenciales de Service Account.")
            else:
                try:
                    import google_sheets_loader as gsl
                    client = gsl.connect_with_service_account(st.session_state["gcp_creds"])
                    loaded = {}

                    with st.spinner("Conectando a Google Sheets..."):
                        if gs_clasificacion_id.strip():
                            loaded["clasificacion"] = gsl.load_clasificacion_from_sheets(
                                client, gs_clasificacion_id.strip()
                            )
                        if gs_tipificaciones_id.strip():
                            loaded["tipificaciones"] = gsl.load_tipificaciones_from_sheets(
                                client, gs_tipificaciones_id.strip()
                            )
                        if gs_graficas_id.strip():
                            loaded["graficas"] = gsl.load_graficas_from_sheets(
                                client, gs_graficas_id.strip()
                            )
                        if gs_base_datos_id.strip():
                            loaded["base_datos"] = gsl.load_base_datos_from_sheets(
                                client, gs_base_datos_id.strip()
                            )

                    st.session_state.gsheets_data = loaded
                    st.session_state.gsheets_connected = True
                    st.success(f"Conectado. Se cargaron {len(loaded)} archivo(s) desde Google Sheets.")
                except Exception as e:
                    st.error(f"Error conectando a Google Sheets: {e}")

    st.markdown("---")
    st.markdown("### 📄 Plantilla Base (Opcional)")
    st.caption("Sube un archivo .docx que servirá como base para el reporte. Se conservarán estilos, encabezados y pies de página.")
    template_file = st.file_uploader(
        "Archivo de plantilla (.docx)",
        type=["docx"],
        key="template_upload",
        help="Sube un archivo Word (.docx) que se usará como plantilla base para generar los reportes.",
    )
    if template_file:
        st.session_state.template_bytes = template_file.read()
        template_file.seek(0)
        # Extract and apply colors/fonts from the template
        rg.apply_template_styles(st.session_state.template_bytes)
        st.success("Plantilla cargada correctamente. Colores y fuentes extraídos del template.")
    else:
        st.session_state.template_bytes = None
        rg.apply_template_styles(None)

    st.markdown("---")
    st.markdown("### ☁️ Subir a Google Drive (Opcional)")
    st.caption("Sube los reportes generados directamente a una carpeta de Drive.")
    drive_folder_input = st.text_input(
        "Carpeta de destino en Drive",
        key="drive_folder",
        placeholder="https://drive.google.com/drive/folders/... o ID",
        help="Pega la URL o ID de la carpeta de Google Drive donde se subirán los reportes.",
    )

    st.markdown("---")
    st.markdown("### 🎨 Leyenda de Colores")
    st.markdown('<span class="status-green">● 98-100%: Excelente</span>', unsafe_allow_html=True)
    st.markdown('<span class="status-yellow">● 95-97%: Muy Bueno</span>', unsafe_allow_html=True)
    st.markdown('<span class="status-orange">● 85-94%: Aceptable</span>', unsafe_allow_html=True)
    st.markdown('<span class="status-red">● <85%: Oportunidad de Mejora</span>', unsafe_allow_html=True)


# ─── LOAD & CACHE DATA (Excel mode) ─────────────────────────────────────────
@st.cache_data(show_spinner=False)
def get_file_data(_clasificacion_key, _tipificaciones_key, _graficas_key, _base_datos_key,
                  clasificacion_bytes, tipificaciones_bytes, graficas_bytes, base_datos_bytes):
    """Cache loaded Excel data keyed by file content hashes."""
    import tempfile

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


def _file_hash(f):
    """Return MD5 hash of file content for stable cache keys."""
    if f is None:
        return None
    import hashlib
    b = f.read()
    f.seek(0)
    return hashlib.md5(b).hexdigest()


# ─── SPECIALTIES LIST ────────────────────────────────────────────────────────
SPECIALTIES = [
    "Medicina General (MEDGEN)",
    "Medicina General Servicio Social (MEDGEN SS)",
    "Medicina Interna (MEDINT)",
    "Pediatría (PEDIA)",
    "Psicología (PSICO)",
    "Nutrición (NUTRI)",
    "Ginecología (GYOBS)",
]

# ─── PERIOD GENERATION ───────────────────────────────────────────────────────
current_year = datetime.now().year
available_years = list(range(current_year - 1, current_year + 2))


# ─── MAIN CONTENT ─────────────────────────────────────────────────────────────
tab1, tab2, tab4, tab3 = st.tabs(["📋 Generar Reportes", "📄 Ver Reportes Generados", "📊 Reporte General", "ℹ️ Ayuda"])

with tab1:
    st.markdown("### 📅 Período de Auditoría")

    col_year, col_period = st.columns([1, 3])
    with col_year:
        selected_year = st.selectbox("Año", available_years, index=available_years.index(current_year))

    periods = generate_periods(selected_year)
    period_labels = [p["label"] for p in periods]

    # Default to a recent period based on current date
    default_idx = 0
    now = datetime.now()
    if selected_year == now.year:
        month_idx = (now.month - 1) * 2
        if now.day > 15:
            month_idx += 1
        default_idx = min(month_idx, len(periods) - 1)

    with col_period:
        selected_period_label = st.selectbox(
            "Período de Auditoría",
            period_labels,
            index=default_idx,
            help="Selecciona el período quincenal a auditar",
        )

    # Get selected period info
    selected_period = next(p for p in periods if p["label"] == selected_period_label)
    period_input = selected_period["short_label"]

    st.info(f"📅 Período seleccionado: **{period_input}**")

    st.markdown("---")

    st.markdown("### 👨‍⚕️ Configuración de Médicos y Especialidad")

    specialty_input = st.selectbox(
        "Especialidad",
        options=SPECIALTIES,
        index=0,
        help="Especialidad médica para buscar en la hoja correspondiente",
        key="specialty_tab1",
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

    # Determine if data is available
    using_gsheets = data_source == "📊 Google Sheets"
    if using_gsheets:
        files_loaded = st.session_state.gsheets_connected and bool(st.session_state.gsheets_data)
    else:
        files_loaded = any([graficas_file, base_datos_file])

    if not api_key:
        st.warning("⚠️ Ingresa tu API Key de Anthropic en el panel lateral para continuar.")
    elif not valid_doctors:
        st.info("ℹ️ Agrega al menos un médico con nombre y código.")
    elif not files_loaded:
        if using_gsheets:
            st.warning("⚠️ Conecta a Google Sheets primero usando el botón en el panel lateral.")
        else:
            st.warning("⚠️ Carga al menos los archivos de Gráficas de Cumplimiento y Base de Datos.")
    else:
        st.success(f"✅ Listo para generar {len(valid_doctors)} reporte(s) para el período: **{period_input}**")

        if st.button("🚀 Generar Reportes", type="primary", use_container_width=True):
            # Load data based on source
            if using_gsheets:
                gs_data = st.session_state.gsheets_data
                graficas_data = gs_data.get("graficas")
                base_datos_data = gs_data.get("base_datos")
                clasificacion_df = gs_data.get("clasificacion")
                tipificaciones_df = gs_data.get("tipificaciones")
            else:
                with st.spinner("Cargando archivos..."):
                    file_cache = get_file_data(
                        _file_hash(clasificacion_file), _file_hash(tipificaciones_file),
                        _file_hash(graficas_file), _file_hash(base_datos_file),
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
                    status_placeholder.info("📝 Generando documentos (DOCX + PDF)...")

                    # Step 4: Parse sections and build structured data
                    parsed_sections = ca.parse_report_sections(full_report)
                    informe_data = build_informe_data(
                        doctor_name=doc_name,
                        doctor_code=doc_code,
                        period=period_input,
                        specialty=specialty_input,
                        sections=parsed_sections,
                    )

                    # Step 4a: Generate DOCX with generate_informe (exact format)
                    docx_bytes = generate_informe(informe_data)

                    # Step 4b: Generate PDF
                    try:
                        pdf_bytes = rpdf.generate_full_report_pdf_from_text(
                            doctor_name=doc_name,
                            doctor_code=doc_code,
                            period=period_input,
                            full_report_text=full_report,
                            specialty=specialty_input,
                        )
                    except Exception as pdf_err:
                        pdf_bytes = None
                        st.warning(f"PDF no disponible: {pdf_err}")

                    # Store in session
                    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in doc_name)
                    filename_docx = f"Reporte_Auditoria_{safe_name}_{doc_code}_{period_input[:10].replace(' ', '_')}.docx"
                    filename_pdf = filename_docx.replace(".docx", ".pdf")
                    st.session_state.generated_reports[doc_code] = {
                        "name": doc_name,
                        "code": doc_code,
                        "period": period_input,
                        "filename": filename_docx,
                        "filename_pdf": filename_pdf,
                        "docx_bytes": docx_bytes,
                        "pdf_bytes": pdf_bytes,
                        "report_text": full_report,
                    }

                    progress.progress(100)
                    status_placeholder.success(f"✅ Reporte generado exitosamente para {doc_name}")

                    col_pdf, col_docx = st.columns(2)
                    if pdf_bytes:
                        col_pdf.download_button(
                            label=f"⬇️ PDF - {doc_name}",
                            data=pdf_bytes,
                            file_name=filename_pdf,
                            mime="application/pdf",
                            key=f"dl_pdf_{doc_code}",
                        )
                    col_docx.download_button(
                        label=f"⬇️ DOCX - {doc_name}",
                        data=docx_bytes,
                        file_name=filename_docx,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{doc_code}",
                    )

                    # Upload to Google Drive if configured
                    if drive_folder_input.strip() and st.session_state.get("gcp_creds"):
                        if st.button(f"☁️ Subir a Drive - {doc_name}", key=f"drive_{doc_code}"):
                            try:
                                import google_sheets_loader as gsl
                                folder_id = gsl.extract_folder_id(drive_folder_input.strip())
                                with st.spinner("Subiendo a Google Drive..."):
                                    link = gsl.upload_file_to_drive(
                                        credentials_info=st.session_state["gcp_creds"],
                                        file_bytes=docx_bytes,
                                        filename=filename_docx,
                                        folder_id=folder_id,
                                    )
                                st.success(f"Subido a Drive: [Abrir archivo]({link})")
                            except Exception as drive_err:
                                st.error(f"Error subiendo a Drive: {drive_err}")

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
                        if rdata.get("pdf_bytes"):
                            zf.writestr(rdata.get("filename_pdf", rdata["filename"].replace(".docx", ".pdf")),
                                        rdata["pdf_bytes"])
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
                    if rdata.get("pdf_bytes"):
                        st.download_button(
                            label="⬇️ Descargar PDF",
                            data=rdata["pdf_bytes"],
                            file_name=rdata.get("filename_pdf", rdata["filename"].replace(".docx", ".pdf")),
                            mime="application/pdf",
                            key=f"dl2_pdf_{code}",
                        )
                    st.download_button(
                        label="⬇️ Descargar DOCX",
                        data=rdata["docx_bytes"],
                        file_name=rdata["filename"],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl2_{code}",
                    )
                    st.markdown(f"**Médico:** {rdata['name']}")
                    st.markdown(f"**Código:** {code}")
                    st.markdown(f"**Período:** {rdata['period']}")

                    if drive_folder_input.strip() and st.session_state.get("gcp_creds"):
                        if st.button(f"☁️ Subir a Drive", key=f"drive2_{code}"):
                            try:
                                import google_sheets_loader as gsl
                                folder_id = gsl.extract_folder_id(drive_folder_input.strip())
                                with st.spinner("Subiendo a Google Drive..."):
                                    link = gsl.upload_file_to_drive(
                                        credentials_info=st.session_state["gcp_creds"],
                                        file_bytes=rdata["docx_bytes"],
                                        filename=rdata["filename"],
                                        folder_id=folder_id,
                                    )
                                st.success(f"[Abrir en Drive]({link})")
                            except Exception as drive_err:
                                st.error(f"Error: {drive_err}")

        if st.button("🗑️ Limpiar Reportes de Sesión", type="secondary"):
            st.session_state.generated_reports = {}
            st.rerun()


with tab4:
    st.markdown("### 📊 Reporte General por Especialidad — Análisis de Pareto")
    st.caption("Genera un análisis de Pareto (80/20) para identificar las causas vitales de No Conformidades y Eventos de Riesgo por especialidad.")

    import chart_generator as cg

    # Specialty selection
    gen_specialty = st.selectbox(
        "Especialidad a evaluar",
        options=SPECIALTIES,
        index=0,
        key="gen_report_specialty",
        help="Selecciona la especialidad para el reporte general de Pareto.",
    )

    # Period selection (reuse same period logic)
    col_yr_gen, col_per_gen = st.columns([1, 3])
    with col_yr_gen:
        gen_year = st.selectbox("Año", available_years, index=available_years.index(current_year), key="gen_year")
    gen_periods = generate_periods(gen_year)
    gen_period_labels = [p["label"] for p in gen_periods]
    gen_default_idx = 0
    if gen_year == now.year:
        gen_month_idx = (now.month - 1) * 2
        if now.day > 15:
            gen_month_idx += 1
        gen_default_idx = min(gen_month_idx, len(gen_periods) - 1)
    with col_per_gen:
        gen_period_label = st.selectbox("Período", gen_period_labels, index=gen_default_idx, key="gen_period")
    gen_selected_period = next(p for p in gen_periods if p["label"] == gen_period_label)
    gen_period = gen_selected_period["short_label"]

    st.info(f"Especialidad: **{gen_specialty}** — Período: **{gen_period}**")

    # Data availability check
    using_gsheets_gen = data_source == "📊 Google Sheets"
    if using_gsheets_gen:
        gen_data_ready = st.session_state.gsheets_connected and bool(st.session_state.gsheets_data)
    else:
        gen_data_ready = any([base_datos_file, graficas_file])

    if not api_key:
        st.warning("⚠️ Ingresa tu API Key de Anthropic en el panel lateral.")
    elif not gen_data_ready:
        if using_gsheets_gen:
            st.warning("⚠️ Conecta a Google Sheets primero.")
        else:
            st.warning("⚠️ Carga al menos el archivo de Base de Datos x Cita.")
    elif st.button("📊 Generar Reporte de Pareto", type="primary", use_container_width=True, key="gen_general"):
        # Load data
        if using_gsheets_gen:
            gs_data = st.session_state.gsheets_data
            base_datos_data_gen = gs_data.get("base_datos")
            tipificaciones_df_gen = gs_data.get("tipificaciones")
        else:
            with st.spinner("Cargando archivos..."):
                file_cache_gen = get_file_data(
                    _file_hash(clasificacion_file), _file_hash(tipificaciones_file),
                    _file_hash(graficas_file), _file_hash(base_datos_file),
                    get_bytes(clasificacion_file), get_bytes(tipificaciones_file),
                    get_bytes(graficas_file), get_bytes(base_datos_file),
                )
            base_datos_data_gen = file_cache_gen.get("base_datos")
            tipificaciones_df_gen = file_cache_gen.get("tipificaciones")

        # Build text from base_datos (all sheets)
        base_datos_text_gen = "No se proporcionó Base de Datos."
        if base_datos_data_gen:
            parts = []
            for sheet_name, df in base_datos_data_gen.items():
                if df is not None and not df.empty:
                    parts.append(f"**Hoja: {sheet_name}**\n{dl.dataframe_to_markdown_table(df)}")
            if parts:
                base_datos_text_gen = "\n\n".join(parts)

        tipificaciones_text_gen = (
            dl.dataframe_to_markdown_table(tipificaciones_df_gen)
            if tipificaciones_df_gen is not None
            else "No se proporcionó archivo de tipificaciones."
        )

        progress_gen = st.progress(0)
        status_gen = st.empty()
        preview_gen = st.empty()

        status_gen.info("🤖 Generando análisis de Pareto...")
        progress_gen.progress(20)

        gen_parts = []

        def gen_stream_cb(chunk):
            gen_parts.append(chunk)
            current = "".join(gen_parts)
            preview_gen.markdown(
                f"**Vista previa:**\n\n```\n{current[-2000:]}\n```"
            )

        try:
            general_text = ca.analyze_general_report(
                base_datos_text=base_datos_text_gen,
                tipificaciones_text=tipificaciones_text_gen,
                period=gen_period,
                specialty=gen_specialty,
                api_key=api_key,
                model=selected_model,
                stream_callback=gen_stream_cb,
            )
            progress_gen.progress(60)
            status_gen.info("📊 Generando gráficas...")

            # Generate charts
            chart_pareto = cg.generate_pareto_chart(general_text, gen_specialty, gen_period)
            chart_accumulated = cg.generate_accumulated_chart(
                st.session_state.generated_reports, gen_specialty
            )
            chart_nc_er = cg.generate_nc_vs_er_chart(general_text, gen_specialty, gen_period)

            progress_gen.progress(80)
            status_gen.info("📝 Generando documentos (PDF + Word)...")

            # Generate Word doc
            gen_docx_bytes = rg.generate_general_report_docx(
                report_text=general_text,
                specialty=gen_specialty,
                period=gen_period,
                chart_pareto=chart_pareto,
                chart_accumulated=chart_accumulated,
                chart_nc_vs_er=chart_nc_er,
                template_bytes=st.session_state.template_bytes,
            )

            # Generate PDF
            try:
                gen_pdf_bytes = rpdf.generate_general_report_pdf(
                    report_text=general_text,
                    specialty=gen_specialty,
                    period=gen_period,
                    chart_pareto=chart_pareto,
                    chart_accumulated=chart_accumulated,
                    chart_nc_vs_er=chart_nc_er,
                )
            except Exception as pdf_err:
                gen_pdf_bytes = None
                st.warning(f"PDF no disponible: {pdf_err}")

            safe_spec = "".join(c if c.isalnum() or c in " _-" else "_" for c in gen_specialty)
            gen_filename = f"Reporte_Pareto_{safe_spec}_{gen_period[:10].replace(' ', '_')}.docx"
            gen_filename_pdf = gen_filename.replace(".docx", ".pdf")

            progress_gen.progress(100)
            status_gen.success("Reporte de Pareto generado exitosamente.")

            # Display charts in Streamlit
            if chart_pareto:
                st.image(chart_pareto, caption="Diagrama de Pareto — Hallazgos del Período", use_container_width=True)
            if chart_nc_er:
                st.image(chart_nc_er, caption="Eventos de Riesgo vs No Conformidades", use_container_width=True)
            if chart_accumulated:
                st.image(chart_accumulated, caption="Hallazgos Acumulados por Período", use_container_width=True)

            st.markdown("---")
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                if gen_pdf_bytes:
                    st.download_button(
                        label="⬇️ Descargar Reporte de Pareto (.pdf)",
                        data=gen_pdf_bytes,
                        file_name=gen_filename_pdf,
                        mime="application/pdf",
                        key="dl_general_pdf",
                    )
                st.download_button(
                    label="⬇️ Descargar Reporte de Pareto (.docx)",
                    data=gen_docx_bytes,
                    file_name=gen_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_general",
                )
            with col_dl2:
                if drive_folder_input.strip() and st.session_state.get("gcp_creds"):
                    if st.button("☁️ Subir a Drive", key="drive_general"):
                        try:
                            import google_sheets_loader as gsl
                            folder_id = gsl.extract_folder_id(drive_folder_input.strip())
                            with st.spinner("Subiendo a Google Drive..."):
                                link = gsl.upload_file_to_drive(
                                    credentials_info=st.session_state["gcp_creds"],
                                    file_bytes=gen_docx_bytes,
                                    filename=gen_filename,
                                    folder_id=folder_id,
                                )
                            st.success(f"[Abrir en Drive]({link})")
                        except Exception as drive_err:
                            st.error(f"Error subiendo a Drive: {drive_err}")

            with st.expander("Ver texto completo del reporte", expanded=False):
                st.text_area("Contenido", value=general_text, height=500, disabled=True, key="gen_text_view")

        except Exception as e:
            progress_gen.progress(0)
            status_gen.error(f"❌ Error generando reporte: {str(e)}")
            st.exception(e)


with tab3:
    st.markdown("""
    ### ℹ️ Guía de Uso

    #### Origen de Datos

    La aplicación soporta **dos modos** de carga de datos:

    | Modo | Descripción |
    |------|-------------|
    | **📤 Subir archivos Excel** | Carga archivos .xlsx/.xls desde tu computadora |
    | **📊 Google Sheets** | Conecta directamente a Google Sheets con Service Account |

    #### Configuración de Google Sheets

    Para usar Google Sheets necesitas:
    1. Crear un **Service Account** en Google Cloud Console
    2. Habilitar la **Google Sheets API** y **Google Drive API**
    3. Descargar el archivo JSON de credenciales
    4. **Compartir** cada spreadsheet con el email del Service Account (permisos de lectura)
    5. Pegar el enlace o ID de cada spreadsheet en el panel lateral

    En Streamlit Cloud, puedes configurar las credenciales en **Settings > Secrets**:
    ```toml
    [gcp_service_account]
    type = "service_account"
    project_id = "tu-proyecto"
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
    client_email = "...@...iam.gserviceaccount.com"
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    ```

    #### Archivos Requeridos

    | Archivo | Descripción |
    |---------|-------------|
    | **Clasificación de No Conformidades** | Define si cada criterio es No Conformidad o Evento de Riesgo |
    | **Tipificaciones de Auditoría** | Lista oficial de tipificaciones permitidas |
    | **Gráficas de Cumplimiento** | Porcentajes de cumplimiento por criterio por médico y período |
    | **Base de Datos x Cita** | Registro de consultas auditadas con IDs, diagnósticos y hallazgos |

    #### Períodos de Auditoría

    Los períodos se generan automáticamente de forma quincenal:
    - **Periodo 1:** 01 al 15 de enero
    - **Periodo 2:** 16 al 31 de enero
    - **Periodo 3:** 01 al 15 de febrero
    - ...y así sucesivamente hasta diciembre

    #### Plantilla Base (Opcional)

    Puedes subir un archivo `.docx` como plantilla base:
    - Se conservan **estilos, encabezados, pies de página y márgenes** del template
    - El contenido del reporte se agrega al final del documento plantilla
    - Útil para mantener el formato corporativo o institucional

    #### Subir a Google Drive (Opcional)

    Puedes subir los reportes generados directamente a una carpeta de Google Drive:
    1. Asegúrate de tener las **credenciales de Service Account** cargadas
    2. El Service Account necesita permisos de **Google Drive API** con acceso de escritura
    3. **Comparte la carpeta de destino** con el email del Service Account (permisos de Editor)
    4. Pega la URL o ID de la carpeta en el panel lateral

    #### Flujo de Trabajo

    1. **Selecciona el origen** de datos (Excel o Google Sheets)
    2. **Carga los archivos** o conecta a Google Sheets
    3. **(Opcional)** Sube una plantilla `.docx` base
    4. **(Opcional)** Configura la carpeta de Google Drive
    5. **Selecciona el año y período** quincenal
    6. **Agrega los médicos** con su nombre y código (COD) exacto
    7. **Haz clic en "Generar Reportes"**
    8. **Descarga** los archivos `.docx` o súbelos directamente a Drive

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
