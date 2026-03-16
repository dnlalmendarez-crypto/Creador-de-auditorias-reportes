# 🏥 Creador de Auditorías y Reportes Médicos

Sistema de análisis de no conformidades y generación automatizada de informes de auditoría médica de calidad, impulsado por Claude AI (Anthropic).

## Características

- Análisis de No Conformidades y Eventos de Riesgo por médico
- Generación de reportes individuales en formato Word (.docx)
- Análisis cualitativo estructurado por componente (Anamnesis, Examen Físico, Diagnóstico, Productos de la Consulta)
- Resumen Ejecutivo para Alta Gerencia
- Cuadro de Cumplimiento con código de colores
- Comparación de períodos auditados
- Descarga individual o en ZIP para múltiples médicos
- Interfaz web amigable con Streamlit

## Requisitos

- Python 3.10+
- API Key de Anthropic

## Instalación

```bash
# Clonar el repositorio
git clone <repo-url>
cd Creador-de-auditorias-reportes

# Crear entorno virtual
python -m venv venv
source venv/bin/activate   # Linux/Mac
venv\Scripts\activate      # Windows

# Instalar dependencias
pip install -r requirements.txt

# Configurar API Key
cp .env.example .env
# Edita .env y agrega tu ANTHROPIC_API_KEY
```

## Uso

```bash
streamlit run app.py
```

Luego abre http://localhost:8501 en tu navegador.

## Archivos de Entrada

Carga los siguientes archivos Excel en la interfaz:

| Archivo | Descripción |
|---------|-------------|
| `Clasificación de no conformidades.XLSX` | Clasificación de criterios como NC o ER |
| `TIPIFICACIONES DE USO COMUN EN AUDITORIA MEDICA DE CALIDAD.XLSX` | Lista de tipificaciones oficiales |
| `🕸️ Graficas de Cumplimiento 2026 \|\| MEDGEN \|\| DoctorSV` | Porcentajes de cumplimiento por médico/período |
| `🛢️ Base de Datos x Cita 2026 \|\| DoctorSv` | Consultas auditadas con IDs y hallazgos |

## Estructura del Proyecto

```
├── app.py                  # Aplicación Streamlit principal
├── data_loader.py          # Carga y procesamiento de archivos Excel
├── claude_analyzer.py      # Integración con Claude AI
├── report_generator.py     # Generación de documentos Word (.docx)
├── requirements.txt        # Dependencias Python
├── .env.example            # Plantilla de variables de entorno
└── README.md               # Este archivo
```

## Categorías de Color (Cumplimiento)

| Color | Rango | Categoría |
|-------|-------|-----------|
| 🟢 Verde | 98% - 100% | Excelente |
| 🟡 Amarillo | 95% - 97% | Muy Bueno |
| 🟠 Anaranjado | 85% - 94% | Aceptable |
| 🔴 Rojo | < 85% | Oportunidad de Mejora |

## Estructura del Reporte

1. **Encabezado** – Médico, código, período, especialidad y fecha
2. **Análisis Cuantitativo** – Tabla ID × Diagnóstico × NC × ER
3. **Análisis Cualitativo**
   - No Conformidades por componente y criterio
   - Eventos de Riesgo por componente y criterio
4. **Resumen Ejecutivo** – Párrafo ejecutivo + hallazgos por componente
5. **Cuadro de Cumplimiento** – Porcentajes exactos del archivo de gráficas
6. **Comentario de Seguimiento** – Comparación entre períodos (si aplica)
