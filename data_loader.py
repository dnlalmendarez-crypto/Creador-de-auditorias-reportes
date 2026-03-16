"""
Data loader module for medical audit report generator.
Handles reading and processing Excel files.
"""

import pandas as pd
import openpyxl
from pathlib import Path
from typing import Optional


def load_excel_safe(path: str, sheet_name=None) -> dict | pd.DataFrame | None:
    """Load an Excel file safely, returning None on failure."""
    try:
        if sheet_name is not None:
            return pd.read_excel(path, sheet_name=sheet_name)
        return pd.read_excel(path, sheet_name=None)  # Returns dict of all sheets
    except Exception as e:
        return None


def get_sheet_names(path: str) -> list[str]:
    """Return all sheet names from an Excel file."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception:
        return []


def load_clasificacion(path: str) -> pd.DataFrame | None:
    """
    Load 'Clasificación de no conformidades.XLSX'.
    Returns a DataFrame with columns for criteria and classification type.
    """
    data = load_excel_safe(path)
    if data is None:
        return None
    # If multiple sheets, try to find the most relevant one
    if isinstance(data, dict):
        # Use the first sheet or look for keywords
        for name, df in data.items():
            if any(kw in name.lower() for kw in ["clasif", "no conform", "riesgo"]):
                return df
        return list(data.values())[0]
    return data


def load_tipificaciones(path: str) -> pd.DataFrame | None:
    """
    Load 'TIPIFICACIONES DE USO COMUN EN AUDITORIA MEDICA DE CALIDAD.XLSX'.
    Returns a DataFrame with typification names.
    """
    data = load_excel_safe(path)
    if data is None:
        return None
    if isinstance(data, dict):
        return list(data.values())[0]
    return data


def load_graficas_cumplimiento(path: str) -> dict | None:
    """
    Load '🕸️ Graficas de Cumplimiento 2026 || MEDGEN || DoctorSV'.
    Returns dict of sheets (one per specialty).
    """
    data = load_excel_safe(path)
    if data is None:
        return None
    if isinstance(data, dict):
        return data
    return {"Sheet1": data}


def load_base_datos_cita(path: str) -> dict | None:
    """
    Load '🛢️ Base de Datos x Cita 2026 || DoctorSv'.
    Returns dict of sheets.
    """
    data = load_excel_safe(path)
    if data is None:
        return None
    if isinstance(data, dict):
        return data
    return {"Sheet1": data}


def find_doctor_compliance(
    graficas_data: dict,
    doctor_name: str,
    doctor_code: str | None = None,
    period: str | None = None,
) -> dict:
    """
    Search for a doctor's compliance data across all specialty sheets.
    Returns a dict with sheet_name, row data, and all available periods.
    """
    results = {}
    name_lower = doctor_name.lower().strip()

    for sheet_name, df in graficas_data.items():
        if df is None or df.empty:
            continue

        # Try to find columns that might be doctor name, code, period
        df_str = df.astype(str)
        cols = list(df.columns)

        # Search by code first if provided
        if doctor_code:
            code_col = None
            for col in cols:
                col_str = str(col).lower()
                if any(kw in col_str for kw in ["cod", "code", "código"]):
                    code_col = col
                    break
            if code_col:
                mask = df_str[code_col].str.upper() == doctor_code.upper()
                matches = df[mask]
                if not matches.empty:
                    if period:
                        # Try to filter by period
                        for col in cols:
                            if period.lower() in str(col).lower():
                                results[sheet_name] = {
                                    "rows": matches,
                                    "period_col": col,
                                    "sheet": sheet_name,
                                }
                                break
                        else:
                            results[sheet_name] = {
                                "rows": matches,
                                "period_col": None,
                                "sheet": sheet_name,
                            }
                    else:
                        results[sheet_name] = {
                            "rows": matches,
                            "period_col": None,
                            "sheet": sheet_name,
                        }

        # Fallback: search by name across all string columns
        if sheet_name not in results:
            for col in cols:
                mask = df_str[col].str.lower().str.contains(
                    name_lower, na=False, regex=False
                )
                if mask.any():
                    results[sheet_name] = {
                        "rows": df[mask],
                        "period_col": None,
                        "sheet": sheet_name,
                    }
                    break

    return results


def find_doctor_consultations(
    base_datos: dict,
    doctor_name: str,
    doctor_code: str | None = None,
    period: str | None = None,
) -> pd.DataFrame:
    """
    Extract all consultation records for a given doctor from the database.
    Returns a combined DataFrame of all relevant records.
    """
    frames = []
    name_lower = doctor_name.lower().strip()

    for sheet_name, df in base_datos.items():
        if df is None or df.empty:
            continue

        df_str = df.astype(str)
        cols = list(df.columns)

        matched = None

        # Search by code
        if doctor_code:
            for col in cols:
                col_str = str(col).lower()
                if any(kw in col_str for kw in ["cod", "code", "código", "medico"]):
                    mask = df_str[col].str.upper() == doctor_code.upper()
                    if mask.any():
                        matched = df[mask].copy()
                        matched["_source_sheet"] = sheet_name
                        break

        # Fallback: search by name
        if matched is None:
            for col in cols:
                mask = df_str[col].str.lower().str.contains(
                    name_lower, na=False, regex=False
                )
                if mask.any():
                    matched = df[mask].copy()
                    matched["_source_sheet"] = sheet_name
                    break

        if matched is not None:
            frames.append(matched)

    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame()


def extract_consultation_ids(df: pd.DataFrame) -> list[str]:
    """
    Extract consultation IDs (5-10 digit numbers) from the DataFrame.
    """
    ids = []
    for col in df.columns:
        col_str = str(col).lower()
        if any(kw in col_str for kw in ["id", "cita", "consulta", "folio", "numero"]):
            for val in df[col].dropna():
                val_str = str(val).strip()
                # Match 5-10 digit numbers
                import re
                if re.match(r"^\d{5,10}$", val_str):
                    ids.append(val_str)
    return list(set(ids))


def get_tipificaciones_list(df: pd.DataFrame) -> list[str]:
    """
    Extract a clean list of typification names from the tipificaciones DataFrame.
    """
    tipificaciones = []
    for col in df.columns:
        for val in df[col].dropna():
            val_str = str(val).strip()
            if len(val_str) > 3 and val_str not in tipificaciones:
                tipificaciones.append(val_str)
    return tipificaciones


def get_clasificacion_map(df: pd.DataFrame) -> dict:
    """
    Build a mapping of criteria -> classification type (No Conformidad / Evento de Riesgo).
    """
    clasificacion = {}
    if df is None or df.empty:
        return clasificacion

    # Try to identify columns
    cols = list(df.columns)
    criteria_col = None
    type_col = None

    for col in cols:
        col_lower = str(col).lower()
        if any(kw in col_lower for kw in ["criterio", "hallazgo", "tipific"]):
            criteria_col = col
        if any(kw in col_lower for kw in ["clasif", "tipo", "categ", "conformid", "riesgo"]):
            type_col = col

    if criteria_col and type_col:
        for _, row in df.iterrows():
            crit = str(row[criteria_col]).strip()
            typ = str(row[type_col]).strip()
            if crit and typ and crit != "nan" and typ != "nan":
                clasificacion[crit.lower()] = typ
    else:
        # Fallback: assume first col is criteria, second is type
        if len(cols) >= 2:
            for _, row in df.iterrows():
                crit = str(row[cols[0]]).strip()
                typ = str(row[cols[1]]).strip()
                if crit and typ and crit != "nan" and typ != "nan":
                    clasificacion[crit.lower()] = typ

    return clasificacion


def dataframe_to_markdown_table(df: pd.DataFrame, max_rows: int = 50) -> str:
    """Convert a DataFrame to a markdown table string for Claude."""
    if df is None or df.empty:
        return "(tabla vacía)"
    df_sample = df.head(max_rows)
    lines = []
    headers = list(df_sample.columns)
    lines.append("| " + " | ".join(str(h) for h in headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
    for _, row in df_sample.iterrows():
        cells = [str(v).replace("|", "/").strip() for v in row]
        lines.append("| " + " | ".join(cells) + " |")
    if len(df) > max_rows:
        lines.append(f"*(mostrando {max_rows} de {len(df)} filas)*")
    return "\n".join(lines)
