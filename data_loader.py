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


def load_reporte_global(path: str) -> dict | None:
    """
    Load '📊 Reporte Global 2026 (Auditoria de Calidad) || DoctorSV'.
    Returns dict of sheets.
    """
    data = load_excel_safe(path)
    if data is None:
        return None
    if isinstance(data, dict):
        return data
    return {"Sheet1": data}


_NAME_COL_KEYWORDS = ["nombre", "name", "doctor", "médico", "medico"]
_CODE_COL_KEYWORDS = ["cod", "code", "código"]


def _is_code_col(col_name: str) -> bool:
    return any(kw in col_name.lower() for kw in _CODE_COL_KEYWORDS)


def _find_name_and_code(cols: list, first_row) -> tuple[str, str]:
    """Extract name and code from a row by scanning column names."""
    name = ""
    code = ""
    for col in cols:
        col_lower = str(col).lower()
        val = str(first_row[col]).strip()
        if not val or val.lower() in ("nan", "none"):
            continue
        if not code and _is_code_col(col_lower):
            code = val
        elif not name and any(kw in col_lower for kw in _NAME_COL_KEYWORDS):
            name = val
    return name, code


def extract_doctor_info_from_data(
    compliance_results: dict | None = None,
    consultations_df: "pd.DataFrame | None" = None,
) -> tuple[str, str]:
    """
    Extract the doctor's name and code from matched data rows.
    Returns (name, code) — empty strings if not found.
    """
    name = ""
    code = ""

    if compliance_results:
        for _sheet, result in compliance_results.items():
            rows = result.get("rows")
            if rows is None or rows.empty:
                continue
            n, c = _find_name_and_code(list(rows.columns), rows.iloc[0])
            if not name and n:
                name = n
            if not code and c:
                code = c
            if name and code:
                break

    if (not name or not code) and consultations_df is not None and not consultations_df.empty:
        n, c = _find_name_and_code(list(consultations_df.columns), consultations_df.iloc[0])
        if not name and n:
            name = n
        if not code and c:
            code = c

    return name, code


def find_doctor_global_compliance(
    reporte_global: dict,
    doctor_name: str,
    doctor_code: str | None = None,
) -> str:
    """
    Search for a doctor's component-level compliance in the Reporte Global.
    Returns markdown table text for the doctor's row(s).
    """
    name_lower = doctor_name.lower().strip()
    name_parts = [p for p in name_lower.split() if len(p) > 2]

    for sheet_name, df in reporte_global.items():
        if df is None or df.empty:
            continue

        df = df.copy()
        # Deduplicate columns
        seen = {}
        new_cols = []
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        df.columns = new_cols
        df_str = df.astype(str)
        cols = list(df.columns)

        matched = None

        # Search by code
        if doctor_code:
            code_upper = doctor_code.upper().strip()
            for col in cols:
                col_str = str(col).lower()
                if any(kw in col_str for kw in ["cod", "code", "código"]):
                    mask = df_str[col].str.strip().str.upper() == code_upper
                    if not mask.any():
                        mask = df_str[col].str.upper().str.contains(
                            code_upper, na=False, regex=False
                        )
                    if mask.any():
                        matched = df[mask]
                        break
            # Scan all columns for code
            if matched is None:
                for col in cols:
                    mask = df_str[col].str.strip().str.upper() == code_upper
                    if mask.any():
                        matched = df[mask]
                        break

        # Search by name (only if name is provided)
        if matched is None and name_lower:
            for col in cols:
                mask = df_str[col].str.lower().str.contains(
                    name_lower, na=False, regex=False
                )
                if mask.any():
                    matched = df[mask]
                    break

        # Name fragments
        if matched is None and len(name_parts) >= 2:
            for col in cols:
                col_lower = df_str[col].str.lower()
                match_count = sum(
                    col_lower.str.contains(part, na=False, regex=False).astype(int)
                    for part in name_parts
                )
                mask = match_count >= min(2, len(name_parts))
                if mask.any():
                    matched = df[mask]
                    break

        if matched is not None and not matched.empty:
            return f"**Hoja: {sheet_name}**\n{dataframe_to_markdown_table(matched)}"

    return ""


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
    # Build name fragments for partial matching (first name, last name)
    name_parts = [p for p in name_lower.split() if len(p) > 2]

    for sheet_name, df in graficas_data.items():
        if df is None or df.empty:
            continue

        # Deduplicate column names to avoid DataFrame returns on df[col]
        df = df.copy()
        seen = {}
        new_cols = []
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        df.columns = new_cols

        df_str = df.astype(str)
        cols = list(df.columns)

        # ── Strategy 1: Search by code ──
        if doctor_code and sheet_name not in results:
            code_upper = doctor_code.upper().strip()
            # Find code column
            code_col = None
            for col in cols:
                col_str = str(col).lower()
                if any(kw in col_str for kw in ["cod", "code", "código"]):
                    code_col = col
                    break
            if code_col:
                # Exact match
                mask = df_str[code_col].str.strip().str.upper() == code_upper
                if not mask.any():
                    # Partial match (code contained in cell)
                    mask = df_str[code_col].str.upper().str.contains(
                        code_upper, na=False, regex=False
                    )
                matches = df[mask]
                if not matches.empty:
                    results[sheet_name] = {
                        "rows": matches,
                        "period_col": None,
                        "sheet": sheet_name,
                    }

            # Also try scanning ALL cells for the code (some sheets may
            # not have a clearly-named code column)
            if sheet_name not in results:
                for col in cols:
                    mask = df_str[col].str.strip().str.upper() == code_upper
                    if mask.any():
                        results[sheet_name] = {
                            "rows": df[mask],
                            "period_col": None,
                            "sheet": sheet_name,
                        }
                        break

        # ── Strategy 2: Search by full name (exact substring) ──
        if sheet_name not in results and name_lower:
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

        # ── Strategy 3: Search by last name fragments ──
        if sheet_name not in results and len(name_parts) >= 2:
            # Try matching with last two name parts (apellidos)
            for col in cols:
                col_lower = df_str[col].str.lower()
                # Match if at least 2 name parts are found in the cell
                match_count = sum(
                    col_lower.str.contains(part, na=False, regex=False).astype(int)
                    for part in name_parts
                )
                mask = match_count >= min(2, len(name_parts))
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

        # Deduplicate column names to avoid DataFrame returns on df[col]
        df = df.copy()
        seen = {}
        new_cols = []
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        df.columns = new_cols

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

        # Fallback: search by name (only if name is provided)
        if matched is None and name_lower:
            for col in cols:
                mask = df_str[col].str.lower().str.contains(
                    name_lower, na=False, regex=False
                )
                if mask.any():
                    matched = df[mask].copy()
                    matched["_source_sheet"] = sheet_name
                    break

        if matched is not None:
            # Try to filter by period if provided
            if period and matched is not None and not matched.empty:
                period_str = str(period).strip()
                period_matched = None

                # Extract numeric part only from period codes like "P001", "P3"
                # (not from date strings like "01 al 15 de enero 2026")
                import re as _re
                is_period_code = bool(_re.match(r"^P?\d{1,3}$", period_str, _re.IGNORECASE))
                num_match = _re.search(r"\d+", period_str) if is_period_code else None
                period_num = int(num_match.group()) if num_match else None

                for col in list(matched.columns):
                    col_str = str(col).lower()
                    if any(kw in col_str for kw in ["periodo", "period", "quincena"]):
                        col_vals = matched[col].astype(str).str.strip()
                        # Try exact match first (e.g., "P001", "1")
                        mask = col_vals.str.upper() == period_str.upper()
                        if not mask.any() and period_num is not None:
                            # Try matching numeric value (column may have "1", "001", "P001", etc.)
                            col_nums = col_vals.str.extract(r"(\d+)", expand=False).astype(float)
                            mask = col_nums == period_num
                        if mask.any():
                            period_matched = matched[mask].copy()
                            break

                    # Also try date columns as fallback for date-mode periods
                    if period_matched is None and any(kw in col_str for kw in ["fecha", "date"]):
                        mask = matched[col].astype(str).str.lower().str.contains(
                            period_str.lower(), na=False, regex=False
                        )
                        if mask.any():
                            period_matched = matched[mask].copy()
                            break

                if period_matched is not None and not period_matched.empty:
                    matched = period_matched
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


def dataframe_to_markdown_table(df: pd.DataFrame, max_rows: int = 200) -> str:
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
