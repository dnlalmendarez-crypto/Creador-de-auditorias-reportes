"""
Period utility module.
Generates audit period options based on biweekly (quincena) ranges for a given year.
"""

MONTH_NAMES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

# Days per month (non-leap year); February handled separately for leap years
DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


def is_leap_year(year: int) -> bool:
    return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)


def get_days_in_month(month: int, year: int) -> int:
    """Return number of days in a given month (1-indexed)."""
    if month == 2 and is_leap_year(year):
        return 29
    return DAYS_IN_MONTH[month - 1]


def generate_periods(year: int) -> list[dict]:
    """
    Generate all biweekly audit periods for a given year.
    Period 1: 01 al 15 de enero
    Period 2: 16 al 31 de enero
    Period 3: 01 al 15 de febrero
    ...and so on.

    Returns list of dicts with keys: number, label, start_day, end_day, month, month_name, year
    """
    periods = []
    period_num = 1

    for month_idx in range(12):
        month = month_idx + 1
        month_name = MONTH_NAMES_ES[month_idx]
        days = get_days_in_month(month, year)

        # First half: 1-15
        periods.append({
            "number": period_num,
            "label": f"Periodo {period_num}: 01 al 15 de {month_name} {year}",
            "short_label": f"01 al 15 de {month_name} {year}",
            "start_day": 1,
            "end_day": 15,
            "month": month,
            "month_name": month_name,
            "year": year,
        })
        period_num += 1

        # Second half: 16-end
        periods.append({
            "number": period_num,
            "label": f"Periodo {period_num}: 16 al {days} de {month_name} {year}",
            "short_label": f"16 al {days} de {month_name} {year}",
            "start_day": 16,
            "end_day": days,
            "month": month,
            "month_name": month_name,
            "year": year,
        })
        period_num += 1

    return periods
