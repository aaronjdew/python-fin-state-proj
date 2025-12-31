"""-------------------------------------------------------
    GLOBAL FUNCTIONS USED THROUGHOUT PROJECT
-------------------------------------------------------"""

import calendar as cal
from datetime import date
import constants as cn


def day_type(day: str) -> str:
    """---------------------------------------------------
    func: Grouping of day into day type
    ---------------------------------------------------"""

    if day not in ('Sat', 'Sun'):
        return 'Week'
    return day


def get_end_of_month(eomdate: date) -> date:
    """---------------------------------------------------
    func: Returns last day of month for argument passed
    ---------------------------------------------------"""

    year = eomdate.year
    month = eomdate.month
    _, num_days = cal.monthrange(year, month)
    end_date = date(year, month, num_days)
    return end_date


def business_year(act_date: date, workbook) -> int:
    """---------------------------------------------------
    func: Returns the business year (not actual year)
    i.e Business becomes operational in 2025 = Year 1
    ---------------------------------------------------"""

    wksbd = cn.WKSBD
    bdcol = cn.WKSBD_ED_COL_1
    bdrw = cn.WKSBD_ED_ROW_1

    # Year 1 Date/Val
    year_1_dt = workbook[wksbd].cell(bdrw, bdcol).value
    year_1 = workbook[wksbd].cell(bdrw, bdcol - 1).value
    # Year 2 Date/Val
    year_2_dt = workbook[wksbd].cell(bdrw + 1, bdcol).value
    year_2 = workbook[wksbd].cell(bdrw + 1, bdcol - 1).value
    # Year 3 Date/Val
    year_3_dt = workbook[wksbd].cell(bdrw + 2, bdcol).value
    year_3 = workbook[wksbd].cell(bdrw + 2, bdcol - 1).value

    if year_1_dt <= act_date < year_2_dt:
        return year_1
    if year_2_dt <= act_date < year_3_dt:
        return year_2
    return year_3


def business_month(act_date: date, workbook) -> int:
    """---------------------------------------------------
    func: Returns the business month (not actual month)
    i.e Business becomes operational in April 2026 = mon 1
    ---------------------------------------------------"""
    wksbd = cn.WKSBD
    bdcol = cn.WKSBD_ED_COL_1
    bdrw = cn.WKSBD_ED_ROW_1
    act_month = act_date.month

    # Set Bus Month 1
    year_1_dt = workbook[wksbd].cell(bdrw, bdcol).value
    bus_mon_1 = year_1_dt.month

    if act_month >= bus_mon_1:
        return (act_month - bus_mon_1) + 1
    return (act_month - bus_mon_1) + 13


if __name__ == "__main__":
    print("This module is intended for import only")
