"""-------------------------------------------------------
    PROJECT CONFIG FILE FOR CONSTANTS ETC..
-------------------------------------------------------"""

from typing import Final

# --- File Path ---
FPEXCEL: Final[str] = "data/pyFinState.xlsx"

# --- Sheets ---
WKSBD: Final[str] = "BusData"
WKSCLDR: Final[str] = "Calendar"
WKSFCALC: Final[str] = "ForeCalc"
WKSFDAI: Final[str] = "ForeDaily"
WKSFMON: Final[str] = "ForeMonth"

# --- Sheet metadata ---

# --- Calendar metadata ---
# *Range columns in sheet
WKSCLDR_ST_COL_1: Final[int] = 0
WKSCLDR_ED_COL_1: Final[int] = 5
WKSCLDR_ST_COL_2: Final[int] = 8
WKSCLDR_ED_COL_2: Final[int] = 10

# --- Bus Data metadata ---
# *Actual columns in sheet
# Business Years
WKSBD_ED_COL_1: Final[int] = 2
# Total Rooms
WKSBD_ED_COL_2: Final[int] = 2

# *Actual rows in sheet
# Business Years
WKSBD_ED_ROW_1: Final[int] = 17
# Total Rooms
WKSBD_ED_ROW_2: Final[int] = 13

# --- Forecast Daily metadata ---
# New df order
WKSFDAI_NEW_ORD: Final[list[str]] = ['Date', 'DayType', 'Month',
                                     'Year', 'BusYear', 'BusMonth']
# *Range columns in sheet
WKSFDAI_ST_COL_1: Final[int] = 0
WKSFDAI_ED_COL_1: Final[int] = 4
WKSFDAI_ST_COL_2: Final[int] = 7
WKSFDAI_ED_COL_2: Final[int] = 9

# --- Forecast Monthly metadata ---
# New df order
WKSFMON_NEW_ORD: Final[list[str]] = ['Month', 'Year', 'Room',
                                     'BusYear', 'BusMonth']
# *Range columns in sheet
WKSFMON_ST_COL_1: Final[int] = 0
WKSFMON_ED_COL_1: Final[int] = 3
WKSFMON_ST_COL_2: Final[int] = 7
WKSFMON_ED_COL_2: Final[int] = 9
