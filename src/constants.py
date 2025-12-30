"""-------------------------------------------------------
    PROJECT CONFIG FILE FOR CONSTANTS ETC..
-------------------------------------------------------"""

# --- File Path ---
FPEXCEL = "../pyFinancialStatements.xlsx"

# --- Sheets ---
WKSBD = "BusData"
WKSCLDR = "Calendar"
WKSFCALC = "ForeCalc"
WKSFDAI = "ForeDaily"
WKSFMON = "ForeMonth"
wksSa1 = "SalesForeYr1"
wksSa2 = "SalesForeYr2"
wksSa3 = "SalesForeYr3"
wksProp = "Property"
wksPr = "Promos"
wksExp = "Expenses"


# --- Sheet metadata ---

# --- Calendar metadata ---
# *Range columns in sheet
WKSCLDR_ST_COL_1 = 0
WKSCLDR_ED_COL_1 = 5
WKSCLDR_ST_COL_2 = 8
WKSCLDR_ED_COL_2 = 10

# --- Bus Data metadata ---
# *Actual columns in sheet
# Business Years
WKSBD_ED_COL_1 = 7
# Total Rooms
WKSBD_ED_COL_2 = 2

# *Actual rows in sheet
# Business Years
WKSBD_ED_ROW_1 = 14
# Total Rooms
WKSBD_ED_ROW_2 = 13

# --- Forecast Daily metadata ---
# New df order
WKSFDAI_NEW_ORD = ['Date', 'DayType', 'Month',
                   'Year', 'BusYear', 'BusMonth']
# *Range columns in sheet
WKSFDAI_ST_COL_1 = 0
WKSFDAI_ED_COL_1 = 4
WKSFDAI_ST_COL_2 = 7
WKSFDAI_ED_COL_2 = 9

# --- Forecast Monthly metadata ---
# New df order
WKSFMON_NEW_ORD = ['Month', 'Year', 'Room',
                   'BusYear', 'BusMonth']
# *Range columns in sheet
WKSFMON_ST_COL_1 = 0
WKSFMON_ED_COL_1 = 3
WKSFMON_ST_COL_2 = 7
WKSFMON_ED_COL_2 = 9
