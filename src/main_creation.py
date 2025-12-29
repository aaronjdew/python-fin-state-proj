"""-------------------------------------------------------
    MAIN PROCEDURE
        -- Clears existing Excel data from various sheets
        -- Creates and inserts new data based on various
           rules set within the workbook itself
-------------------------------------------------------"""

import create_cal as cc
import create_daily_data as cd
import create_monthly_data as md
from utilities import create_bus_years

print("Starting process for data generation")

# Create Business Years
create_bus_years()
# Clear existing calendar
cc.clear_data_wkscldr()
# Create new calendar
cal_data = cc.insert_data_wkscldr()
# Clear existing daily data
cd.clear_data_wksfdai()
# Create daily data
dai_data = cd.insert_data_wksfdai(cal_data)
# Clear existing monthly data
md.clear_data_wksfmon()
# Create monthly data
md.insert_data_wksfmon(dai_data)
print("All data has been generated and loaded into the workbook.")
print("Process Complete!")
