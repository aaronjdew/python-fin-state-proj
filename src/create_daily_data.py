"""-------------------------------------------------------
    CREATION OF FORECAST DATA (WEEKLY)
        -- Clears existing data
        -- Uses Calendar data to create new weekly data
-------------------------------------------------------"""

import warnings as wn
import numpy as np
import pandas as pd
from openpyxl import load_workbook as lw
import constants as cn


def clear_data_wksfdai():
    # --Initialisation--
    filepath = cn.FPEXCEL
    worksheet = cn.WKSFDAI
    colstart = cn.WKSFDAI_ST_COL_1
    colend = cn.WKSFDAI_ED_COL_2
    df_fdai = pd.read_excel(filepath, worksheet,
                            usecols=range(colstart, colend))

    print('Clearing Daily Data')

    # --Procedure--
    # Loop to re-initialise calender st/ed variables
    for aloop in range(1, 3):

        if aloop == 1:
            colst = cn.WKSFDAI_ST_COL_1
            coled = cn.WKSFDAI_ED_COL_1
            coled = coled + 1
        else:
            colst = cn.WKSFDAI_ST_COL_2
            coled = cn.WKSFDAI_ED_COL_2

        # set columns
        cols = range(colst, coled)

        # Ignore the data type warnings
        with wn.catch_warnings():
            wn.simplefilter("ignore", category=FutureWarning)

            # Update excel to missing data (NULLS)
            df_fdai.iloc[:, cols] = np.nan

            # Select part of df to export
            df_export = df_fdai.iloc[:, cols]

            # Clear create data by setting to NULL
            with pd.ExcelWriter(
                    filepath,
                    engine='openpyxl',
                    mode='a',
                    if_sheet_exists='overlay') as writer:

                # Export data and convert to NULLs
                df_export.to_excel(writer, sheet_name=worksheet,
                                   index=False, startcol=colst, startrow=1,
                                   na_rep='', header=False)

    print("Clearing Daily Data Complete!")


def insert_data_wksfdai(df_fdai):
    # --Initialisation--
    filepath = cn.FPEXCEL
    worksheet = cn.WKSFDAI
    neworder = cn.WKSFDAI_NEW_ORD
    df_fdai = df_fdai[neworder]

    # --Initalise BusData variables--
    # Load workbook
    workbook = lw(filepath, data_only=True)
    wksbd = cn.WKSBD
    bdcol = cn.WKSBD_ED_COL_2
    bdrw = cn.WKSBD_ED_ROW_2
    no_of_rooms = workbook[wksbd].cell(bdrw, bdcol).value

    # Create temporary dataframe
    df_temp = df_fdai.copy(deep=True)

    print('Creating Daily Data')

    # Create calendar df for all rooms
    for room_num in range(1, no_of_rooms + 1):
        if room_num == 1:
            df_temp['Room'] = room_num
        else:
            df_temp1 = df_temp
            df_temp2 = df_fdai.copy(deep=True)
            df_temp2['Room'] = room_num
            df_temp = pd.concat([df_temp1, df_temp2], ignore_index=True)

    df_fdai = df_temp

    # Loop to re-initialise calender st/ed variables
    for aloop in range(1, 4):
        print(f'Exporting Daily Data {aloop}')
        if aloop == 1:
            colst = cn.WKSFDAI_ST_COL_1
            coled = cn.WKSFDAI_ED_COL_1
            xlout_col = colst

        elif aloop == 2:
            colst = cn.WKSFDAI_ED_COL_1
            coled = df_fdai.shape[1]
            coled = coled - 1
            xlout_col = cn.WKSFDAI_ST_COL_2
        else:
            colst = df_fdai.shape[1]
            colst = colst - 1
            coled = df_fdai.shape[1]
            xlout_col = cn.WKSFDAI_ED_COL_1

        cols = range(colst, coled)
        df_export = df_fdai.iloc[:, cols]

        # Add data to spreadsheet
        with pd.ExcelWriter(
                filepath,
                engine='openpyxl',
                mode='a',
                if_sheet_exists='overlay') as writer:

            # Export data
            df_export.to_excel(writer, sheet_name=worksheet,
                               index=False, startcol=xlout_col, startrow=1,
                               na_rep='', header=False)

        print(f'Exporting Daily Data {aloop} Complete!')

    return df_fdai
