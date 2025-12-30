"""-------------------------------------------------------
    CREATION OF FORECAST DATA (MONTHLY)
        -- Clears existing Data
        -- Uses Weekly data to create an aggregated view 
           for monthly data
-------------------------------------------------------"""

import warnings as wn
import numpy as np
import pandas as pd
import constants as cn


def clear_data_wksfmon():
    # --Initialisation--
    filepath = cn.FPEXCEL
    worksheet = cn.WKSFMON
    colstart = cn.WKSFMON_ST_COL_1
    colend = cn.WKSFMON_ED_COL_2
    df_fmon = pd.read_excel(filepath, worksheet,
                            usecols=range(colstart, colend))

    print('Clearing Monthly Data')

    # --Procedure--
    # Loop to re-initialise calender st/ed variables
    for aloop in range(1, 3):

        if aloop == 1:
            colst = cn.WKSFMON_ST_COL_1
            coled = cn.WKSFMON_ED_COL_1
        else:
            colst = cn.WKSFMON_ST_COL_2
            coled = cn.WKSFMON_ED_COL_2

        # set columns
        cols = range(colst, coled)

        # Ignore the data type warnings
        with wn.catch_warnings():
            wn.simplefilter("ignore", category=FutureWarning)

            # Update excel to missing data (NULLS)
            df_fmon.iloc[:, cols] = np.nan

            # Select part of df to export
            df_export = df_fmon.iloc[:, cols]

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

    print("Clearing Monthly Data Complete!")


def insert_data_wksfmon(df_fdai):
    # --Initialisation--
    filepath = cn.FPEXCEL
    worksheet = cn.WKSFMON
    neworder = cn.WKSFMON_NEW_ORD
    # New Col order
    df_fdai = df_fdai[neworder]

    print('Creating Monthly Data')

    # Distinct vals only - create new variable
    monthly_df = df_fdai.drop_duplicates().reset_index(drop=True)

    # Loop to re-initialise calender st/ed variables
    for aloop in range(1, 3):
        print(f'Exporting Monthly Data {aloop}')
        if aloop == 1:
            colst = cn.WKSFMON_ST_COL_1
            coled = cn.WKSFMON_ED_COL_1
            xlout_col = colst
        else:
            colst = cn.WKSFMON_ED_COL_1
            coled = monthly_df.shape[1]
            xlout_col = cn.WKSFMON_ST_COL_2

        cols = range(colst, coled)

        df_export = monthly_df.iloc[:, cols]

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

        print(f'Exporting Monthly Data {aloop} Complete!')
