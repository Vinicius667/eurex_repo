import os
import numpy as np
import pandas as pd
from typing import List
import json
from datetime import datetime, timedelta
import pickle
from openpyxl.reader.excel import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple

def create_folder(path):
    # Create a foder in case it has not been created already
    if not os.path.exists(path):
        os.mkdir(path)
        print(f"Directory created: {path}")


#  Auxiliary functions for styling the table

def scale_col_range(col : pd.Series, range):
    """Scale column to range
    col = [0,10,100], range = 10 => [0,1,10]
    """
    return ((range) * ((col - col.min())/(col.max() - col.min()) )).astype(np.int32)

def bold(col:pd.Series)->List:
    """
    Return list string in col in bold using HTML syntax
    """
    return [f"<b>{element}</b>" for element in col]

def posneg_binary_color(col:pd.Series,pos_color,neg_color):
    """
    Return series with colors to differentiate positive and negative values
    """
    aux = col.reset_index()
    aux['color'] = neg_color #
    aux.loc[col>=0,'color'] = pos_color
    return np.array(aux['color'])

def posneg_gradient(col:pd.Series):
    aux = col.reset_index()
    aux['r'] = aux['g']  = aux['b'] = 255
    aux.loc[aux.Änderung<0,'g'] = aux.loc[aux.Änderung<0,'b'] = 255 - 150*aux.Änderung/(aux.Änderung.min()) 
    aux.loc[aux.Änderung>0,'g'] = aux.loc[aux.Änderung>0,'r'] = 255 - 150*aux.Änderung/(aux.Änderung.max())
    aux = aux.astype(np.int32)
    return  np.array([f"rgb({aux.loc[i,'r']},{aux.loc[i,'g']},{aux.loc[i,'b']})" for i in range(aux.shape[0]) ])

#################################################################################################

def save_as_pickle(variable, path):
    with open(path, 'wb') as handle:
        pickle.dump(variable, handle, protocol=pickle.HIGHEST_PROTOCOL)

def read_pickle(path):
    with open(path, 'rb') as handle:
        variable = pickle.load(handle)
        return variable
    


def next_business_day(input_datetime):
    # Define a list of weekdays that are not considered business days (Saturday and Sunday)
    non_business_days = [5, 6]  # 5 represents Saturday, and 6 represents Sunday

    # Start with the next day from the input date
    next_day = input_datetime + timedelta(days=1)

    # Check if the next day is a non-business day (Saturday or Sunday)
    while next_day.weekday() in non_business_days:
        next_day += timedelta(days=1)

    return next_day


def parse_excel(auswahl : int, excel_path : str):
    dict_auswahl_prefix = {
        0 : "",
        1 : "STOXX_"
    }

    wb = load_workbook(excel_path, data_only=True, keep_vba=True)
    Ueberhaenge_sheet = wb['Ueberhaenge']
    heute = Ueberhaenge_sheet.cell(*coordinate_to_tuple("G5")).value
    expiry = Ueberhaenge_sheet.cell(*coordinate_to_tuple("G3")).value
    expiry_1 = Ueberhaenge_sheet.cell(*coordinate_to_tuple("R6")).value
    stock_price = Ueberhaenge_sheet.cell(*coordinate_to_tuple("X4")).value
    InterestRate = Ueberhaenge_sheet.cell(*coordinate_to_tuple("N3")).value
    ZentralKurs =  Ueberhaenge_sheet.cell(*coordinate_to_tuple("C3")).value  
    volatility = Ueberhaenge_sheet.cell(*coordinate_to_tuple("N4")).value

    Summery_sheet = wb[f'{dict_auswahl_prefix[auswahl]}Summery']
    nbd_heute = Summery_sheet.cell(*coordinate_to_tuple("C9")).value
    nbd_last = Summery_sheet.cell(*coordinate_to_tuple("C9")).value
    nbd_last = next_business_day(heute)
    nbd_dict = {'heute' : nbd_heute, 'last' : nbd_last}


    tage_bis_verfall = (expiry - heute).days #+ 1

    print(f"tage_bis_verfall = {tage_bis_verfall}")
    print(f"stock_price = {stock_price}")
    print(f"InterestRate = {InterestRate}")
    print(f"volatility = {volatility}")
    print(f"InterestRate = {InterestRate}")


    dict_prod_bus = {}
    for productdate_idx in [0,1]:
        dict_prod_bus[productdate_idx] = {}
        for busdate_idx in [0,1]:
            print(f"busdate_idx = {busdate_idx}   | productdate_idx = {productdate_idx}")
            if (busdate_idx == 1) and (productdate_idx == 1):
                # No need to request data for this case so we skip this iteration
                break

            if (busdate_idx == 0) and (productdate_idx == 0):
                contractsCall_aux_df =  pd.read_excel(excel_path, sheet_name = f"{dict_auswahl_prefix[auswahl]}Call_Front")
                contractsPut_aux_df =  pd.read_excel(excel_path, sheet_name = f"{dict_auswahl_prefix[auswahl]}Put_Front")

            if (busdate_idx == 1) and (productdate_idx == 0):
                contractsCall_aux_df= pd.read_excel(excel_path, sheet_name = f"{dict_auswahl_prefix[auswahl]}CallFront-1")
                contractsPut_aux_df = pd.read_excel(excel_path, sheet_name = f"{dict_auswahl_prefix[auswahl]}PutFront-1")

            if (busdate_idx == 0) and (productdate_idx == 1):
                contractsCall_aux_df = pd.read_excel(excel_path, sheet_name = f"{dict_auswahl_prefix[auswahl]}Put+01")
                contractsPut_aux_df= pd.read_excel(excel_path, sheet_name = f"{dict_auswahl_prefix[auswahl]}Call+01")


            aux_df = contractsCall_aux_df[['strike','openInterest']].merge(
                contractsPut_aux_df[['strike','openInterest']],
                on = "strike",
                how="left",
                suffixes=('_CF', '_PF')
            )
            dict_prod_bus[productdate_idx][busdate_idx] = aux_df
    return auswahl, heute, expiry, expiry_1, stock_price, InterestRate, volatility, ZentralKurs, tage_bis_verfall, nbd_dict, dict_prod_bus