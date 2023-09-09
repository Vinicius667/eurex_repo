import os
import numpy as np
import pandas as pd
from typing import List
import json
from datetime import datetime, timedelta
import pickle

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
    return ((range) * (col/(col.max() - col.min()) )).astype(np.int32)

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