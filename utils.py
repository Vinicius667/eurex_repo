import os
import numpy as np
import pandas as pd
from typing import List
import win32com.client as client


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

def check_emails_available_outlook(outlook:client.CDispatch,list_emails:list)->bool:
    return_error = False
    list_emails_error = []
    for email in list_emails:
        if not outlook.Session.Accounts[email]:
            return_error = True
            list_emails_error.append(email)

    if return_error:
        print("E-mails not available:")
        for email in list_emails_error:
            print(email)
        raise ValueError(f"E-mails are not available on outlook.")

    return True


if __name__ == "__main__":
    # construct Outlook application instance
    print("running")
    outlook = client.Dispatch('Outlook.Application')

    if check_emails_available_outlook(outlook, [email['from'] for email in list_emails]):
        print("All e-mails to send from are available on the outlook session.")

    for email in list_emails:
        account = outlook.Session.Accounts[email['from']]
        # construct the email item object
        message = outlook.CreateItem(0)
        message.Subject = email['subject']
        message.Body = open(os.path.join(email_messages_path,email['text']),encoding="UTF-8").read()
        message.To = ";".join(email['mailto'])
        message._oleobj_.Invoke(*(64209, 0, 8, 0, account))
        message.Attachments.Add(os.path.join(os.getcwd(), 'requirements.txt'))
        #message.Display()
        message.Save()
        message.Send()