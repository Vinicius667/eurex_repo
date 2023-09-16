from utils import *
from variables import *
from utils import parse_excel

dict_auswahl_prefix = {0 : "", 1 : "STOXX_"}
auswahl = 1
excel_path = "./excel_files/STOXX_OI-Abfrage 2023-09-11 V9.00.xlsm"


auswahl, ZentralKurs, volatility, InterestRate, tage_bis_verfall, nbd_dict, dict_prod_bus, stock_price, expiry, expiry_1, heute, list_email_send_selection, future_date_col = parse_excel(auswahl, excel_path)

Summery_df, HedgeBedarf_df, HedgeBedarf1_df, delta = hedge(auswahl, ZentralKurs, volatility, InterestRate, tage_bis_verfall, dict_prod_bus, stock_price, expiry, expiry_1, heute)