from send_mail import send_mail_outlook, check_emails_available_outlook, client
from variables import *
from utils import create_folder, parse_eurex, hedge, generate_pdfs, save_as_pickle
import shutil

heute_input = None
list_email_send = []

# Create results folder in case they still were not created
list_folder_results =  [current_results_path, old_results_path, temp_results_path]
for folder in list_folder_results:
    create_folder(folder)

for auswahl in [1,0]:
    result = parse_eurex(auswahl, heute = heute_input)

    if result:
        auswahl, ZentralKurs, volatility, InterestRate, tage_bis_verfall, nbd_dict, dict_prod_bus, stock_price, expiry, expiry_1, heute, list_email_send_selection, future_date_col = result

        list_email_send += list_email_send_selection

        save_as_pickle(result, os.path.join(temp_results_path, f'{dict_index_stock[auswahl]}.pickle'))

        Summery_df, HedgeBedarf_df, HedgeBedarf1_df, Ueberhaenge_df, delta= hedge(auswahl, ZentralKurs, volatility, InterestRate, tage_bis_verfall, dict_prod_bus, stock_price, expiry, expiry_1, heute)

        file_path = os.path.join(current_results_path, f"{dict_index_stock[auswahl]}.pdf")

        generate_pdfs(auswahl, Summery_df, HedgeBedarf_df, HedgeBedarf1_df, stock_price, heute, nbd_dict, tage_bis_verfall, delta, expiry, expiry_1, file_path)



list_email_send = set(list_email_send)

print(f"list_email_send = {list_email_send}")
outlook = client.Dispatch('Outlook.Application')  
for email in list_emails:
    if check_emails_available_outlook(outlook, email['from']):
        if email['id'] in list_email_send:
            send_mail_outlook(
                outlook = outlook,
                send_from = email['from'],
                send_to = email['mailto'],
                subject = email['subject'],
                body = open(os.path.join(email_messages_path,email['text']),encoding="UTF-8").read(),
                attachments = [
                        os.path.join(current_results_path, f"{email['index']}.pdf")
                                    ],
            )
            print(f"{email['from']}: File was sent.")
