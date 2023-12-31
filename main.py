from send_mail import send_mail_outlook, check_emails_available_outlook, client
from variables import *
from utils import create_folder, parse_eurex, hedge, generate_pdfs, save_as_pickle
import shutil

heute_input = None

# List of emails to send
list_email_send = []

# Create results folder in case they still were not created
list_folder_results =  [current_results_path, old_results_path, temp_results_path]
for folder in list_folder_results:
    create_folder(folder)

for auswahl in [0,1]:
    # Get result from parse_eurex function 
    result = parse_eurex(auswahl, heute = heute_input)

    if result: # If result is not None which can happen if, based on the emails to be sent, there is no need to parse Eurex

        # Unpack result from parse_eurex function
        auswahl, ZentralKurs, volatility, InterestRate, tage_bis_verfall, nbd_dict, dict_prod_bus, stock_price, expiry, expiry_1, heute, list_email_send_selection, future_date_col = result

        # Update list_email_send
        list_email_send += list_email_send_selection

        # Save result as pickle for testing purposes
        save_as_pickle(result, os.path.join(temp_results_path, f'{dict_index_stock[auswahl]}.pickle'))

        # Perform hedge calculation
        Summery_df, HedgeBedarf_df, HedgeBedarf1_df, Ueberhaenge_df, delta= hedge(auswahl, ZentralKurs, volatility, InterestRate, tage_bis_verfall, dict_prod_bus, stock_price, expiry, expiry_1, heute)

        # Generate pdfs
        generate_pdfs(auswahl, Summery_df, HedgeBedarf_df, HedgeBedarf1_df, stock_price, heute, nbd_dict, tage_bis_verfall, delta, expiry, expiry_1)

# Remove duplicates from list_email_send
list_email_send = set(list_email_send)

print(f"list_email_send = {list_email_send}")
outlook = client.Dispatch('Outlook.Application')
for email in list_emails:
    if check_emails_available_outlook(outlook, email['from']):
        if email['id'] in list_email_send:
            # File to be sent
            file_path = os.path.join(current_results_path, f"{email['index']}_{heute.strftime('%d_%m_%Y')}.pdf")
            print(f"{email['mailto']}: {file_path}")
            send_mail_outlook(
                outlook = outlook,
                send_from = email['from'],
                send_to = email['mailto'],
                subject = email['subject'],
                body = "\n".join([open(os.path.join(email_messages_path,text),encoding="UTF-8").read() for text in email['text']]),
                attachments = [
                        file_path
                                    ],
            )
            print(f"{email['from']}: File was sent.")