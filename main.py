from generate_files import *
from send_mail import send_mail_outlook, check_emails_available_outlook, client
from send_mail import send_mail


list_email_send = set(generate_parquets(heute= None))
print(list_email_send)

raise Exception("Stop here")
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
