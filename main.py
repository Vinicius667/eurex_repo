from generate_files import generate_files
from variables import *
import traceback
from send_mail import send_mail

# Amount of times they we will try to generate the files 
tries = 20
is_generation_successful = False


for i in range(tries):
    try:
        generate_files(0)
        is_generation_successful = True
        break
    except:
        pass


if is_generation_successful:
    dict_conditions = {
        "always" : True,
    }


    for email in list_emails:
        credentials = list_credentials[email['from']]
        smtp_config = list_emails_providers[credentials['provider']]['smtp']
        if dict_conditions[email['condition']]:
            send_mail(
                send_from = email['from'],
                send_to = email['mailto'],
                subject = email['subject'],
                body = open(os.path.join(email_messages_path,email['text']),encoding="UTF-8").read(),
                attachments = [list_pdf_files[0]],
                smtp_url = smtp_config['server'],
                port = smtp_config['port'],
                username = email['from'], 
                password = credentials['password']
            )
else:
    print("Emails were not sent since PDF file was not generated.")