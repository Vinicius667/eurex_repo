from generate_files import generate_files
from variables import *
from send_mail import send_mail_outlook, check_emails_available_outlook, client
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
    outlook = client.Dispatch('Outlook.Application')


    dict_conditions = {
        "always" : True,
    }

    for email in list_emails:
        if check_emails_available_outlook(outlook, email['from']):
            if dict_conditions[email['condition']]:
                send_mail_outlook(
                    outlook = outlook,
                    send_from = email['from'],
                    send_to = email['mailto'],
                    subject = email['subject'],
                    body = open(os.path.join(email_messages_path,email['text']),encoding="UTF-8").read(),
                    attachments = [list_pdf_files[0]],
                )
                print(f"{email['from']}: File was sent.")


else:
    print("Emails were not sent since PDF file was not generated.")