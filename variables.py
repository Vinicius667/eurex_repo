import os
from datetime import datetime
import json

# User defined values

Spannweite = 2000
ZentralKurs = 16500
Schritt = 50
DetailSchritte = 20

auswahl_dict = {
        0 : "DAX",
        1 : "STOXX"
    }

# Folders where results will be stored
current_results_path = "./Results"
temp_results_path = "./Results/temp"
old_results_path = "./Results/Old"

# Source file folder
src_path = "./src/"

# Source files
src_html = os.path.join(src_path,"template.html")
src_css = os.path.join(src_path,"style.css")

email_messages_path = os.path.join(src_path,"email_messages")

list_emails_path = os.path.join(src_path,"list_emails_test_Dieball.json")
with open(list_emails_path) as file:
    list_emails = json.load(file)['emails']


list_emails_providers_path = os.path.join(src_path,"email_providers.json")
with open(list_emails_providers_path) as file:
    list_emails_providers = json.load(file)['providers']

list_creditials_path = os.path.join(src_path,"credentials.json")
with open(list_creditials_path) as file:
    list_credentials = json.load(file)['credentials']


# Results files
list_pdf_files = [
    os.path.join(current_results_path, "eurex.pdf"),
    os.path.join(old_results_path, f"eurex_{datetime.today().strftime('%Y_%m_%d')}.pdf")
]
    
list_excel_files = [
    os.path.join(current_results_path, "eurex.xlsx"),
    os.path.join(old_results_path, f"eurex_{datetime.today().strftime('%Y_%m_%d')}.xlsx")
]

result_html = os.path.join(current_results_path,"eurex.html")
result_css = os.path.join(current_results_path,"style.css")
result_image = os.path.join(current_results_path,"image.svg")
