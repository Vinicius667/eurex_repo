import os
from datetime import datetime
import json


save_files_to_debug = True

dict_index_stock = {
        0 : "DAX",
        1 : "STOXX"
    }

dict_stock_index = {v: k for k, v in dict_index_stock.items()}

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

list_emails_path = os.path.join(src_path,"list_emails_test.json")
with open(list_emails_path) as file:
    list_emails = json.load(file)['emails']


list_emails_providers_path = os.path.join(src_path,"email_providers.json")
with open(list_emails_providers_path) as file:
    list_emails_providers = json.load(file)['providers']

list_creditials_path = os.path.join(src_path,"credentials.json")
with open(list_creditials_path) as file:
    list_credentials = json.load(file)['credentials']


result_html = os.path.join(current_results_path,"eurex.html")
result_css = os.path.join(current_results_path,"style.css")
result_image = os.path.join(current_results_path,"image.svg")

