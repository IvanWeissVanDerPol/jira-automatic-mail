from distutils.log import info
import shutil
import re
import os
from pytest import skip
from workers import worker_class
from selenium import webdriver  
from selenium.webdriver.common.keys import Keys  
import time
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook
import numpy as np
from IPython.display import HTML
from datetime import date



pd.set_option('display.max_colwidth', -1)

##### functions #####
def log_in(browser,email,password):
    user_block_Xpath = '/html/body/div[1]/section/dashboard/div[1]/div/div[5]/div/div[2]/div/div/form/div[1]/input'
    pass_block_Xpath =  '/html/body/div[1]/section/dashboard/div[1]/div/div[5]/div/div[2]/div/div/form/div[2]/input'
    log_in_button = '/html/body/div[1]/section/dashboard/div[1]/div/div[5]/div/div[2]/div/div/form/div[4]/div/input'
    url = "http://192.168.1.3:8080/secure/Dashboard.jspa"
    browser.get(url)
    time.sleep(sleeptime)
    browser.find_element(By.XPATH,user_block_Xpath).send_keys(email)
    browser.find_element(By.XPATH,pass_block_Xpath).send_keys(password)
    browser.find_element(By.XPATH,log_in_button).click()
    time.sleep(sleeptime)
    
def monthToNum(shortMonth):
    return {
            'Jan': 1,
            'Feb': 2,
            'Mar': 3,
            'Apr': 4,
            'May': 5,
            'Jun': 6,
            'Jul': 7,
            'Aug': 8,
            'Sep': 9, 
            'Oct': 10,
            'Nov': 11,
            'Dec': 12
    }[shortMonth]
    
def make_clickable(value):
    url = "https://custom.url/{}"
    return '=HYPERLINK("%s", "%s")' % (url.format(value), value)

###### Email ######
sender_address = 'weissvanderpol@gmail.com'
sender_pass = 'Polivan123Gmail!'
subject = 'automated email here are all your open tickets'
mail_content = ""

####### paths to the different elements #######
ticket_page = "http://192.168.1.3:8080/issues/?jql=status%20in%20(Open%2C%20%22In%20Progress%22%2C%20Reopened%2C%20Resolved)%20ORDER%20BY%20due%20ASC%2C%20priority%20ASC%2C%20assignee%20ASC%2C%20status%20ASC%2C%20project%20ASC"
table_xpath =           "/html/body/div[1]/section/div[1]/div[3]/div/div/div/div/div/div/div[2]/div/issuetable-web-component/table/tbody"
user_path =             "/td[4]/span/a"
issue_ID_path =         "/html/body/div[1]/section/div[1]/div[3]/div/div/div/div/div/div/div[2]/div/issuetable-web-component/table/tbody/tr[1]/td[2]/a"
next_page_xpath =       "/html/body/div[1]/section/div[1]/div[3]/div/div/div/div/div/div/div[4]/div[2]/div/a[5]/span"

####### variables  #######
sleeptime = 3
workers = []
workers_found = []
workers_that_are_gone = ['Gaston Hendriks','Francisco Bogado','Arturo Fleitas','Marcelino Jara De Leon']


# browser = webdriver.Chrome("driver\\chromedriver.exe")  
# browser.maximize_window()  
# log_in(browser,"ivan","ivan")
# browser.get(ticket_page)

# ###############################################################################################################
# ########                        open jira and clean the columns fro the current project                 #######
# ###############################################################################################################
# time.sleep(sleeptime)
# list_of_columns = ["Key","Summary","Assignee","Status","Priority","Project","Due Date","Labels","Issue ID","Created"]
# jira_column_done_button_path =  '//*[@id="inline-dialog-column-picker-dialog"]/div[1]/div/form/div[4]/input'
# jira_first_columns_list_path = '//*[@id="user-column-sparkler-suggestions"]/div/ul[1]/li/label/input'
# jira_columns_seach_box_path = '//*[@id="user-column-sparkler-input"]'
# jira_show_column_button_path = '//*[@id="content"]/div[1]/div[3]/div/div/div/div/div/div/div[1]/div[2]/div/button'

# jira_show_column_button = browser.find_element(By.XPATH,jira_show_column_button_path) 
# jira_show_column_button.click()
# #remove the active columns
# while len(browser.find_elements(By.XPATH,'//*[@id="user-column-sparkler-suggestions"]/div/ul')) != 1:
#     jira_column_active_columns = browser.find_elements(By.XPATH,jira_first_columns_list_path)
#     jira_column_done_button = browser.find_element(By.XPATH,jira_column_done_button_path)
#     for active_column_jira in jira_column_active_columns:
#         active_column_jira.click()
#     jira_column_done_button.click()
#     browser.get(ticket_page)
#     jira_show_column_button = browser.find_element(By.XPATH,jira_show_column_button_path) 
#     jira_show_column_button.click()
# #add the correct columns
# for x in range(len(list_of_columns)):
#     jira_column_search = browser.find_element(By.XPATH,jira_columns_seach_box_path)
#     jira_column_search.send_keys(list_of_columns[x])
#     jira_column_deactivated_columns = browser.find_elements(By.XPATH,jira_first_columns_list_path)
#     jira_column_done_button = browser.find_element(By.XPATH,jira_column_done_button_path)
#     for deactivated_column_jira in jira_column_deactivated_columns:
#         column_name = deactivated_column_jira.accessible_name
#         if list_of_columns[x] == column_name:
#             deactivated_column_jira.click()
#             jira_column_done_button.click()
#             time.sleep(1.5)
#             jira_show_column_button.click()
#             break

# ###############################################################################################################
# ########                       Extract important info from jira and put in dataframe                    #######
# ###############################################################################################################

# issue_key_list = []
# link_list = []
# summary_list = []
# user_name_list = []
# status_list = []
# prio_list = []
# project_list = []
# due_date_list = []
# labelCell_list = []
# issue_ID_list = []
# creo_date_list = []

# # while not in last page (breaks when next page is not found)
# ticket_num = 0
# while True:
#     time.sleep(sleeptime)
#     ticket_table = browser.find_element(By.XPATH,table_xpath)
#     rows = ticket_table.find_elements(By.XPATH,'./tr')
#     for row in rows:
#         ticket_num      = ticket_num + 1 
#         issue_key       = row.find_element(By.XPATH,"./td[1]").text
#         summary         = row.find_element(By.XPATH,"./td[2]").accessible_name
#         user_name       = row.find_element(By.XPATH,"./td[3]").accessible_name
#         status          = row.find_element(By.XPATH,"./td[4]").accessible_name
#         prio            = re.sub("[^0-9]", "", row.find_element(By.XPATH,"./td[5]").accessible_name)
#         project         = row.find_element(By.XPATH,"./td[6]").accessible_name
#         due_date        = row.find_element(By.XPATH,"./td[7]").text.split("/")
#         labelCell       = row.find_element(By.XPATH,'./td[8]').text.split(" ")
#         issue_ID_cell   = row.find_element(By.XPATH,"./td[9]").text.split(" ")
#         creo_date       = row.find_element(By.XPATH,"./td[10]").text.split("/")
#         #change month name to number
#         if due_date[0] != " ":
#             due_date[1] = str(monthToNum(due_date[1]))
#             due_date = "-".join(due_date)
#         creo_date[1] = str(monthToNum(creo_date[1]))
        
#         creo_date = "-".join(creo_date)
        
#         if prio == "":
#             prio = "00"
#         real_prio = int(prio)
#         for letter in project:
#             letter_number_value = ord(letter)
#             real_prio = real_prio + letter_number_value
#         for label in labelCell:
#             for issue_ID in issue_ID_cell:
                
#                 issue_key_list.append(issue_key)
#                 summary_list.append(summary)
#                 user_name_list.append(user_name)
#                 status_list.append(status)
#                 prio_list.append(prio)
#                 project_list.append(project)
#                 due_date_list.append(due_date)
#                 labelCell_list.append(label)
#                 issue_ID_list.append(issue_ID)
#                 creo_date_list.append(creo_date)

                
#                 link_list.append("http://192.168.1.3:8080/browse/" + issue_key)
                
#         #!separates by worker and adds to list to make a txt for each worker
#         # if user_name not in workers_found and user_name not in workers_that_are_gone:
#         #     new_worker = worker_class(user_name)
#         #     new_worker.append_ticket(issue_key,summary,status,prio,due_date,ticket_num,real_prio)
#         #     workers.append(new_worker)
#         #     workers_found.append(user_name)
#         # else:
#         #     for worker in workers:
#         #         if worker.name == user_name:
#         #             worker.append_ticket(issue_key,summary,status,prio,due_date,ticket_num,real_prio)
#     if ticket_num < 51:
#         element = browser.find_element(By.XPATH,"/html/body/div[1]/section/div[1]/div[3]/div/div/div/div/div/div/div[4]/div[2]/div/a[5]")
#         element.click()
        
#     else:
#         try:
#             element = browser.find_element(By.XPATH,"/html/body/div[1]/section/div[1]/div[3]/div/div/div/div/div/div/div[4]/div[2]/div/a[6]")
#             element.click()
#         except Exception:
#             break
        
# #create the jira dataframe
# df_jira = pd.DataFrame({"Issue Key":issue_key_list,"Summary":summary_list,"Assignee":user_name_list,"Status":status_list,"Priority":prio_list,"Project":project_list,"Due date":due_date_list,"Label":labelCell_list,"Issue ID":issue_ID_list,"Creation date":creo_date_list,"Links":link_list})
# df_jira.to_excel("jira_data.xlsx",sheet_name = "jira tickets", index=False)
# df_jira.to_excel("jira_data_backup.xlsx",sheet_name = "jira tickets", index=False)
###############################################################################################################
########                        create a sheet for each worker                                          #######
###############################################################################################################
excel_path = "jira_data.xlsx"
excel_full_path = r"C:\Users\Ivan\Documents\projects\WPG\automation\automatically mail open projects\jira_data.xlsx"
backup_excel_full_path = r"C:\Users\Ivan\Documents\projects\WPG\automation\automatically mail open projects\jira_data_backup.xlsx"
todays_date = date.today()
min_creation_date = '01-01-'+str(todays_date.year)

if os.path.exists(excel_full_path):
    os.remove(excel_full_path)
shutil.copyfile(backup_excel_full_path,excel_full_path)
df_jira = pd.read_excel(excel_path)
df_jira = df_jira.replace("[' ']", '')
list_of_labels = list(filter(None, df_jira["Label"].drop_duplicates().str.findall(r'\b[a-zA-Z]{2}\b').to_list()))
ExcelWorkbook = load_workbook(excel_path)
writer = pd.ExcelWriter(excel_path, engine = 'openpyxl')
writer.book = ExcelWorkbook
for label in list_of_labels:
    label = label[0]
    new_worker_sheet = df_jira.copy()
    new_worker_sheet['Links'] = new_worker_sheet['Links'].apply(lambda x: '=HYPERLINK("{}", "{}")'.format(x, x))
    HTML(new_worker_sheet.to_html(escape=False))
    new_sheet_name = label + "-long"
    new_worker_sheet = new_worker_sheet[new_worker_sheet['Label']==(label)]
    new_worker_sheet.drop("Label",axis=1,inplace=True)
    new_worker_sheet.to_excel(writer,sheet_name = new_sheet_name, index=False)
    new_sheet_name = label + "-short"
    new_worker_sheet = new_worker_sheet[pd.to_datetime(new_worker_sheet['Creation date']) > pd.to_datetime(min_creation_date)]
    new_worker_sheet.to_excel(writer,sheet_name = new_sheet_name, index=False)
    
writer.save()
writer.close()










with open('emails.json') as json_file:
    data = json.load(json_file)

for worker in workers:
    for J_key in data:
        J_worker = data[J_key][0]
        j_name  = J_worker['name']
        j_email = J_worker['email']
        if worker.name == j_name:
            worker.mail = j_email

    mail_content= ""
    worker_name = worker.name
    worker_tickets = worker.ticket_list
    file_name = worker_name + ".txt" 
    worker_file = open(file_name, "w+")
    line = "priority          status                          Due Date                          issue_key                                                                         summary\n"
    worker_file.write(line)
    mail_content = mail_content + line
    for ticket_obj in worker_tickets:
        ticket = ticket_obj
        statusLen = len(ticket.status)
        keyLen = len(ticket.issue_key)
        dateLen = len(ticket.due_date)
        statusTabs = " " * (30 - statusLen)
        keyTabs = " " * (30 - keyLen)
        dateTabs = " " * (30 - dateLen)
        line = ticket.prio + "                " + ticket.status + statusTabs +  ticket.due_date + dateTabs + "http://192.168.1.3:8080/browse/" + ticket.issue_key + keyTabs + ticket.summary + "\n"
        worker_file.write(line)
        mail_content = mail_content + line
        
        
    #! delete this
    worker.mail = "weissvanderpol@gmail.com"
    receiver_address = worker.mail
    message = MIMEMultipart()
    subject = "Open tickets - " + worker_name
    message['Subject'] = subject
    message['From'] = sender_address
    message['To'] = worker.mail

    message.attach(MIMEText(mail_content, 'plain'))
    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login(sender_address, sender_pass) #login with mail_id and password
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    session.quit()
    print('Mail Sent')
    
    # if worker.name == "Unassigned":
    #     worker.mail = "jacquelin.fleytas@wpg.com.py"

    # if worker.mail == "emilio.ginzobenitez@wpg.com.py":

    #     receiver_address = worker.mail
    #     message = MIMEMultipart()
    #     message['From'] = sender_address
    #     message['To'] = worker.mail

    #     message.attach(MIMEText(mail_content, 'plain'))
    #     #Create SMTP session for sending the mail
    #     session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    #     session.starttls() #enable security
    #     session.login(sender_address, sender_pass) #login with mail_id and password
    #     text = message.as_string()
    #     session.sendmail(sender_address, receiver_address, text)
    #     session.quit()
    #     print('Mail Sent')

    # if worker.mail == "jacquelin.fleytas@wpg.com.py":
        
    #     receiver_address = worker.mail
    #     message = MIMEMultipart()
    #     message['From'] = sender_address
    #     message['To'] = worker.mail

    #     message.attach(MIMEText(mail_content, 'plain'))
    #     #Create SMTP session for sending the mail
    #     session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    #     session.starttls() #enable security
    #     session.login(sender_address, sender_pass) #login with mail_id and password
    #     text = message.as_string()
    #     session.sendmail(sender_address, receiver_address, text)
    #     session.quit()
    #     print('Mail Sent')
    worker_file.close()
    print(worker_name)
        




