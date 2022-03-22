from distutils.log import info
import re
from workers import worker_class
from selenium import webdriver  
from selenium.webdriver.common.keys import Keys  
import time
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.webdriver.common.by import By


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

###### Email ######
sender_address = 'processmail@gmail.com'
sender_pass = 'processPass'
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
workers_that_are_gone = ['jira accounts user name and surname complete']


browser = webdriver.Chrome("driver\\chromedriver.exe")  
browser.maximize_window()  
log_in(browser,"jiraLogInUser","jiraLogInPass")

browser.get(ticket_page)
# while not in last page (breaks when next page is not found)
ticket_num = 0
while True:
    time.sleep(sleeptime)
    ticket_table = browser.find_element(By.XPATH,table_xpath)
    rows = ticket_table.find_elements(By.XPATH,'./tr')
    for row in rows:
        ticket_num = ticket_num + 1 
        # put jira columns in this order 
        issue_key = row.find_element(By.XPATH,"./td[1]").text
        summary   = row.find_element(By.XPATH,"./td[2]").accessible_name
        user_name = row.find_element(By.XPATH,"./td[3]").accessible_name
        status    = row.find_element(By.XPATH,"./td[4]").accessible_name
        prio = re.sub("[^0-9]", "", row.find_element(By.XPATH,"./td[5]").accessible_name)
        project   =  row.find_element(By.XPATH,"./td[6]").accessible_name
        due_date  = row.find_element(By.XPATH,"./td[7]").text
        if prio == "":
            prio = "00"
        real_prio = int(prio)
        for letter in project:
            letter_number_value = ord(letter)
            real_prio = real_prio + letter_number_value
        if user_name not in workers_found and user_name not in workers_that_are_gone:
            new_worker = worker_class(user_name)
            #new_worker.add_ticket(issue_key,summary,status,prio,ticket_num,real_prio)
            new_worker.append_ticket(issue_key,summary,status,prio,due_date,ticket_num,real_prio)
            workers.append(new_worker)
            workers_found.append(user_name)
        else:
            for worker in workers:
                if worker.name == user_name:
                    #worker.add_ticket(issue_key,summary,status,prio,ticket_num,real_prio)
                    worker.append_ticket(issue_key,summary,status,prio,due_date,ticket_num,real_prio)
    if ticket_num < 51:
        element = browser.find_element(By.XPATH,"/html/body/div[1]/section/div[1]/div[3]/div/div/div/div/div/div/div[4]/div[2]/div/a[5]")
        element.click()
        
    else:
        try:
            element = browser.find_element(By.XPATH,"/html/body/div[1]/section/div[1]/div[3]/div/div/div/div/div/div/div[4]/div[2]/div/a[6]")
            element.click()
        except Exception:
            break

with open('emails.json') as json_file:
    data = json.load(json_file)

for worker in workers:
    for J_key in data:
        J_worker = data[J_key][0]
        j_name  = J_worker['name']
        j_email = J_worker['email']
        if worker.name == j_name:
            worker.mail = j_email

    #worker.sort()
    mail_content= ""
    worker_name = worker.name
    worker_tickets = worker.ticket_list
    file_name = worker_name + ".txt" 
    worker_file = open(file_name, "w+")
    line = "priority          status                          Due Date                          issue_key                                                                         summary\n"
    worker_file.write(line)
    mail_content = mail_content + line
    for ticket_obj in worker_tickets:
        #!uncomment this
        # ticket = ticket_obj[2]
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
     
    if worker.name == "Unassigned":
        worker.mail = "admin mail"

        receiver_address = worker.mail
        message = MIMEMultipart()
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
    worker_file.close()
    print(worker_name)
        




