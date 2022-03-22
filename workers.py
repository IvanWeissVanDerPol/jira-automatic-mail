import heapq
from datetime import datetime
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
class ticket:
    def __init__(self,issue_key,summary,status,prio,due_date):
        self.issue_key =  issue_key
        self.summary   =  summary
        self.status    =  status
        self.prio      =  prio
        self.due_date  =  due_date


class worker_class:

    def __init__(self,name):
        self.name = name
        self.mail = ""
        self.number_of_tickets = 0
        self.ticket_list = []
    
    def append_ticket(self,issue_key,summary,status,prio,due_date,ticket_num,real_prio):
        pyDate = present = ""
        if due_date != " ":
            splitDueDate = due_date.split("/")
            pyDate = datetime(int(splitDueDate[2]), monthToNum(splitDueDate[1]), int(splitDueDate[0]))
            present = datetime.now()
        if not pyDate > present or due_date == " ":
            new_ticket = ticket(issue_key,summary,status,prio,due_date)
            self.ticket_list.append(new_ticket)
            self.number_of_tickets += 1
            
    def sort(self):
        sorted_list = sorted(self.ticket_list, key=lambda x: (x.prio, x.issue_key))
        print(sorted_list)
    
