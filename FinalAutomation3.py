import win32com.client as client
import pygsheets
import sys
import time
import xlsxwriter
import pythoncom

#Setting up the connection to Outlook
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
account = namespace.Folders['macquarie.instructions@legalstream.com.au']
inbox = account.Folders['Inbox']

#OLD GLOBAL VARIABLES, IF CODE DOESN'T WORK, USE THIS:
# messages = inbox.Items
# message = messages.GetLast()
#Empty List to start storing subject line and time and date.

emptylist2 = []
emptylist3 = []
dataList2 = []


# #Excel Automation for testing purposes on an excel sheet before implementing it on a google sheet.
# workbook = xlsxwriter.Workbook(r'D:\Projects for SQL & Tableau\Python Project\instructions.xlsx')
# worksheet = workbook.add_worksheet()
# worksheet.write(0, 0, "Date and Time")
# worksheet.write(0, 1, "APP Number")

#Google Sheet Automation:
scope = ['https://www.googleapis.com/auth/spreadsheets,' 'https://www.googleapis.com/auth/drive.file',
         'https://www.googleapis.com/auth/drive']
service_file = r'C:\Users\LEGUser\OneDrive\Desktop\Projects\data-processor-336022-951e72ecd818.json'
gc = pygsheets.authorize(service_file=service_file)
Worksheet = gc.open("Macquarie Doc Prep Pipeline")
Status_Sheet = Worksheet.worksheet_by_title("Status")
Bot_Review_Sheet = Worksheet.worksheet_by_title("Bot Review")

def Email_Extraction():
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", False)
    for msg in messages:
        if msg.UnRead == True:
            emptyList = []
            y = str(msg.Subject.encode('ascii', 'ignore')[-12:])
            p = str(y[-12:-1])
            date = msg.SentOn.strftime("%d/%m/%Y %H:%M:%S")
            print(p, date)
            emptyList.insert(0, date)
            emptyList.insert(1, p)
            emptyList.insert(2, "2. Bots Running")
            emptylist2.append(emptyList)
            emptylist3.append(emptyList[0:2])
            msg.UnRead = False

    print("Total Instructions to be processed: {}".format(len(emptylist2)))

def findEmptyCell_status():
    list_of_lists = Status_Sheet.get_col(2)
    for emptycell in range(1,len(list_of_lists)):
        if len(list_of_lists[emptycell]) == 0:
            return emptycell
            break

def checkEmptyCell_Status():
    list_of_lists = Status_Sheet.get_col(2)
    emptycell = findEmptyCell_status()
    total_files = len(emptylist2)
    for i in range(emptycell, emptycell+total_files):
        if len(list_of_lists[i]) != 0:
            print("SYSTEM EXIT: There is a row already filled in at Row {}".format(i+1))
            sys.exit()
    print("The bot has started writing at Row {} on the Status Tab".format(emptycell+1))

def findEmptyCell_Bot_Review():
    list_of_lists = Bot_Review_Sheet.get_col(3)
    for emptycell in range(1, len(list_of_lists)):
        if len(list_of_lists[emptycell]) == 0:
            print("The bot has started writing at Row {} on the Bot_Review Tab".format(emptycell+1))
            return emptycell
            break

def writeEmptyCell_status():
    B_row = 1
    emptycell = findEmptyCell_status()
    x = str(emptycell + B_row)
    Status_Sheet.update_values('B{}'.format(x), emptylist2)

def writeEmptyCell_Bot_Review():

    C_row = 1
    emptycell = findEmptyCell_Bot_Review()
    x = str(emptycell + C_row)
    Bot_Review_Sheet.update_values('C{}'.format(x), emptylist3)


class Handler_Class(object):
    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            try:
                emptylist2.clear()
                emptylist3.clear()
                dataList2.clear()
                Email_Extraction()
                checkEmptyCell_Status()
                writeEmptyCell_status()
                writeEmptyCell_Bot_Review()
                time.sleep(5)
            except:
                pass


outlook = client.DispatchWithEvents("Outlook.Application", Handler_Class)

#and then an infinite loop that waits for events.
pythoncom.PumpMessages()





