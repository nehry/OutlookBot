import win32com.client as client
import pygsheets
import sys
import time
import xlsxwriter

#Google Sheet Automation:tting up the connection to Outlook
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')
account = namespace.Folders['macquarie.instructions@legalstream.com.au']
inbox = account.Folders['Inbox']

#Empty List to start storing subject line and time and date
emptylist2 = []
emptylist3 = []
dataList2 = []

scope = ['https://www.googleapis.com/auth/spreadsheets,' 'https://www.googleapis.com/auth/drive.file',
         'https://www.googleapis.com/auth/drive']
service_file = r'D:\Projects for SQL & Tableau\Python Project\cryptotracker-327411-b7e2a6da147f.json'
gc = pygsheets.authorize(service_file=service_file)
Worksheet = gc.open("Macquarie Doc Prep Pipeline")
Status_Sheet = Worksheet.worksheet_by_title("Status")
Bot_Review_Sheet = Worksheet.worksheet_by_title("Bot Review")

#Email Extraction of the inbox for all unread files. Utilises string manipulation to extract the APP reference number and time library to extract the date and time. 
#Marks message as read once this function has been run.
def Email_Extraction():
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", False)
    for msg in messages:
        if msg.Unread == True:
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
            msg.Unread = False

    print("Total Instructions to be processed: {}".format(len(emptylist2)))

# This function finds the first empty row in the Status spreadsheet and stores it.
def findEmptyCell_status():
    list_of_lists = Status_Sheet.get_col(2)

    for emptycell in range(1,len(list_of_lists)):
        if len(list_of_lists[emptycell]) == 0:
            return emptycell
            break

# This function checks the next empty cells to see if they are empty depending on how many apps needs to be processed.
# If the rows are filled in, the program will exit to prevent overwriting data.
def checkEmptyCell_Status():
    list_of_lists = Status_Sheet.get_col(2)
    emptycell = findEmptyCell_status()
    total_files = len(emptylist2)
    for i in range(emptycell, emptycell+total_files):
        if len(list_of_lists[i]) != 0:
            print("SYSTEM EXIT: There is a row already filled in at Row {}".format(i+1))
            sys.exit()
    print("The bot has started writing at Row {} on the Status Tab".format(emptycell+1))


# This function finds the first empty row in the Bot Review spreadsheet and stores it.
def findEmptyCell_Bot_Review():
    list_of_lists = Bot_Review_Sheet.get_col(3)
    for emptycell in range(1, len(list_of_lists)):

        if len(list_of_lists[emptycell]) == 0:
            print("The bot has started writing at Row {} on the Bot_Review Tab".format(emptycell+1))
            return emptycell
            break

# This function will now process the matrix and batch update the rows all at once on the status sheet.
def writeEmptyCell_status():
    B_row = 1
    emptycell = findEmptyCell_status()
    x = str(emptycell + B_row)
    Status_Sheet.update_values('B{}'.format(x), emptylist2)
         
# This function will now process the matrix and batch update the rows all at once on the bot_review sheet.
def writeEmptyCell_Bot_Review():

    C_row = 1
    emptycell = findEmptyCell_Bot_Review()
    x = str(emptycell + C_row)
    Bot_Review_Sheet.update_values('C{}'.format(x), emptylist3)


Email_Extraction()
checkEmptyCell_Status()
writeEmptyCell_status()
writeEmptyCell_Bot_Review()
