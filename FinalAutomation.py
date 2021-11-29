import win32com.client as client
import pygsheets
import sys
import time
import xlsxwriter

#Setting up the connection to Outlook
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')
account = namespace.Folders['macquarie.instructions@legalstream.com.au']
inbox = account.Folders['Inbox']

#OLD GLOBAL VARIABLES, IF CODE DOESN'T WORK, USE THIS:
# messages = inbox.Items
# message = messages.GetLast()

#Empty List to start storing subject line and time and date
emptyList = []
timeList = []

# #Excel Automation for testing purposes on an excel sheet before implementing it on a google sheet.
# workbook = xlsxwriter.Workbook(r'D:\Projects for SQL & Tableau\Python Project\instructions.xlsx')
# worksheet = workbook.add_worksheet()
# worksheet.write(0, 0, "Date and Time")
# worksheet.write(0, 1, "APP Number")

#Google Sheet Automation:
scope = ['https://www.googleapis.com/auth/spreadsheets,' 'https://www.googleapis.com/auth/drive.file',
         'https://www.googleapis.com/auth/drive']
service_file = r'C:\Users\LEGUser\Desktop\Projects\cryptotracker-327411-b7e2a6da147f.json'
gc = pygsheets.authorize(service_file=service_file)
Worksheet = gc.open("Macquarie Doc Prep Pipeline")
Status_Sheet = Worksheet.worksheet_by_title("Status")
Bot_Review_Sheet = Worksheet.worksheet_by_title("Bot Review")

def Email_Extraction():
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", False)
    for msg in messages:
        if msg.Unread == True:
            y = str(msg.Subject.encode('ascii', 'ignore')[-12:])
            p = str(y[-12:-1])
            date = msg.SentOn.strftime("%d/%m/%Y %H:%M:%S")
            emptyList.append(p)
            timeList.append(date)
            msg.Unread = False

    print(emptyList)
    print(timeList)
    print("Total Instructions to be processed: {}".format(len(emptyList)))


# BELOW IS THE OLD CODE, IF SOMETHING GOES WRONG, USE THIS CODE AGAIN
    # Macq_Inst = [message for message in inbox.Items if message.Unread == True]
    # for msg in Macq_Inst:
    #     y = str(msg.Subject.encode('ascii', 'ignore')[-12:])
    #     p = str(y[-12:-1])
    #     emptyList.append(p)
    #     # msg.Unread = False
    # # emptyList.reverse()
    #
    # print(emptyList)
    #
    # for item in messages:
    #     if msg.Unread == True:
    #         date = item.SentOn.strftime("%d/%m/%Y %H:%M:%S")
    #         timeList.append(date)
    # # timeList.reverse()
    #
    # print(timeList)
    # print("Data exported to Google Sheets Successfully")

#OLD CODE, USE THIS TO FILTER FOR AN EMPTY CELL IF CURRENT CODE DOESN'T WORK.
    # str_list = list(filter(None, google_sheets.col_values(1)))
    # str_list2 = len(google_sheets.get_all_values()) + 1

def findEmptyCell_status():
    list_of_lists = Status_Sheet.get_col(2)

    for emptycell in range(1,len(list_of_lists)):
        if len(list_of_lists[emptycell]) == 0:
            print("The bot has started writing at Row {} on the Status Tab".format(emptycell+1))
            return emptycell
            break


    # list_of_lists = Status_Sheet.get_all_values()
    # for emptycell in range(1, len(list_of_lists)):
    #
    #     if len(list_of_lists[emptycell][1]) == 0:
    #         print("The bot has started writing at Row {} on the Status Tab".format(emptycell+1))
    #         return emptycell
    #         break

def findEmptyCell_Bot_Review():
    list_of_lists = Bot_Review_Sheet.get_all_values()
    for emptycell in range(1, len(list_of_lists)):

        if len(list_of_lists[emptycell][2]) == 0:
            print("The bot has started writing at Row {} on the Bot_Review Tab".format(emptycell+1))
            return emptycell
            break

def writeEmptyCell_status():
    B_row = 1
    C_row = 1
    D_row = 1
    emptycell = findEmptyCell_status()
    for dateandtime in timeList:
        if Status_Sheet.get_value("B{}".format(emptycell+B_row)) == "":
            x = str(emptycell + B_row)
            Status_Sheet.update_value('B{}'.format(x), dateandtime)
            B_row = B_row + 1
        else:
            print("There is a row already filled in at Row {}".format(emptycell + B_row))
            sys.exit()

    for APP in emptyList:
        x = str(emptycell + C_row)
        Status_Sheet.update_value('C{}'.format(x), APP)
        C_row = C_row + 1
        y = str(emptycell + D_row)
        Status_Sheet.update_value('D{}'.format(y), "2. Bots Running")
        D_row = D_row + 1

    # for Status in emptyList:
    #     x = str(emptycell + D_row)
    #     Status_Sheet.update_value('D{}'.format(x), "2. Bots Running")
    #     D_row = D_row + 1

def writeEmptyCell_Bot_Review():

    C_row = 1
    D_row = 1
    emptycell = findEmptyCell_Bot_Review()

    for dateandtime in timeList:
        x = str(emptycell + C_row)
        Bot_Review_Sheet.update_value('C{}'.format(x), dateandtime)
        C_row = C_row + 1

    for APP in emptyList:
        x = str(emptycell + D_row)
        Bot_Review_Sheet.update_value('D{}'.format(x), APP)
        D_row = D_row + 1


Email_Extraction()
writeEmptyCell_status()
writeEmptyCell_Bot_Review()

