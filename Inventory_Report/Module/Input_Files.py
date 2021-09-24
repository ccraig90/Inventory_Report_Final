import pandas as pd
import config as config
import win32com.client
import Module.Date_File_Path as date_file_path
import shutil

def Read_File_from_Location(File_Path,Sheet_Name,Skip_Rows):
    File = pd.read_excel(File_Path,sheet_name = Sheet_Name,skiprows = Skip_Rows,index_col=None)
    return File


def TW_Email_Input(Email_Subject_Line):

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    message = messages.GetFirst()
    date = message.SentOn.strftime("%m-%d-%y")
    i=0
    for item in Email_Subject_Line:
        Subject_Line = Email_Subject_Line[item]
        while message.Subject != Subject_Line:
            message = messages.GetNext()
        else:
            attachments = message.Attachments
            print(message.Subject)
            attachment = attachments.Item(2)
            x = str(i)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project\\BloombergTW'+x+'.csv')
            date = message.SentOn.strftime("%m-%d-%y")
            message = messages.GetFirst()
            i += 1
    TW_Excel_File0 = pd.read_csv('C:/Users/ccraig/Desktop/PNL Project/BloombergTW0.csv',skiprows=3)
    TW_Excel_File1 = pd.read_csv('C:/Users/ccraig/Desktop/PNL Project/BloombergTW1.csv',skiprows=3)
    TW_Excel_File2 = pd.read_csv('C:/Users/ccraig/Desktop/PNL Project/BloombergTW2.csv',skiprows=3)
    TW_Excel_File3 = pd.read_csv('C:/Users/ccraig/Desktop/PNL Project/BloombergTW3.csv',skiprows=3)

    Bloomberg_Inventory = pd.concat([TW_Excel_File0,TW_Excel_File1,TW_Excel_File2,TW_Excel_File3],sort = True)
    Bloomberg_Inventory.dropna(subset = ['Symbol'],inplace = True)
    Bloomberg_Inventory = Bloomberg_Inventory[Bloomberg_Inventory.Security != 'USD Total']
    Bloomberg_Inventory.to_excel(str(config.BBG_File_Path)+date+'.xlsx')
    return Bloomberg_Inventory
    

def HT_FTP_File_Input(Sheet_Name,Skip_Rows):
    Recent_Date = date_file_path.Date_Recent()
    Previous_Date = date_file_path.Date_Previous()
    HT_File_String_Recent = date_file_path.File_Path(Recent_Date,config.HT_File_Path)
    HT_File_String_Previous = date_file_path.File_Path(Previous_Date,config.HT_File_Path)
    Recent_Hilltop_File = Read_File_from_Location(HT_File_String_Recent,Sheet_Name,Skip_Rows)
    Previous_Hilltop_File = Read_File_from_Location(HT_File_String_Previous,Sheet_Name,Skip_Rows)
    return Recent_Hilltop_File, Previous_Hilltop_File

def HT_FT_File_Archive():
    Recent_Date = date_file_path.Date_Recent()
    HT_File_String_Recent = date_file_path.File_Path(Recent_Date,config.HT_File_Path)
    shutil.copyfile(HT_File_String_Recent,'P:/2. Corps/PNL_Daily_Report/HT_Files/'+Recent_Date+'.xlsx')



