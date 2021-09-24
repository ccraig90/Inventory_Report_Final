import win32com.client
import config as config
from win32com.client import Dispatch, constants
const=win32com.client.constants

def Send_Email():
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "Inventory Report"
    newMail.To = config.Email_Address_List
    newMail.BCC = config.BCC_Email_Addres_List
    newMail.Body = 'Inventory Report'
    newMail.HTMLBody ='<a href="file:///P:/2.%20Corps/PNL_Daily_Report/Reports/PNL_Report_1.xlsx">Inventory Report</a>'
    newMail.Attachments.Add(config.Excel_File_Address)
    newMail.Send()
