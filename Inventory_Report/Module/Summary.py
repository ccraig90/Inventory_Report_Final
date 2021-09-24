
import datetime as datetime
import config as config
import Module.Date_File_Path as date_file_path
import Module.Input_Files as input_file

def Hilltop_Trade_Acct_Summary():
    """
    Reads in "Summary Section" from Recent and Previous Hilltop FTP files.
    Reformats data, dropping unused data and adding a "Change" column.
    "Change" column = Recent - Previous
    """
    HT_Files = input_file.HT_FTP_File_Input('Trade Acct Summary',2)
    for item in HT_Files:
        item.drop(columns=['Office','RR','Unnamed: 7','Repo Adjustment','Account'],inplace = True)
    HT_Recent_Trade_Summary = HT_Files[0].transpose()
    HT_Previous_Trade_Summary = HT_Files[1].transpose()
    HT_Recent_Trade_Summary['Change'] = HT_Recent_Trade_Summary[0] - HT_Previous_Trade_Summary[0]
    HT_Recent_Trade_Summary.rename(columns={0:'Available Funds'},inplace = True)
    HT_Recent_Trade_Summary.rename_axis('Item',inplace = True)
    return HT_Recent_Trade_Summary