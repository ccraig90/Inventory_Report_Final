
import datetime as datetime
import pandas as pd
import config as config
import Module.Date_File_Path as date_file_path
import Module.Input_Files as input_file
import Module.Summary as summary
import Module.Month_Summary as month_summary
import win32com.client
import xlsxwriter
import Module.Excel_Summary as excel_summary
import Module.Email_Send as email_send
import Module.Position_DSP as position_dsp

def Update_Cleared_Position():
    New_Updates = pd.read_excel(config.Excel_File_Address,sheet_name = 'Position DSP')
    New_Updates = New_Updates[['Security','Account Name','CUSIP','QTY DSP','Position Notes']]
    New_Updates.dropna(inplace = True)
    Cleared_Position_File = pd.read_excel(config.File_Path_Text['QTY_DSP_Cleared_File_Path'])
    New_Updates = Cleared_Position_File.append(New_Updates)
    New_Updates.drop_duplicates(subset = ['CUSIP'],keep = 'first',inplace = True)
    New_Updates.to_excel(config.File_Path_Text['QTY_DSP_Cleared_File_Path'],index = False)

    return New_Updates

def Update_Running_PnL_DSP(HT_Merged):
    # gets new DSP items loaded from saved PnL xlsx file on public drive
    PnL_DSP_Yesterday = pd.read_excel(config.Excel_File_Address,sheet_name = 'Real PnL DSP')  #DSP Items from previous report
    Additions_to_Running_PnL_DSP = PnL_DSP_Yesterday[['Date','Security','Account Name','CUSIP','Real PnL DSP','Notes']]  #DSP Items from previous report sorted
    Additions_to_Running_PnL_DSP = Additions_to_Running_PnL_DSP[Additions_to_Running_PnL_DSP['Notes'].isnull()]
    # gets running PnL items loaded from saved PnL xlsx file on public drive
    Yesterday_Running_PnL = PnL_DSP_Yesterday[['Unresolved PnL DSP', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13','Unnamed: 14']]
    Yesterday_Running_PnL.rename(columns={'Unresolved PnL DSP':'Date','Unnamed: 8':'Security','Unnamed: 9':'Account Name','Unnamed: 10':'CUSIP',
                                                                  'Unnamed: 11':'Previous PnL DSP','Unnamed: 12':'Current PnL DSP','Unnamed: 13':'Net PnL DSP','Unnamed: 14': 'Notes'},inplace = True)
    Yesterday_Running_PnL = Yesterday_Running_PnL.drop(Yesterday_Running_PnL.index[0])

    Yesterday_Running_PnL = Yesterday_Running_PnL[(Yesterday_Running_PnL['Net PnL DSP'] > 10) | (Yesterday_Running_PnL['Net PnL DSP'] < - 10)] # filters out 'closed' Positions


    Yesterday_Running_PnL = Yesterday_Running_PnL[Yesterday_Running_PnL['Notes'].isnull()]

    Yesterday_Running_PnL.rename(columns={'Net PnL DSP':'Real PnL DSP'},inplace = True)
    Yesterday_Running_PnL = Yesterday_Running_PnL[['Date','Security','Account Name','CUSIP','Real PnL DSP']]
    New_Running_PnL_List =[Yesterday_Running_PnL,Additions_to_Running_PnL_DSP]
    Current_Running_PnL_DSP = pd.concat(New_Running_PnL_List)
    Current_Running_PnL_DSP.rename(columns={'Real PnL DSP':'Previous PnL DSP'},inplace = True)


    Running_PnL_DSP = pd.merge(Current_Running_PnL_DSP,HT_Merged, on = 'CUSIP', how = 'left')
    
    Running_PnL_DSP.dropna(thresh = 4,inplace = True)
    Running_PnL_DSP.fillna(0,inplace = True)

    Running_PnL_DSP['Net PnL DSP'] = Running_PnL_DSP['Previous PnL DSP'] + Running_PnL_DSP['Real PnL DSP']
    Running_PnL_DSP.rename(columns={'Date_x':'Date','Security_x':'Security','Account Name_x':'Account Name','Real PnL DSP':'Current PnL DSP'},inplace = True)
    
    Running_PnL_DSP = Running_PnL_DSP[['Date','Security','Account Name','CUSIP','Previous PnL DSP','Current PnL DSP','Net PnL DSP']]

    Running_PnL_DSP.to_excel(config.File_Path_Text['Running_PnL_DSP_File_Path'],index = False)
 
    return Running_PnL_DSP


 