
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
import Module.System_Clear as system_clear
import numpy as np



def Position_DSP():
    Bloomberg_Inventory = input_file.TW_Email_Input(config.TOMS_Email_Subjects)
    Bloomberg_Inventory['Position'] = Bloomberg_Inventory['Position']*1000
    Bloomberg_Inventory['BBG MTG Position'] = Bloomberg_Inventory['Original Face Long P'] *1000
    Bloomberg_Inventory = Bloomberg_Inventory.groupby(['CUSIP']).agg({'P&L':'sum',
                                                                   'Security':'first',
                                                                   'Position':'sum',
                                                                   'Symbol':'first',
                                                                   'Book':'first',
                                                                   'BBG MTG Position':'sum'})
    HT_Files = input_file.HT_FTP_File_Input('Inv Detail',2)
    HT_Files_Recent = HT_Files[0].groupby(['CUSIP']).agg({'Quantity':'sum',
                                                                   'Unreal PnL':'sum',
                                                                   'Real PnL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Description':'first',
                                                                   'Price':'mean',
                                                                   'Account Name':'first'})
    HT_Files_Recent.reset_index(inplace = True)
    HT_Files_Recent['CUSIP'] = HT_Files_Recent['CUSIP'].str[:9]

    Cleared_Positions = pd.read_excel(config.File_Path_Text['QTY_DSP_Cleared_File_Path'])

    HT_Merged = pd.merge(HT_Files_Recent,Bloomberg_Inventory,on ='CUSIP',how = 'left')
    HT_Merged['Position'] = np.where((HT_Merged['Account Name'] == 'K76 S P IN'),HT_Merged['BBG MTG Position'],HT_Merged['Position'])
    HT_Merged['Position'] = np.where((HT_Merged['Account Name'] == 'M64 SIERRA'),HT_Merged['BBG MTG Position'],HT_Merged['Position'])
    HT_Merged['QTY DSP'] = HT_Merged['Position'] - HT_Merged['Quantity']
    HT_Merged = HT_Merged.loc[(HT_Merged['QTY DSP'] > 1) | (HT_Merged['QTY DSP'] < -1)]
    
    HT_Merged = HT_Merged.reindex(HT_Merged['QTY DSP'].abs().sort_values(ascending=False).index)
    HT_Merged.rename(columns={'Quantity':'HT Quantity','Position':'BBG Position','BBG Position':'BBG Pervious Position'},inplace = True)
    HT_Merged = HT_Merged.append(Cleared_Positions)
    Cleared_Positions = Cleared_Positions[['Security','Account Name','CUSIP','QTY DSP','Position Notes']]
    HT_Merged.loc[HT_Merged['Security']==0,'Security'] = HT_Merged['Description']
    HT_Merged = HT_Merged[['Security','Account Name','CUSIP','QTY DSP']]
    HT_Merged.drop_duplicates(subset = ['CUSIP'],keep = False,inplace = True)
    
    
    
    return HT_Merged,Cleared_Positions
    
    