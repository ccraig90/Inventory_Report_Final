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


def PnL_DSP():
    Item_Date = date_file_path.Date_Recent()
    Bloomberg_Inventory = input_file.TW_Email_Input(config.TOMS_Email_Subjects)
    Bloomberg_Inventory['Position'] = Bloomberg_Inventory['Position']*1000


    HT_Files = input_file.HT_FTP_File_Input('Inv Detail',2)
    HT_Files_Recent = HT_Files[0].groupby(['CUSIP']).agg({'Quantity':'sum',
                                                                   'Unreal PnL':'sum',
                                                                   'Real PnL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Description':'first',
                                                                   'Price':'mean',
                                                                   'Account Name':'first'})
    HT_Files_Previous = HT_Files[1].groupby(['CUSIP']).agg({'Quantity':'sum',
                                                                   'Unreal PnL':'sum',
                                                                   'Real PnL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Description':'first',
                                                                   'Price':'mean',
                                                                   'Account Name':'first'})
    HT_Merged = pd.merge(HT_Files_Recent,HT_Files_Previous,on = 'CUSIP',how = 'left')
    HT_Merged.reset_index(inplace = True)
    HT_Merged['CUSIP'] = HT_Merged['CUSIP'].str[:9]
    HT_Merged = pd.merge(HT_Merged,Bloomberg_Inventory,on ='CUSIP',how = 'left')
    HT_Merged.fillna(0,inplace = True)
    HT_Merged['HT Real PnL Change'] = HT_Merged['Real PnL_x'] - HT_Merged['Real PnL_y']
    
    HT_Merged['Real PnL DSP'] = HT_Merged['HT Real PnL Change'] - HT_Merged['P&L']
    HT_Merged.rename(columns={'Quantity_x':'HT Recent Position',
                              'Account Name_x':'Account Name',
                              'Quantity_y':'HT Previous Position',
                              'Real PnL_x':'HT Recent Real PnL',
                              'Real PnL_y':'HT Previous Real PnL',
                              'P&L':'Bloomberg PnL',
                              'Position':'Bloomberg Position'},inplace = True)

    HT_Merged['Date'] = Item_Date
    HT_Merged.loc[HT_Merged['Security']==0,'Security'] = HT_Merged['Description_x']

    HT_Merged = HT_Merged[['Date','Security','Account Name','CUSIP','Real PnL DSP']]
    HT_Merged = HT_Merged.loc[(HT_Merged['Real PnL DSP'] >= 25) | (HT_Merged['Real PnL DSP'] <= -25)]
    Running_PnL_DSP = system_clear.Update_Running_PnL_DSP(HT_Merged)
    HT_Merged = HT_Merged.reindex(HT_Merged['Real PnL DSP'].abs().sort_values(ascending=False).index)
    
    
    return HT_Merged,Running_PnL_DSP
