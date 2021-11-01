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
import numpy as np

def HT_Detail_Generate():
    Bloomberg_Inventory = input_file.TW_Email_Input(config.TOMS_Email_Subjects)
    Bloomberg_Inventory['Position'] = Bloomberg_Inventory['Position']*1000
    Bloomberg_Inventory['BBG MTG Position'] = Bloomberg_Inventory['Original Face Long P'] *1000
    
    Bloomberg_Inventory.reset_index(inplace = True)

    Bloomberg_Inventory = Bloomberg_Inventory.reindex(Bloomberg_Inventory['P&L'].abs().sort_values(ascending=False).index)
    Bloomberg_Inventory = Bloomberg_Inventory[['Security','CUSIP','P&L','Book','Position','BBG MTG Position','Cumulative Average C']]
    #Bloomberg_Inventory.to_excel( 'C:\\Users\\ccraig\\Desktop\\test1.xlsx')
    

    HT_Files = input_file.HT_FTP_File_Input('Inv Detail',2)
    HT_Files_Recent = HT_Files[0].groupby(['CUSIP','Account Name']).agg({'Quantity':'sum',
                                                                   'Unreal PnL':'sum',
                                                                   'Real PnL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Description':'first',
                                                                   'Price':'mean',
                                                            #       'Account Name':'first',
                                                                   'Req%':'sum'})
    HT_Files_Previous = HT_Files[1].groupby(['CUSIP','Account Name']).agg({'Quantity':'sum',
                                                                   'Unreal PnL':'sum',
                                                                   'Real PnL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Description':'first',
                                                                   'Price':'mean',
                                                             #      'Account Name':'first',
                                                                   'Req%':'sum'})
    HT_Merged = pd.merge(HT_Files_Recent,HT_Files_Previous,on = ['CUSIP','Account Name'],how = 'outer')
    HT_Merged.fillna(0,inplace = True)
    HT_Merged['Requirement Change'] = HT_Merged['Requirement_x']-HT_Merged['Requirement_y']
    HT_Merged.reset_index(inplace = True)
    HT_Merged['CUSIP'] = HT_Merged['CUSIP'].str[:9]
    
    HT_Merged = pd.merge(HT_Merged,Bloomberg_Inventory,on ='CUSIP',how = 'left')
    # change 9/24/2020 - looking from major HT and BBG DSP
    HT_Merged.fillna(0,inplace = True)


    # Adjusting for Mortgage Bond Positions

    HT_Merged['Position'] = np.where((HT_Merged['Account Name'] == 'K76 S P IN'),HT_Merged['BBG MTG Position'],HT_Merged['Position'])
    HT_Merged['Position'] = np.where((HT_Merged['Account Name'] == 'M64 SIERRA'),HT_Merged['BBG MTG Position'],HT_Merged['Position'])

    HT_Merged['QTY DSP'] = HT_Merged['Position'] - HT_Merged['Quantity_x']
    HT_Merged['HT Real PnL Change'] = HT_Merged['Real PnL_x'] - HT_Merged['Real PnL_y']
    HT_Merged['Real PnL DSP'] = HT_Merged['HT Real PnL Change'] - HT_Merged['P&L']

  
    HT_Merged['HT QTY Change'] = HT_Merged['Quantity_x']-HT_Merged['Quantity_y']
    HT_Merged['TW - HT Quantity DSP'] = HT_Merged['Position']-HT_Merged['Quantity_x']
    HT_Merged['Real PnL Change'] = HT_Merged['Real PnL_x']-HT_Merged['Real PnL_y']
    HT_Merged['HT Real PnL Change'] = HT_Merged['Real PnL_x']-HT_Merged['Real PnL_y']
    HT_Merged['Adj Unreal PnL Change'] = HT_Merged['Unreal PnL_x']-HT_Merged['Unreal PnL_y']+HT_Merged['HT Real PnL Change']
    HT_Merged['HT-TW PnL DSP'] = HT_Merged['HT Real PnL Change']-HT_Merged['P&L']
    HT_Merged['Filter Column'] = HT_Merged['TW - HT Quantity DSP'] + HT_Merged['HT QTY Change'] + HT_Merged['Adj Unreal PnL Change'] + HT_Merged['HT-TW PnL DSP'] + HT_Merged['Requirement Change']
    HT_Merged.rename(columns={
        'Price_x':'Price',
        'Position':'BBG QTY',
        'Quantity_x':'HT QTY',
        'Unreal PnL_x': 'HT Current Unreal PnL',
        'Unreal PnL_y':'HT Previous Unreal PnL',
        'P&L':'BBG PnL',
        'Requirement_x':'Current Requirement',
        'Security_x':'Security',
        'Account Name':'Account',
        'Req%_x':'Req%'},inplace = True)
    HT_Merged.loc[HT_Merged['Security']==0,'Security'] = HT_Merged['Description_x']
    HT_Merged.loc[HT_Merged['Security']==0,'Security'] = HT_Merged['Description_y']
    HT_Merged.drop_duplicates(subset=['CUSIP','ACCOUNT'],keep='first')
    
    #HT_Merged.loc[HT_Merged['Account']==0,'Security'] = HT_Merged['Account Name_y']
    HT_Detail = HT_Merged[['Security','CUSIP','Account','Price','BBG QTY','HT QTY','QTY DSP','HT QTY Change','HT Current Unreal PnL','HT Previous Unreal PnL',
                           'Real PnL Change','Adj Unreal PnL Change','BBG PnL','HT-TW PnL DSP','Requirement Change','Current Requirement','Req%']]
    HT_Detail = HT_Detail.reindex(HT_Detail['BBG PnL'].abs().sort_values(ascending=False).index)
   
    
    Adj_Unreal_Change = HT_Merged[['Security','CUSIP','Account','Adj Unreal PnL Change']]
    Adj_Unreal_Change = Adj_Unreal_Change.reindex(Adj_Unreal_Change['Adj Unreal PnL Change'].abs().sort_values(ascending=False).index)

    Requirement_Change = HT_Merged[['Security','CUSIP','Account','Requirement Change']]
    Requirement_Change.drop_duplicates(subset = ['CUSIP'],keep = 'first',inplace = True)
    #HT_Detail.to_excel('C:Users/ccraig/Desktop/test.xlsx')
    Requirement_Change = Requirement_Change.reindex(Requirement_Change['Requirement Change'].abs().sort_values(ascending=False).index)
    return HT_Detail,Adj_Unreal_Change,Requirement_Change,Bloomberg_Inventory
     
