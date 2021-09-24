
import datetime as datetime
File_Path_Text = {
    'Past_Position_File_Path':'P:/2. Corps/PNL_Daily_Report/Reports/',
    'QTY_DSP_Cleared_File_Path':'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx',
    'Running_PnL_DSP_File_Path':'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/PNL_DSP_History.xlsx'
    }
HT_File_Path = 'P:/DEV/HillTop_FTP/inventory-margin-sierra-'
TW_File_Path = 'C:\\Users\\ccraig\\Desktop\\PNL Project\\Bloomberg TW.xls'
BBG_File_Path = 'P:/2. Corps/PNL_Daily_Report/BBG_Files/'
HT_History_File_Path = 'P:/2. Corps/PNL_Daily_Report/HT_Files/'
TOMS_Email_Subjects = {
    'Toms Email 1': 'Report "TW 16 22" is generated.',
    'Toms Email 2': 'Report "TW 1 5" is generated.',
    'Toms Email 3': 'Report "TW 6 10" is generated.',
    'Toms Email 4': 'Report "TW 11 15" is generated.'
    }
Archive_File_Path = 'P:/2. Corps/PNL_Daily_Report/Reports/'

US_Trading_Holiday = [datetime.datetime(2020,4,10).strftime("%Y-%m-%d"),
                      datetime.datetime(2020,5,25).strftime("%Y-%m-%d"),
                      datetime.datetime(2020,7,3).strftime("%Y-%m-%d"),
                      datetime.datetime(2020,9,7).strftime("%Y-%m-%d"),
                      datetime.datetime(2020,10,12).strftime("%Y-%m-%d"),
                      datetime.datetime(2020,11,11).strftime("%Y-%m-%d"),
                      datetime.datetime(2020,11,26).strftime("%Y-%m-%d"),
                      datetime.datetime(2020,12,25).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,1,1).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,1,18).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,2,15).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,4,2).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,5,31).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,7,5).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,9,6).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,10,11).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,11,11).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,11,25).strftime("%Y-%m-%d"),
                      datetime.datetime(2021,12,24).strftime("%Y-%m-%d")]

Muni_Accounts = ['K72','K78','K79','K80','K81','K82','K0P60','K0P61','K0P66','K0P67','Total']
CMO_Accounts = ['K76','M64','CMO Total']
Email_Address_List = 'ccraig@sierrapacificsecurities.com;jblamire@sierrapacificsecurities.com;lankowsky@bloomberg.net;jdean@sierrapacificsecurities.com;bburdick@sierrapacificsecurities.com;tpedersen@sierrapacificsecurities.com'
BCC_Email_Addres_List = 'tcarney@sierrapacificsecurities.com;mleck@sierrapacificsecurities.com;dbrooks@sierrapacificsecurities.com;jcurran@sierrapacificsecurities.com;dgrijalva@sierrapacificsecurities.com;jread@sierrapacificsecurities.com'
#Excel_File_Address = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
Excel_File_Address = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report_1.xlsx'