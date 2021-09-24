import datetime as datetime
import config as config
import Module.Date_File_Path as date_file_path
import Module.Input_Files as input_file
import pandas as pd





def Month_Summary():
    HT_Files = input_file.HT_FTP_File_Input('Inv Acct Summary',2)
    for item in HT_Files:
        item.drop(columns=['Office','RR','Unnamed: 7','Repo Adjustment','Account'],inplace = True)
        item['Unique Identifier'] = item['Account Name'] +' - ' + item['Position Type']
    
    Month_Summary = HT_Files[0]
    Month_Summary = Month_Summary[['Account Name','Position Type','Market Value','Requirement','Unreal PnL','Real PnL','Unique Identifier']]
    # -- Missing Account Short Correction --
    Blank_Dictionary = {'Account Name':[' '],'Position Type': [' '],'Market Value':[0],'Requirement':[0],'Unreal PnL':[0],'Real PnL':[0],'Unique Identifier':['x']}
    Blank_Dictionary_df = pd.DataFrame(data = Blank_Dictionary)

    # -- Corp Section --
    N88_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'N88 CORP   - Long']
    N88_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'N88 CORP   - Short']
    if N88_Short.empty:
        N88_Short = Blank_Dictionary_df
    accounts = [N88_Long,N88_Short]
    N88_Account = pd.concat(accounts,join = 'outer')
    N88_Account.loc['Total'] = N88_Account.sum(numeric_only = True,axis = 0)

    N90_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'N90 CD     - Long']
    N90_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'N90 CD     - Short']
    if N90_Short.empty:
        N90_Short = Blank_Dictionary_df
    accounts = [N90_Long,N90_Short]
    N90_Account = pd.concat(accounts,join = 'outer')
    N90_Account.loc['Total'] = N90_Account.sum(numeric_only = True,axis = 0)

    P01_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'P01 CORP   - Long']
    P01_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'P01 CORP   - Short']
    if P01_Short.empty:
        P01_Short = Blank_Dictionary_df
    accounts = [P01_Long,P01_Short]
    P01_Account = pd.concat(accounts,join = 'outer')
    P01_Account.loc['Total'] = P01_Account.sum(numeric_only = True,axis = 0)

    P02_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'P02 CORP   - Long']
    P02_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'P02 CORP   - Short']
    if P02_Short.empty:
        P02_Short = Blank_Dictionary_df
    accounts = [P02_Long,P02_Short]
    P02_Account = pd.concat(accounts,join = 'outer')
    P02_Account.loc['Total'] = P02_Account.sum(numeric_only = True,axis = 0)

    K74_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K74 CORPOR - Long']
    K74_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K74 CORPOR - Short']
    if K74_Short.empty:
        K74_Short = Blank_Dictionary_df
    accounts = [K74_Long,K74_Short]
    K74_Account = pd.concat(accounts,join = 'outer')
    K74_Account.loc['Total'] = K74_Account.sum(numeric_only = True,axis = 0)

    L81_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'L81 SIERRA - Long']
    L81_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'L81 SIERRA - Short']
    if L81_Short.empty:
        L81_Short = Blank_Dictionary_df
    accounts = [L81_Long,L81_Short]
    L81_Account = pd.concat(accounts,join = 'outer')
    L81_Account.loc['Total'] = L81_Account.sum(numeric_only = True,axis = 0)

    P03_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'P03 NEW C  - Long']
    P03_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'P03 NEW C  - Short']
    if P03_Short.empty:
        P03_Short = Blank_Dictionary_df
    accounts = [P03_Long,P03_Short]
    P03_Account = pd.concat(accounts,join = 'outer')
    P03_Account.loc['Total'] = P03_Account.sum(numeric_only = True,axis = 0)


    N87_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'N87 CORP   - Long']
    N87_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'N87 CORP   - Short']
    if N87_Short.empty:
        N87_Short = Blank_Dictionary_df
    accounts = [N87_Long,N87_Short]
    N87_Account = pd.concat(accounts,join = 'outer')
    N87_Account.loc['Total'] = N87_Account.sum(numeric_only = True,axis = 0)


    N89_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'N89 CORP   - Long']
    N89_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'N89 CORP   - Short']
    if N89_Long.empty:
        N89_Long = Blank_Dictionary_df
    if N89_Short.empty:
        N89_Short = Blank_Dictionary_df
    accounts = [N89_Long,N89_Short]
    N89_Account = pd.concat(accounts,join = 'outer')
    N89_Account.loc['Total'] = N89_Account.sum(numeric_only = True,axis = 0)

    corp_accounts =[N88_Account,
                    N90_Account,
                    P01_Account,
                    P02_Account,
                    K74_Account,
                    L81_Account,
                    P03_Account,
                    N87_Account,
                    N89_Account]
    Corp_Section = pd.concat(corp_accounts,join = 'outer')
    Corp_Section = Corp_Section[['Account Name','Position Type','Market Value','Requirement','Unreal PnL','Real PnL']]

    # -- Muni Section --

    K72_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K72 MUNI I - Long']
    K72_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K72 MUNI I - Short']
    if K72_Short.empty == False: # check to see if short position exists
        K72_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K72_is_short = ''        # '' assigned indicating there was no short present
    if K72_Short.empty:
        K72_Short = Blank_Dictionary_df
    accounts = [K72_Long,K72_Short]
    K72_Account = pd.concat(accounts,join = 'outer')
    

    K78_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K78 TAXABL - Long']
    K78_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K78 TAXABL - Short']
    if K78_Short.empty == False: # check to see if short position exists
        K78_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K78_is_short = ''        # '' assigned indicating there was no short present
    if K78_Short.empty:
        K78_Short = Blank_Dictionary_df
    accounts = [K78_Long,K78_Short]
    K78_Account = pd.concat(accounts,join = 'outer')
    

    K79_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K79 CALI T - Long']
    K79_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K79 CALI T - Short']
    if K79_Short.empty == False: # check to see if short position exists
        K79_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K79_is_short = ''        # '' assigned indicating there was no short present
    if K79_Short.empty:
        K79_Short = Blank_Dictionary_df
    accounts = [K79_Long,K79_Short]
    K79_Account = pd.concat(accounts,join = 'outer')

    K80_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K80 MUNI T - Long']
    K80_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K80 MUNI T - Short']
    if K80_Short.empty == False: # check to see if short position exists
        K80_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K80_is_short = ''        # '' assigned indicating there was no short present
    if K80_Short.empty:
        K80_Short = Blank_Dictionary_df
    accounts = [K80_Long,K80_Short]
    K80_Account = pd.concat(accounts,join = 'outer')

    K81_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K81 MUNI T - Long']
    K81_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K81 MUNI T - Short']
    if K81_Short.empty == False: # check to see if short position exists
        K81_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K81_is_short = ''        # '' assigned indicating there was no short present
    if K81_Short.empty:
        K81_Short = Blank_Dictionary_df
    accounts = [K81_Long,K81_Short]
    K81_Account = pd.concat(accounts,join = 'outer')

    K82_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K82 TAX 0  - Long']
    K82_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K82 TAX 0  - Short']
    if K82_Short.empty == False: # check to see if short position exists
        K82_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K82_is_short = ''          # '' assigned indicating there was no short present
    if K82_Short.empty:
        K81_Short = Blank_Dictionary_df
    accounts = [K82_Long,K82_Short]
    K82_Account = pd.concat(accounts,join = 'outer')

    K0P60_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == '0P60 MUNI  - Long']
    K0P60_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == '0P60 MUNI  - Short']
    if K0P60_Short.empty == False:   # check to see if short position exists
        K0P60_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P60_is_short = ''          # '' assigned indicating there was no short present
    if K0P60_Short.empty:
        K0P60_Short = Blank_Dictionary_df
    accounts = [K0P60_Long,K0P60_Short]
    K0P60_Account = pd.concat(accounts,join = 'outer')

    K0P61_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == '0P61 MUNI  - Long']
    K0P61_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == '0P61 MUNI  - Short']
    if K0P61_Short.empty == False:  # check to see if short position exists
        K0P61_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P61_is_short = ''          # '' assigned indicating there was no short present
    if K0P61_Short.empty:
        K0P61_Short = Blank_Dictionary_df
    accounts = [K0P61_Long,K0P61_Short]
    K0P61_Account = pd.concat(accounts,join = 'outer')

    K0P66_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == '0P66 MUNI  - Long']
    K0P66_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == '0P66 MUNI  - Short']
    if K0P61_Short.empty == False:  # check to see if short position exists
        K0P61_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P66_is_short = ''          # '' assigned indicating there was no short present
    if K0P66_Short.empty:
        K0P66_Short = Blank_Dictionary_df
    accounts = [K0P66_Long,K0P66_Short]
    K0P66_Account = pd.concat(accounts,join = 'outer')

    K0P67_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == '0P67 MUNI  - Long']
    K0P67_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == '0P67 MUNI  - Short']
    if K0P67_Short.empty == False:  # check to see if short position exists
        K0P67_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P67_is_short = ''          # '' assigned indicating there was no short present
    if K0P67_Short.empty:
        K0P67_Short = Blank_Dictionary_df
    accounts = [K0P67_Long,K0P67_Short]
    K0P67_Account = pd.concat(accounts,join = 'outer')

    muni_accounts =[K72_Account,
                    K78_Account,
                    K79_Account,
                    K80_Account,
                    K81_Account,
                    K82_Account,
                    K0P60_Account,
                    K0P61_Account,
                    K0P66_Account,
                    K0P67_Account]

    for item in muni_accounts:
        item.loc['Total'] = item.sum(numeric_only = True, axis = 0)

    Muni_Section = pd.concat(muni_accounts,join = 'outer')
    Muni_Section.reset_index(inplace = True)
    Muni_Section = Muni_Section.loc[Muni_Section['index'] == 'Total']
    Muni_Section.loc['Muni Total'] = Muni_Section.sum(numeric_only = True,axis = 0)
    Muni_Section['Account Name'] = config.Muni_Accounts
    Muni_Section = Muni_Section[['Account Name','Position Type','Market Value','Requirement','Unreal PnL','Real PnL']]

    # -- CMO Section--
 
    K76_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K76 S P IN - Long']
    K76_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K76 S P IN - Short']
    if K76_Short.empty == False:  # check to see if short position exists
        K76_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K76_is_short = ''          # '' assigned indicating there was no short present
    if K76_Short.empty:
        K76_Short = Blank_Dictionary_df
    accounts = [K76_Long,K76_Short]
    K76_Account = pd.concat(accounts,join = 'outer')


    M64_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'M64 SIERRA - Long']
    M64_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'M64 SIERRA - Short']
    if M64_Short.empty == False:  # check to see if short position exists
        M64_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        M64_is_short = ''          # '' assigned indicating there was no short present
    if M64_Short.empty:
        M64_Short = Blank_Dictionary_df
    accounts = [M64_Long,M64_Short]
    M64_Account = pd.concat(accounts,join = 'outer')

    cmo_accounts = [K76_Account,
                    M64_Account]

    for item in cmo_accounts:
        item.loc['Total'] = item.sum(numeric_only = True, axis = 0)
    CMO_Section = pd.concat(cmo_accounts,join = 'outer')
    CMO_Section.reset_index(inplace = True)
    CMO_Section = CMO_Section.loc[CMO_Section['index'] == 'Total']
    CMO_Section.loc['CMO Total'] = CMO_Section.sum(numeric_only = True,axis = 0)
    CMO_Section['Account Name'] = config.CMO_Accounts


    #CMO_Section.loc['CMO Total'] = CMO_Section.sum(numeric_only = True,axis = 0)
    # -- CD Section --

    K77_Long = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K77 CD INV - Long']
    K77_Short = Month_Summary.loc[Month_Summary['Unique Identifier'] == 'K77 CD INV - Short']
    if K77_Short.empty == False:  # check to see if short position exists
        K77_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K77_is_short = ''          # '' assigned indicating there was no short present
    if K77_Short.empty:
        K77_Short = Blank_Dictionary_df

    cd_accounts = [K77_Long,K77_Short]

    CD_Section = pd.concat(cd_accounts,join = 'outer')
    CD_Section.loc['Total'] = CD_Section.sum(numeric_only = True, axis = 0)
    CD_Section.reset_index(inplace = True)
    CD_Section = CD_Section.loc[CD_Section['index'] == 'Total']


    #CD_Section.loc['CD Total'] = CD_Section.sum(numeric_only = True,axis = 0)
    Is_Short = {'K72':K72_is_short,'K78':K78_is_short,'K79':K79_is_short,'K80':K80_is_short,'K81':K81_is_short,'K82':K82_is_short,'K0P60':K0P60_is_short,'K0P61':K0P61_is_short,'K76':K76_is_short,'M64':M64_is_short,'K77':K77_is_short}

    return Muni_Section,Corp_Section,CD_Section,CMO_Section,Is_Short

























def Month_Summary_Daily_Change():
    HT_Files = input_file.HT_FTP_File_Input('Inv Acct Summary',2)
    for item in HT_Files:
        item.drop(columns=['Office','RR','Unnamed: 7','Repo Adjustment','Account'],inplace = True)
        item['Unique Identifier'] = item['Account Name'] +' - ' + item['Position Type']
    New_File = HT_Files[0]
    Old_File = HT_Files[1]
    # -- Missing Account Short Correction --
    Blank_Dictionary = {'Account Name':[' '],'Position Type': [' '],'Market Value':[0],'Requirement':[0],'Unreal PnL':[0],'Real PnL':[0],'Unique Identifier':['x']}
    Blank_Dictionary_df = pd.DataFrame(data = Blank_Dictionary)

        # -- Corp Section --
    N88_Long = New_File.loc[New_File['Unique Identifier'] == 'N88 CORP   - Long']
    N88_Short = New_File.loc[New_File['Unique Identifier'] == 'N88 CORP   - Short']
    if N88_Short.empty:
        N88_Short = Blank_Dictionary_df
    accounts = [N88_Long,N88_Short]
    N88_Account = pd.concat(accounts,join = 'outer')
    N88_Account.loc['Total'] = N88_Account.sum(numeric_only = True,axis = 0)

    N90_Long = New_File.loc[New_File['Unique Identifier'] == 'N90 CD     - Long']
    N90_Short = New_File.loc[New_File['Unique Identifier'] == 'N90 CD     - Short']
    if N90_Short.empty:
        N90_Short = Blank_Dictionary_df
    accounts = [N90_Long,N90_Short]
    N90_Account = pd.concat(accounts,join = 'outer')
    N90_Account.loc['Total'] = N90_Account.sum(numeric_only = True,axis = 0)

    P01_Long = New_File.loc[New_File['Unique Identifier'] == 'P01 CORP   - Long']
    P01_Short = New_File.loc[New_File['Unique Identifier'] == 'P01 CORP   - Short']
    if P01_Short.empty:
        P01_Short = Blank_Dictionary_df
    accounts = [P01_Long,P01_Short]
    P01_Account = pd.concat(accounts,join = 'outer')
    P01_Account.loc['Total'] = P01_Account.sum(numeric_only = True,axis = 0)

    P02_Long = New_File.loc[New_File['Unique Identifier'] == 'P02 CORP   - Long']
    P02_Short = New_File.loc[New_File['Unique Identifier'] == 'P02 CORP   - Short']
    if P02_Short.empty:
        P02_Short = Blank_Dictionary_df
    accounts = [P02_Long,P02_Short]
    P02_Account = pd.concat(accounts,join = 'outer')
    P02_Account.loc['Total'] = P02_Account.sum(numeric_only = True,axis = 0)

    K74_Long = New_File.loc[New_File['Unique Identifier'] == 'K74 CORPOR - Long']
    K74_Short = New_File.loc[New_File['Unique Identifier'] == 'K74 CORPOR - Short']
    if K74_Short.empty:
        K74_Short = Blank_Dictionary_df
    accounts = [K74_Long,K74_Short]
    K74_Account = pd.concat(accounts,join = 'outer')
    K74_Account.loc['Total'] = K74_Account.sum(numeric_only = True,axis = 0)

    L81_Long = New_File.loc[New_File['Unique Identifier'] == 'L81 SIERRA - Long']
    L81_Short = New_File.loc[New_File['Unique Identifier'] == 'L81 SIERRA - Short']
    if L81_Short.empty:
        L81_Short = Blank_Dictionary_df
    accounts = [L81_Long,L81_Short]
    L81_Account = pd.concat(accounts,join = 'outer')
    L81_Account.loc['Total'] = L81_Account.sum(numeric_only = True,axis = 0)

    P03_Long = New_File.loc[New_File['Unique Identifier'] == 'P03 NEW C  - Long']
    P03_Short = New_File.loc[New_File['Unique Identifier'] == 'P03 NEW C  - Short']
    if P03_Short.empty:
        P03_Short = Blank_Dictionary_df
    accounts = [P03_Long,P03_Short]
    P03_Account = pd.concat(accounts,join = 'outer')
    P03_Account.loc['Total'] = P03_Account.sum(numeric_only = True,axis = 0)


    N87_Long = New_File.loc[New_File['Unique Identifier'] == 'N87 CORP   - Long']
    N87_Short = New_File.loc[New_File['Unique Identifier'] == 'N87 CORP   - Short']
    if N87_Short.empty:
        N87_Short = Blank_Dictionary_df
    accounts = [N87_Long,N87_Short]
    N87_Account = pd.concat(accounts,join = 'outer')
    N87_Account.loc['Total'] = N87_Account.sum(numeric_only = True,axis = 0)


    N89_Long = New_File.loc[New_File['Unique Identifier'] == 'N89 CORP   - Long']
    N89_Short = New_File.loc[New_File['Unique Identifier'] == 'N89 CORP   - Short']
    if N89_Long.empty:
        N89_Long = Blank_Dictionary_df
    if N89_Short.empty:
        N89_Short = Blank_Dictionary_df
    accounts = [N89_Long,N89_Short]
    N89_Account = pd.concat(accounts,join = 'outer')
    N89_Account.loc['Total'] = N89_Account.sum(numeric_only = True,axis = 0)

    corp_accounts =[N88_Account,
                    N90_Account,
                    P01_Account,
                    P02_Account,
                    K74_Account,
                    L81_Account,
                    P03_Account,
                    N87_Account,
                    N89_Account]
    Corp_Section = pd.concat(corp_accounts,join = 'outer')
    New_Corp_Section = Corp_Section[['Account Name','Position Type','Market Value','Requirement','Unreal PnL','Real PnL']]
    New_Corp_Section.reset_index(inplace = True)
    New_Corp_Section.reset_index(inplace = True)


    # -- Muni Section --

    K72_Long = New_File.loc[New_File['Unique Identifier'] == 'K72 MUNI I - Long']
    K72_Short = New_File.loc[New_File['Unique Identifier'] == 'K72 MUNI I - Short']
    if K72_Short.empty == False: # check to see if short position exists
        K72_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K72_is_short = ''        # '' assigned indicating there was no short present
    if K72_Short.empty:
        K72_Short = Blank_Dictionary_df
    accounts = [K72_Long,K72_Short]
    K72_Account = pd.concat(accounts,join = 'outer')
    

    K78_Long = New_File.loc[New_File['Unique Identifier'] == 'K78 TAXABL - Long']
    K78_Short = New_File.loc[New_File['Unique Identifier'] == 'K78 TAXABL - Short']
    if K78_Short.empty == False: # check to see if short position exists
        K78_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K78_is_short = ''        # '' assigned indicating there was no short present
    if K78_Short.empty:
        K78_Short = Blank_Dictionary_df
    accounts = [K78_Long,K78_Short]
    K78_Account = pd.concat(accounts,join = 'outer')
    

    K79_Long = New_File.loc[New_File['Unique Identifier'] == 'K79 CALI T - Long']
    K79_Short = New_File.loc[New_File['Unique Identifier'] == 'K79 CALI T - Short']
    if K79_Short.empty == False: # check to see if short position exists
        K79_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K79_is_short = ''        # '' assigned indicating there was no short present
    if K79_Short.empty:
        K79_Short = Blank_Dictionary_df
    accounts = [K79_Long,K79_Short]
    K79_Account = pd.concat(accounts,join = 'outer')

    K80_Long = New_File.loc[New_File['Unique Identifier'] == 'K80 MUNI T - Long']
    K80_Short = New_File.loc[New_File['Unique Identifier'] == 'K80 MUNI T - Short']
    if K80_Short.empty == False: # check to see if short position exists
        K80_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K80_is_short = ''        # '' assigned indicating there was no short present
    if K80_Short.empty:
        K80_Short = Blank_Dictionary_df
    accounts = [K80_Long,K80_Short]
    K80_Account = pd.concat(accounts,join = 'outer')

    K81_Long = New_File.loc[New_File['Unique Identifier'] == 'K81 MUNI T - Long']
    K81_Short = New_File.loc[New_File['Unique Identifier'] == 'K81 MUNI T - Short']
    if K81_Short.empty == False: # check to see if short position exists
        K81_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K81_is_short = ''        # '' assigned indicating there was no short present
    if K81_Short.empty:
        K81_Short = Blank_Dictionary_df
    accounts = [K81_Long,K81_Short]
    K81_Account = pd.concat(accounts,join = 'outer')

    K82_Long = New_File.loc[New_File['Unique Identifier'] == 'K82 TAX 0  - Long']
    K82_Short = New_File.loc[New_File['Unique Identifier'] == 'K82 TAX 0  - Short']
    if K82_Short.empty == False: # check to see if short position exists
        K82_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K82_is_short = ''          # '' assigned indicating there was no short present
    if K82_Short.empty:
        K81_Short = Blank_Dictionary_df
    accounts = [K82_Long,K82_Short]
    K82_Account = pd.concat(accounts,join = 'outer')

    K0P60_Long = New_File.loc[New_File['Unique Identifier'] == '0P60 MUNI  - Long']
    K0P60_Short = New_File.loc[New_File['Unique Identifier'] == '0P60 MUNI  - Short']
    if K0P60_Short.empty == False:   # check to see if short position exists
        K0P60_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P60_is_short = ''          # '' assigned indicating there was no short present
    if K0P60_Short.empty:
        K0P60_Short = Blank_Dictionary_df
    accounts = [K0P60_Long,K0P60_Short]
    K0P60_Account = pd.concat(accounts,join = 'outer')

    K0P61_Long = New_File.loc[New_File['Unique Identifier'] == '0P61 MUNI  - Long']
    K0P61_Short = New_File.loc[New_File['Unique Identifier'] == '0P61 MUNI  - Short']
    if K0P61_Short.empty == False:  # check to see if short position exists
        K0P61_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P61_is_short = ''          # '' assigned indicating there was no short present
    if K0P61_Short.empty:
        K0P61_Short = Blank_Dictionary_df
    accounts = [K0P61_Long,K0P61_Short]
    K0P61_Account = pd.concat(accounts,join = 'outer')

    K0P66_Long =New_File.loc[New_File['Unique Identifier'] == '0P66 MUNI  - Long']
    K0P66_Short = New_File.loc[New_File['Unique Identifier'] == '0P66 MUNI  - Short']
    if K0P61_Short.empty == False:  # check to see if short position exists
        K0P61_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P66_is_short = ''          # '' assigned indicating there was no short present
    if K0P66_Short.empty:
        K0P66_Short = Blank_Dictionary_df
    accounts = [K0P66_Long,K0P66_Short]
    K0P66_Account = pd.concat(accounts,join = 'outer')

    K0P67_Long = New_File.loc[New_File['Unique Identifier'] == '0P67 MUNI  - Long']
    K0P67_Short = New_File.loc[New_File['Unique Identifier'] == '0P67 MUNI  - Short']
    if K0P67_Short.empty == False:  # check to see if short position exists
        K0P67_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P67_is_short = ''          # '' assigned indicating there was no short present
    if K0P67_Short.empty:
        K0P67_Short = Blank_Dictionary_df
    accounts = [K0P61_Long,K0P67_Short]
    K0P67_Account = pd.concat(accounts,join = 'outer')

    muni_accounts =[K72_Account,
                    K78_Account,
                    K79_Account,
                    K80_Account,
                    K81_Account,
                    K82_Account,
                    K0P60_Account,
                    K0P61_Account,
                    K0P66_Account,
                    K0P67_Account]

    for item in muni_accounts:
        item.loc['Total'] = item.sum(numeric_only = True, axis = 0)

    Muni_Section = pd.concat(muni_accounts,join = 'outer')
    Muni_Section.reset_index(inplace = True)
    Muni_Section = Muni_Section.loc[Muni_Section['index'] == 'Total']
    Muni_Section.loc['Muni Total'] = Muni_Section.sum(numeric_only = True,axis = 0)
    print(Muni_Section['Account Name'])
    Muni_Section['Account Name'] = config.Muni_Accounts
    New_Muni_Section = Muni_Section[['Account Name','Position Type','Market Value','Requirement','Unreal PnL','Real PnL']]
    New_Muni_Section.reset_index(inplace=True)
    New_Muni_Section.reset_index(inplace=True)
    # -- CMO Section--
 
    K76_Long = New_File.loc[New_File['Unique Identifier'] == 'K76 S P IN - Long']
    K76_Short = New_File.loc[New_File['Unique Identifier'] == 'K76 S P IN - Short']
    if K76_Short.empty == False:  # check to see if short position exists
        K76_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K76_is_short = ''          # '' assigned indicating there was no short present
    if K76_Short.empty:
        K76_Short = Blank_Dictionary_df
    accounts = [K76_Long,K76_Short]
    K76_Account = pd.concat(accounts,join = 'outer')


    M64_Long = New_File.loc[New_File['Unique Identifier'] == 'M64 SIERRA - Long']
    M64_Short = New_File.loc[New_File['Unique Identifier'] == 'M64 SIERRA - Short']
    if M64_Short.empty == False:  # check to see if short position exists
        M64_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        M64_is_short = ''          # '' assigned indicating there was no short present
    if M64_Short.empty:
        M64_Short = Blank_Dictionary_df
    accounts = [M64_Long,M64_Short]
    M64_Account = pd.concat(accounts,join = 'outer')

    cmo_accounts = [K76_Account,
                    M64_Account]
    
    for item in cmo_accounts:
        item.loc['Total'] = item.sum(numeric_only = True, axis = 0)
    New_CMO_Section = pd.concat(cmo_accounts,join = 'outer')
    New_CMO_Section.reset_index(inplace=True)
    New_CMO_Section = New_CMO_Section.loc[New_CMO_Section['index'] == 'Total']
    New_CMO_Section.loc['CMO Total'] = New_CMO_Section.sum(numeric_only = True,axis = 0)
    New_CMO_Section['Account Name'] = config.CMO_Accounts



    #CMO_Section.loc['CMO Total'] = CMO_Section.sum(numeric_only = True,axis = 0)
    # -- CD Section --

    K77_Long = New_File.loc[New_File['Unique Identifier'] == 'K77 CD INV - Long']
    K77_Short = New_File.loc[New_File['Unique Identifier'] == 'K77 CD INV - Short']
    if K77_Short.empty == False:  # check to see if short position exists
        K77_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K77_is_short = ''          # '' assigned indicating there was no short present
    if K77_Short.empty:
        K77_Short = Blank_Dictionary_df

    cd_accounts = [K77_Long,K77_Short]

    New_CD_Section = pd.concat(cd_accounts,join = 'outer')
    New_CD_Section.loc['Total'] = New_CD_Section.sum(numeric_only = True, axis = 0)
    New_CD_Section.reset_index(inplace = True)
    New_CD_Section = New_CD_Section.loc[New_CD_Section['index'] == 'Total']
    New_CD_Section['Account Name'] = 'CD'

    #CD_Section.loc['CD Total'] = CD_Section.sum(numeric_only = True,axis = 0)
    New_Is_Short = {'K72':K72_is_short,'K78':K78_is_short,'K79':K79_is_short,'K80':K80_is_short,'K81':K81_is_short,'K82':K82_is_short,'K0P60':K0P60_is_short,'K0P61':K0P61_is_short,'K76':K76_is_short,'M64':M64_is_short,'K77':K77_is_short}






    """
    Old Section Parse
    """


    # -- Corp Section --
    N88_Long = Old_File.loc[Old_File['Unique Identifier'] == 'N88 CORP   - Long']
    N88_Short = Old_File.loc[Old_File['Unique Identifier'] == 'N88 CORP   - Short']
    if N88_Short.empty:
        N88_Short = Blank_Dictionary_df
    accounts = [N88_Long,N88_Short]
    N88_Account = pd.concat(accounts,join = 'outer')
    N88_Account.loc['Total'] = N88_Account.sum(numeric_only = True,axis = 0)

    N90_Long = Old_File.loc[Old_File['Unique Identifier'] == 'N90 CD     - Long']
    N90_Short = Old_File.loc[Old_File['Unique Identifier'] == 'N90 CD     - Short']
    if N90_Short.empty:
        N90_Short = Blank_Dictionary_df
    accounts = [N90_Long,N90_Short]
    N90_Account = pd.concat(accounts,join = 'outer')
    N90_Account.loc['Total'] = N90_Account.sum(numeric_only = True,axis = 0)

    P01_Long = Old_File.loc[Old_File['Unique Identifier'] == 'P01 CORP   - Long']
    P01_Short = Old_File.loc[Old_File['Unique Identifier'] == 'P01 CORP   - Short']
    if P01_Short.empty:
        P01_Short = Blank_Dictionary_df
    accounts = [P01_Long,P01_Short]
    P01_Account = pd.concat(accounts,join = 'outer')
    P01_Account.loc['Total'] = P01_Account.sum(numeric_only = True,axis = 0)

    P02_Long = Old_File.loc[Old_File['Unique Identifier'] == 'P02 CORP   - Long']
    P02_Short = Old_File.loc[Old_File['Unique Identifier'] == 'P02 CORP   - Short']
    if P02_Short.empty:
        P02_Short = Blank_Dictionary_df
    accounts = [P02_Long,P02_Short]
    P02_Account = pd.concat(accounts,join = 'outer')
    P02_Account.loc['Total'] = P02_Account.sum(numeric_only = True,axis = 0)

    K74_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K74 CORPOR - Long']
    K74_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K74 CORPOR - Short']
    if K74_Short.empty:
        K74_Short = Blank_Dictionary_df
    accounts = [K74_Long,K74_Short]
    K74_Account = pd.concat(accounts,join = 'outer')
    K74_Account.loc['Total'] = K74_Account.sum(numeric_only = True,axis = 0)

    L81_Long = Old_File.loc[Old_File['Unique Identifier'] == 'L81 SIERRA - Long']
    L81_Short = Old_File.loc[Old_File['Unique Identifier'] == 'L81 SIERRA - Short']
    if L81_Short.empty:
        L81_Short = Blank_Dictionary_df
    accounts = [L81_Long,L81_Short]
    L81_Account = pd.concat(accounts,join = 'outer')
    L81_Account.loc['Total'] = L81_Account.sum(numeric_only = True,axis = 0)

    P03_Long = Old_File.loc[Old_File['Unique Identifier'] == 'P03 NEW C  - Long']
    P03_Short = Old_File.loc[Old_File['Unique Identifier'] == 'P03 NEW C  - Short']
    if P03_Short.empty:
        P03_Short = Blank_Dictionary_df
    accounts = [P03_Long,P03_Short]
    P03_Account = pd.concat(accounts,join = 'outer')
    P03_Account.loc['Total'] = P03_Account.sum(numeric_only = True,axis = 0)


    N87_Long = Old_File.loc[Old_File['Unique Identifier'] == 'N87 CORP   - Long']
    N87_Short = Old_File.loc[Old_File['Unique Identifier'] == 'N87 CORP   - Short']
    if N87_Short.empty:
        N87_Short = Blank_Dictionary_df
    accounts = [N87_Long,N87_Short]
    N87_Account = pd.concat(accounts,join = 'outer')
    N87_Account.loc['Total'] = N87_Account.sum(numeric_only = True,axis = 0)


    N89_Long = Old_File.loc[Old_File['Unique Identifier'] == 'N89 CORP   - Long']
    N89_Short = Old_File.loc[Old_File['Unique Identifier'] == 'N89 CORP   - Short']
    if N89_Long.empty:
        N89_Long = Blank_Dictionary_df
    if N89_Short.empty:
        N89_Short = Blank_Dictionary_df
    accounts = [N89_Long,N89_Short]
    N89_Account = pd.concat(accounts,join = 'outer')
    N89_Account.loc['Total'] = N89_Account.sum(numeric_only = True,axis = 0)

    corp_accounts =[N88_Account,
                    N90_Account,
                    P01_Account,
                    P02_Account,
                    K74_Account,
                    L81_Account,
                    P03_Account,
                    N87_Account,
                    N89_Account]
    Corp_Section = pd.concat(corp_accounts,join = 'outer')
    Old_Corp_Section = Corp_Section[['Account Name','Position Type','Market Value','Requirement','Unreal PnL','Real PnL']]
    Old_Corp_Section.reset_index(inplace = True)
    Old_Corp_Section.reset_index(inplace = True)

    # -- Muni Section --

    K72_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K72 MUNI I - Long']
    K72_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K72 MUNI I - Short']
    if K72_Short.empty == False: # check to see if short position exists
        K72_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K72_is_short = ''        # '' assigned indicating there was no short present
    if K72_Short.empty:
        K72_Short = Blank_Dictionary_df
    accounts = [K72_Long,K72_Short]
    K72_Account = pd.concat(accounts,join = 'outer')
    

    K78_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K78 TAXABL - Long']
    K78_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K78 TAXABL - Short']
    if K78_Short.empty == False: # check to see if short position exists
        K78_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K78_is_short = ''        # '' assigned indicating there was no short present
    if K78_Short.empty:
        K78_Short = Blank_Dictionary_df
    accounts = [K78_Long,K78_Short]
    K78_Account = pd.concat(accounts,join = 'outer')
    

    K79_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K79 CALI T - Long']
    K79_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K79 CALI T - Short']
    if K79_Short.empty == False: # check to see if short position exists
        K79_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K79_is_short = ''        # '' assigned indicating there was no short present
    if K79_Short.empty:
        K79_Short = Blank_Dictionary_df
    accounts = [K79_Long,K79_Short]
    K79_Account = pd.concat(accounts,join = 'outer')

    K80_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K80 MUNI T - Long']
    K80_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K80 MUNI T - Short']
    if K80_Short.empty == False: # check to see if short position exists
        K80_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K80_is_short = ''        # '' assigned indicating there was no short present
    if K80_Short.empty:
        K80_Short = Blank_Dictionary_df
    accounts = [K80_Long,K80_Short]
    K80_Account = pd.concat(accounts,join = 'outer')

    K81_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K81 MUNI T - Long']
    K81_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K81 MUNI T - Short']
    if K81_Short.empty == False: # check to see if short position exists
        K81_is_short = 'Short'   # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K81_is_short = ''        # '' assigned indicating there was no short present
    if K81_Short.empty:
        K81_Short = Blank_Dictionary_df
    accounts = [K81_Long,K81_Short]
    K81_Account = pd.concat(accounts,join = 'outer')

    K82_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K82 TAX 0  - Long']
    K82_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K82 TAX 0  - Short']
    if K82_Short.empty == False: # check to see if short position exists
        K82_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K82_is_short = ''          # '' assigned indicating there was no short present
    if K82_Short.empty:
        K81_Short = Blank_Dictionary_df
    accounts = [K82_Long,K82_Short]
    K82_Account = pd.concat(accounts,join = 'outer')

    K0P60_Long = Old_File.loc[Old_File['Unique Identifier'] == '0P60 MUNI  - Long']
    K0P60_Short = Old_File.loc[Old_File['Unique Identifier'] == '0P60 MUNI  - Short']
    if K0P60_Short.empty == False:   # check to see if short position exists
        K0P60_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P60_is_short = ''          # '' assigned indicating there was no short present
    if K0P60_Short.empty:
        K0P60_Short = Blank_Dictionary_df
    accounts = [K0P60_Long,K0P60_Short]
    K0P60_Account = pd.concat(accounts,join = 'outer')

    K0P61_Long = Old_File.loc[Old_File['Unique Identifier'] == '0P61 MUNI  - Long']
    K0P61_Short = Old_File.loc[Old_File['Unique Identifier'] == '0P61 MUNI  - Short']
    if K0P61_Short.empty == False:  # check to see if short position exists
        K0P61_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P61_is_short = ''          # '' assigned indicating there was no short present
    if K0P61_Short.empty:
        K0P61_Short = Blank_Dictionary_df
    accounts = [K0P61_Long,K0P61_Short]
    K0P61_Account = pd.concat(accounts,join = 'outer')

    K0P66_Long = Old_File.loc[Old_File['Unique Identifier'] == '0P66 MUNI  - Long']
    K0P66_Short = Old_File.loc[Old_File['Unique Identifier'] == '0P66 MUNI  - Short']
    if K0P61_Short.empty == False:  # check to see if short position exists
        K0P61_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P66_is_short = ''          # '' assigned indicating there was no short present
    if K0P66_Short.empty:
        K0P66_Short = Blank_Dictionary_df
    accounts = [K0P66_Long,K0P66_Short]
    K0P66_Account = pd.concat(accounts,join = 'outer')

    K0P67_Long = Old_File.loc[Old_File['Unique Identifier'] == '0P67 MUNI  - Long']
    K0P67_Short = Old_File.loc[Old_File['Unique Identifier'] == '0P67 MUNI  - Short']
    if K0P67_Short.empty == False:  # check to see if short position exists
        K0P67_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K0P67_is_short = ''          # '' assigned indicating there was no short present
    if K0P67_Short.empty:
        K0P67_Short = Blank_Dictionary_df
    accounts = [K0P61_Long,K0P67_Short]
    K0P67_Account = pd.concat(accounts,join = 'outer')

    muni_accounts =[K72_Account,
                    K78_Account,
                    K79_Account,
                    K80_Account,
                    K81_Account,
                    K82_Account,
                    K0P60_Account,
                    K0P61_Account,
                    K0P66_Account,
                    K0P67_Account]

    for item in muni_accounts:
        item.loc['Total'] = item.sum(numeric_only = True, axis = 0)

    Muni_Section = pd.concat(muni_accounts,join = 'outer')
    Muni_Section.reset_index(inplace = True)
    Muni_Section = Muni_Section.loc[Muni_Section['index'] == 'Total']
    Muni_Section.loc['Muni Total'] = Muni_Section.sum(numeric_only = True,axis = 0)
    Muni_Section['Account Name'] = config.Muni_Accounts
    Old_Muni_Section = Muni_Section[['Account Name','Position Type','Market Value','Requirement','Unreal PnL','Real PnL']]
    Old_Muni_Section.reset_index(inplace = True)
    Old_Muni_Section.reset_index(inplace = True)
    # -- CMO Section--
 
    K76_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K76 S P IN - Long']
    K76_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K76 S P IN - Short']
    if K76_Short.empty == False:  # check to see if short position exists
        K76_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K76_is_short = ''          # '' assigned indicating there was no short present
    if K76_Short.empty:
        K76_Short = Blank_Dictionary_df
    accounts = [K76_Long,K76_Short]
    K76_Account = pd.concat(accounts,join = 'outer')


    M64_Long = Old_File.loc[Old_File['Unique Identifier'] == 'M64 SIERRA - Long']
    M64_Short = Old_File.loc[Old_File['Unique Identifier'] == 'M64 SIERRA - Short']
    if M64_Short.empty == False:  # check to see if short position exists
        M64_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        M64_is_short = ''          # '' assigned indicating there was no short present
    if M64_Short.empty:
        M64_Short = Blank_Dictionary_df
    accounts = [M64_Long,M64_Short]
    M64_Account = pd.concat(accounts,join = 'outer')

    cmo_accounts = [K76_Account,
                    M64_Account]
    for item in cmo_accounts:
        item.loc['Total'] = item.sum(numeric_only = True, axis = 0)
    Old_CMO_Section = pd.concat(cmo_accounts,join = 'outer')
    Old_CMO_Section.reset_index(inplace=True)
    Old_CMO_Section = Old_CMO_Section.loc[Old_CMO_Section['index'] == 'Total']
    Old_CMO_Section.loc['CMO Total'] = Old_CMO_Section.sum(numeric_only = True,axis = 0)
    Old_CMO_Section['Account Name'] = config.CMO_Accounts


    # -- CD Section --

    K77_Long = Old_File.loc[Old_File['Unique Identifier'] == 'K77 CD INV - Long']
    K77_Short = Old_File.loc[Old_File['Unique Identifier'] == 'K77 CD INV - Short']
    if K77_Short.empty == False:  # check to see if short position exists
        K77_is_short = 'Short'     # assigns 'short' to identifier to be used to show if an account has a short position contained
    else:
        K77_is_short = ''          # '' assigned indicating there was no short present
    if K77_Short.empty:
        K77_Short = Blank_Dictionary_df

    cd_accounts = [K77_Long,K77_Short]

    Old_CD_Section = pd.concat(cd_accounts,join = 'outer')
    Old_CD_Section.loc['Total'] = Old_CD_Section.sum(numeric_only = True, axis = 0)
    Old_CD_Section.reset_index(inplace = True)
    Old_CD_Section = Old_CD_Section.loc[Old_CD_Section['index'] == 'Total']
    Old_CD_Section['Account Name'] = 'CD'
    #CD_Section.loc['CD Total'] = CD_Section.sum(numeric_only = True,axis = 0)
    
    Old_Is_Short = {'K72':K72_is_short,'K78':K78_is_short,'K79':K79_is_short,'K80':K80_is_short,'K81':K81_is_short,'K82':K82_is_short,'K0P60':K0P60_is_short,'K0P61':K0P61_is_short}
    Old_Is_Short = pd.DataFrame.from_dict(Old_Is_Short,orient = 'index')


    Corp_Section_Daily_Change = pd.merge(New_Corp_Section,Old_Corp_Section,how = 'inner',on = 'level_0')
    Corp_Section_Daily_Change['Market Value'] = Corp_Section_Daily_Change['Market Value_x'] - Corp_Section_Daily_Change['Market Value_y']
    Corp_Section_Daily_Change['Requirement'] = Corp_Section_Daily_Change['Requirement_x'] - Corp_Section_Daily_Change['Requirement_y']  
    Corp_Section_Daily_Change['Unreal PnL'] = Corp_Section_Daily_Change['Unreal PnL_x'] - Corp_Section_Daily_Change['Unreal PnL_y']
    Corp_Section_Daily_Change['Real PnL'] = Corp_Section_Daily_Change['Real PnL_x'] - Corp_Section_Daily_Change['Real PnL_y']
    Corp_Section_Daily_Change = Corp_Section_Daily_Change[['Account Name_x','Position Type_x','Market Value','Requirement','Unreal PnL','Real PnL']]
    Corp_Section_Daily_Change.rename(columns={'Account Name_x':'Account Name','Position Type_x':'Position Type'},inplace = True)

    Muni_Section_Daily_Change = pd.merge(New_Muni_Section,Old_Muni_Section,how = 'inner',on = 'level_0')
    Muni_Section_Daily_Change['Market Value'] = Muni_Section_Daily_Change['Market Value_x'] - Muni_Section_Daily_Change['Market Value_y']
    Muni_Section_Daily_Change['Requirement'] = Muni_Section_Daily_Change['Requirement_x'] - Muni_Section_Daily_Change['Requirement_y']  
    Muni_Section_Daily_Change['Unreal PnL'] = Muni_Section_Daily_Change['Unreal PnL_x'] - Muni_Section_Daily_Change['Unreal PnL_y']
    Muni_Section_Daily_Change['Real PnL'] = Muni_Section_Daily_Change['Real PnL_x'] - Muni_Section_Daily_Change['Real PnL_y']
    Muni_Section_Daily_Change = Muni_Section_Daily_Change[['Account Name_x','Position Type_x','Market Value','Requirement','Unreal PnL','Real PnL']]
    Muni_Section_Daily_Change.rename(columns={'Account Name_x':'Account Name','Position Type_x':'Position Type'},inplace = True)

    CMO_Section_Daily_Change = pd.merge(New_CMO_Section,Old_CMO_Section,how = 'inner',on = 'Account Name')
    CMO_Section_Daily_Change['Market Value'] = CMO_Section_Daily_Change['Market Value_x'] - CMO_Section_Daily_Change['Market Value_y']
    CMO_Section_Daily_Change['Requirement'] = CMO_Section_Daily_Change['Requirement_x'] - CMO_Section_Daily_Change['Requirement_y']  
    CMO_Section_Daily_Change['Unreal PnL'] = CMO_Section_Daily_Change['Unreal PnL_x'] - CMO_Section_Daily_Change['Unreal PnL_y']
    CMO_Section_Daily_Change['Real PnL'] = CMO_Section_Daily_Change['Real PnL_x'] - CMO_Section_Daily_Change['Real PnL_y']
    CMO_Section_Daily_Change = CMO_Section_Daily_Change[['Account Name','Position Type_x','Market Value','Requirement','Unreal PnL','Real PnL']]
    CMO_Section_Daily_Change.rename(columns={'Position Type_x':'Position Type'},inplace = True)

    CD_Section_Daily_Change = pd.merge(New_CD_Section,Old_CD_Section,how = 'inner',on = 'Account Name')
    CD_Section_Daily_Change['Market Value'] = CD_Section_Daily_Change['Market Value_x'] - CD_Section_Daily_Change['Market Value_y']
    CD_Section_Daily_Change['Requirement'] = CD_Section_Daily_Change['Requirement_x'] - CD_Section_Daily_Change['Requirement_y']  
    CD_Section_Daily_Change['Unreal PnL'] = CD_Section_Daily_Change['Unreal PnL_x'] - CD_Section_Daily_Change['Unreal PnL_y']
    CD_Section_Daily_Change['Real PnL'] = CD_Section_Daily_Change['Real PnL_x'] - CD_Section_Daily_Change['Real PnL_y']
    CD_Section_Daily_Change = CD_Section_Daily_Change[['Account Name','Position Type_x','Market Value','Requirement','Unreal PnL','Real PnL']]
    CD_Section_Daily_Change.rename(columns={'Position Type_x':'Position Type'},inplace = True)

    return Muni_Section_Daily_Change,Corp_Section_Daily_Change,CD_Section_Daily_Change,CMO_Section_Daily_Change,Old_Is_Short