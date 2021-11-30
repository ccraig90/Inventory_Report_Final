import Module.Input_Files as input_file
import pandas as pd

def Short_Check():
    HT_Files = input_file.HT_FTP_File_Input('Inv Detail',2)
    HT_Files_Recent = HT_Files[0].groupby(['CUSIP']).agg({'Quantity':'sum',
                                                                'Unreal PnL':'sum',
                                                                'Real PnL':'sum',
                                                                'Requirement':'sum',
                                                                'Description':'first',
                                                                'Price':'mean',
                                                                'Account Name':'first'})

    K72_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == 'K72 MUNI I']
    K78_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == 'K78 TAXABL']
    K79_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == 'K79 CALI T']
    K80_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == 'K80 MUNI T']
    K81_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == 'K81 MUNI T']
    K82_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == 'K82 TAX 0']
    K0P60_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == '0P66 MUNI']
    K0P61_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == '0P61 MUNI']
    K0P66_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == '0P66 MUNI']
    K0P67_Muni = HT_Files_Recent.loc[HT_Files_Recent['Account Name'] == '0P67 MUNI']
    

    K72 = (K72_Muni['Quantity'].values < 0).any()
    K78 = (K78_Muni['Quantity'].values < 0).any()
    K79 = (K79_Muni['Quantity'].values < 0).any()
    K80 = (K80_Muni['Quantity'].values < 0).any()
    K81 = (K81_Muni['Quantity'].values < 0).any()
    K82 = (K82_Muni['Quantity'].values < 0).any()
    K0P60 = (K0P60_Muni['Quantity'].values < 0).any()
    K0P61 = (K0P61_Muni['Quantity'].values < 0).any()
    K0P66 = (K0P66_Muni['Quantity'].values < 0).any()
    K0P67 = (K0P67_Muni['Quantity'].values < 0).any()
    
    Muni_Short_Check = {
        'K72':K72,
        'K78':K78,
        'K79':K79,
        'K80':K80,
        'K81':K81,
        'K82':K82,
        'K0P60':K0P60,
        'K0P61':K0P61,
        'K0P66':K0P66,
        'K0P67':K0P67}
    Muni_Short_Check = pd.DataFrame.from_dict(Muni_Short_Check,orient = 'index')
    return Muni_Short_Check


