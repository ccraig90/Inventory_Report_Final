
import datetime as datetime
import pandas as pd
import config as config
import Module.Date_File_Path as date_file_path
import Module.Input_Files as input_file
import Module.Summary as summary
import Module.Month_Summary as month_summary
import Module.Real_PnL_DSP as real_pnl_dsp
import Module.Position_DSP as position_dsp
import Module.HT_Detail as ht_detail
import win32com.client
import xlsxwriter
import Module.Muni_Short_Check as muni_short_check


def Excel_Summary():
    Book_Month_Summary_Daily_Change = month_summary.Month_Summary_Daily_Change()
    Book_Month_Summary_Static = month_summary.Month_Summary()
    Firm_Summary = summary.Hilltop_Trade_Acct_Summary()
    Muni_Short_Check = muni_short_check.Short_Check()

    Is_Short = Book_Month_Summary_Daily_Change[4]
    writer = pd.ExcelWriter(config.Excel_File_Address, engine='xlsxwriter')
    Firm_Summary.to_excel(writer,sheet_name = 'Summary',index = True, startcol = 38,startrow = 1)
    Book_Month_Summary_Static[0].to_excel(writer,sheet_name = 'Summary', index = False,startcol=38, startrow = 11)
    Book_Month_Summary_Daily_Change[0].to_excel(writer,sheet_name = 'Summary', index = False,startcol = 48, startrow = 11)
    Book_Month_Summary_Static[1].to_excel(writer,sheet_name = 'Summary', index = False, startcol = 38,startrow = 23,header = False)
    Book_Month_Summary_Daily_Change[1].to_excel(writer,sheet_name = 'Summary', index = False,startcol = 48, startrow = 23,header = False)
    Book_Month_Summary_Static[2].to_excel(writer,sheet_name = 'Summary', index = False,startcol = 38, startrow = 49,header = False)
    Book_Month_Summary_Daily_Change[2].to_excel(writer,sheet_name = 'Summary', index = False,startcol = 48, startrow = 49,header = False)
    Book_Month_Summary_Static[3].to_excel(writer,sheet_name = 'Summary', index = False,startcol = 38, startrow = 55,header = False)
    Book_Month_Summary_Daily_Change[3].to_excel(writer,sheet_name = 'Summary', index = False,startcol = 48, startrow = 55,header = False)
    Muni_Short_Check.to_excel(writer,sheet_name='Summary',index = False, startcol = 50, startrow=100,header = False)
    for item in Is_Short:
        print(Is_Short[item])
    Is_Short.to_excel(writer,sheet_name = 'Summary', index = False,startcol = 49,startrow = 11)

    """
    XLSX Summary Formatting
    """

    workbook = writer.book
    worksheet_summary = writer.sheets['Summary']

    format_top_summary = workbook.add_format({'bg_color':'#000e6b',
                                              'font_size':'10',
                                              'font_color':'white'})
    format_mini_total = workbook.add_format({'num_format': '#,##0',
                                             'font_size':'8',
                                             'bold': True,
                                             'top':1})
    format_general = workbook.add_format({'num_format': '#,##0',
                                             'font_size':'8'})
    format_general_decimal = workbook.add_format({'num_format': '#,##0.00',
                                            'font_size':'8'})
    format_blank_blue = workbook.add_format({'bg_color':'#4267b8',
                                             'font_size':'8',
                                             'font_color':'white'})
    format_top_summary = workbook.add_format({'bg_color':'#000e6b',
                                              'font_size':'10',
                                              'font_color':'white'})
    format_top_summary_textwrap = workbook.add_format({'bg_color':'#000e6b',
                                            'font_size':'10',
                                            'font_color':'white',
                                            'text_wrap':True})

    format_grey_columnhead = workbook.add_format({'bg_color':'#d4d4d4',
                                                  'font_size':'8',
                                                  'bottom':1})

    format_subtotal = workbook.add_format({'num_format': '#,##0',
                                           'bold':True,
                                           'font_size':'10',
                                           'bottom':2,
                                           'top':1})
    format_general_row = workbook.add_format({'font_size':'8',
                                              'num_format': '#,##0'})
    format_general_row_green = workbook.add_format({'font_size':'8',
                                                    'num_format': '#,##0',
                                                    'font_color':'green'})
    format_general_row_red = workbook.add_format({'font_size':'8',
                                                  'num_format': '#,##0',
                                                  'font_color':'red'})
    format_subtotal_row = workbook.add_format({'font_size':'10',
                                              'num_format': '#,##0',
                                              'bottom':1,
                                              'top':1})
    format_group_total = workbook.add_format({'font_size':'10',
                                              'num_format':'#,##0',
                                              'bold': True})
    format_column = workbook.add_format({'bottom':0,
                                         'top':0,
                                         'border_color':'white'})
    format_url_links = workbook.add_format({'font_size':'10',
                                           'font_color':'blue',
                                           'underline': 1})

    worksheet_summary.conditional_format('E3:E9', {'type':'no_errors',
                                            'format':   format_general_row})

    merge_format = workbook.add_format({
        'font_size':'10',
        'bold': 1,
        'border': 0,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color':'#000e6b',
        'font_color':'white'})


    """Format Top Summary"""
    worksheet_summary.write('D1', 'Summary',format_top_summary) 
    worksheet_summary.write('E1', ' ',format_top_summary)
    worksheet_summary.write('F1', ' ',format_top_summary)
    worksheet_summary.write('D2', 'Item',format_grey_columnhead)
    worksheet_summary.write('E2', 'Available Funds',format_grey_columnhead)
    worksheet_summary.write('F2', 'Change',format_grey_columnhead)
    worksheet_summary.write('D3','Available Funds',format_general)
    worksheet_summary.write('D4','Cost',format_general)
    worksheet_summary.write('D5','Market Value',format_mini_total)
    worksheet_summary.write('D6','Unreal Pnl',format_general)
    worksheet_summary.write('D7','Requirement',format_general)
    worksheet_summary.write('D8','Real PnL',format_general)
    worksheet_summary.write('D9','Call/Excess',format_mini_total)
    worksheet_summary.write_formula('E3','=AN3',format_general)
    worksheet_summary.write_formula('E4','=AN4',format_general)
    worksheet_summary.write_formula('E5','=AN5',format_mini_total)
    worksheet_summary.write_formula('E6','=AN6',format_general)
    worksheet_summary.write_formula('E7','=AN7',format_general)
    worksheet_summary.write_formula('E8','=AN8',format_general)
    worksheet_summary.write_formula('E9','=AN9',format_mini_total)
    worksheet_summary.write_formula('F3','=AO3',format_general)
    worksheet_summary.write_formula('F4','=AO4',format_general)
    worksheet_summary.write_formula('F5','=AO5',format_mini_total)
    worksheet_summary.write_formula('F6','=AO6',format_general)
    worksheet_summary.write_formula('F7','=AO7',format_general)
    worksheet_summary.write_formula('F8','=AO8',format_general)
    worksheet_summary.write_formula('F9','=AO9',format_mini_total)



    """Format Month/Daily Change Summary"""
    worksheet_summary.merge_range('A11:F11', 'Month Summary',merge_format)
    worksheet_summary.merge_range('I11:N11', 'Daily Change',merge_format)
    worksheet_summary.write('A12', 'Account Name',format_grey_columnhead)
    worksheet_summary.write('B12', 'Position Type',format_grey_columnhead)
    worksheet_summary.write('C12', 'Market Value',format_grey_columnhead)
    worksheet_summary.write('D12', 'Requirement',format_grey_columnhead)
    worksheet_summary.write('E12', 'Unreal PNL',format_grey_columnhead)
    worksheet_summary.write('F12', 'Real PNL',format_grey_columnhead)
    worksheet_summary.write('I12', 'Account Name',format_grey_columnhead)
    worksheet_summary.write('J12', 'Position Type',format_grey_columnhead)
    worksheet_summary.write('K12', 'Market Value',format_grey_columnhead)
    worksheet_summary.write('L12', 'Requirement',format_grey_columnhead)
    worksheet_summary.write('M12', 'UnrealPNL',format_grey_columnhead)
    worksheet_summary.write('N12', 'Real PNL',format_grey_columnhead)

    worksheet_summary.write_formula('A13','=AM13',format_general)
    worksheet_summary.write_formula('A14','=AM14',format_general)
    worksheet_summary.write_formula('A15','=AM15',format_general)
    worksheet_summary.write_formula('A16','=AM16',format_general)
    worksheet_summary.write_formula('A17','=AM17',format_general)
    worksheet_summary.write_formula('A18','=AM18',format_general)
    worksheet_summary.write_formula('A19','=AM19',format_general)
    worksheet_summary.write_formula('A20','=AM20',format_general)
    worksheet_summary.write_formula('A21','=AM21',format_general)
    worksheet_summary.write_formula('A22','=AM22',format_general)
    worksheet_summary.write('A23','Muni Total',format_subtotal)


    worksheet_summary.write_formula('B23','=if(AN23="","","")',format_subtotal)

    worksheet_summary.write_formula('C13','=AO13',format_general)
    worksheet_summary.write_formula('C14','=AO14',format_general)
    worksheet_summary.write_formula('C15','=AO15',format_general)
    worksheet_summary.write_formula('C16','=AO16',format_general)
    worksheet_summary.write_formula('C17','=AO17',format_general)
    worksheet_summary.write_formula('C18','=AO18',format_general)
    worksheet_summary.write_formula('C19','=AO19',format_general)
    worksheet_summary.write_formula('C20','=AO20',format_general)
    worksheet_summary.write_formula('C21','=AO21',format_general)
    worksheet_summary.write_formula('C22','=AO22',format_general)
    worksheet_summary.write_formula('C23','=AO23',format_subtotal)

    worksheet_summary.write_formula('D13','=AP13',format_general)
    worksheet_summary.write_formula('D14','=AP14',format_general)
    worksheet_summary.write_formula('D15','=AP15',format_general)
    worksheet_summary.write_formula('D16','=AP16',format_general)
    worksheet_summary.write_formula('D17','=AP17',format_general)
    worksheet_summary.write_formula('D18','=AP18',format_general)
    worksheet_summary.write_formula('D19','=AP19',format_general)
    worksheet_summary.write_formula('D20','=AP20',format_general)
    worksheet_summary.write_formula('D21','=AP21',format_general)
    worksheet_summary.write_formula('D22','=AP22',format_general)
    worksheet_summary.write_formula('D23','=AP23',format_subtotal)

    worksheet_summary.write_formula('E13','=AQ13',format_general)
    worksheet_summary.write_formula('E14','=AQ14',format_general)
    worksheet_summary.write_formula('E15','=AQ15',format_general)
    worksheet_summary.write_formula('E16','=AQ16',format_general)
    worksheet_summary.write_formula('E17','=AQ17',format_general)
    worksheet_summary.write_formula('E18','=AQ18',format_general)
    worksheet_summary.write_formula('E19','=AQ19',format_general)
    worksheet_summary.write_formula('E20','=AQ20',format_general)
    worksheet_summary.write_formula('E21','=AQ21',format_general)
    worksheet_summary.write_formula('E22','=AQ22',format_general)
    worksheet_summary.write_formula('E23','=AQ23',format_subtotal)

    worksheet_summary.write_formula('F13','=AR13',format_general)
    worksheet_summary.write_formula('F14','=AR14',format_general)
    worksheet_summary.write_formula('F15','=AR15',format_general)
    worksheet_summary.write_formula('F16','=AR16',format_general)
    worksheet_summary.write_formula('F17','=AR17',format_general)
    worksheet_summary.write_formula('F18','=AR18',format_general)
    worksheet_summary.write_formula('F19','=AR19',format_general)
    worksheet_summary.write_formula('F20','=AR20',format_general)
    worksheet_summary.write_formula('F21','=AR21',format_general)
    worksheet_summary.write_formula('F22','=AR22',format_general)
    worksheet_summary.write_formula('F23','=AR23',format_subtotal)



    worksheet_summary.write_formula('I13','=AW13',format_general)
    worksheet_summary.write_formula('I14','=AW14',format_general)
    worksheet_summary.write_formula('I15','=AW15',format_general) 
    worksheet_summary.write_formula('I16','=AW16',format_general)
    worksheet_summary.write_formula('I17','=AW17',format_general)
    worksheet_summary.write_formula('I18','=AW18',format_general)
    worksheet_summary.write_formula('I19','=AW19',format_general)
    worksheet_summary.write_formula('I20','=AW20',format_general)
    worksheet_summary.write_formula('I21','=AW21',format_general)
    worksheet_summary.write_formula('I22','=AW22',format_general)
    worksheet_summary.write('I23','Muni Total',format_subtotal)

    worksheet_summary.write_formula('J13','=if(AX13="","",AX13)',format_general)
    worksheet_summary.write_formula('J14','=if(AX14="","",AX14)',format_general)
    worksheet_summary.write_formula('J15','=if(AX15="","",AX15)',format_general)
    worksheet_summary.write_formula('J16','=if(AX16="","",AX16)',format_general)
    worksheet_summary.write_formula('J17','=if(AX17="","",AX17)',format_general)
    worksheet_summary.write_formula('J18','=if(AX18="","",AX18)',format_general)
    worksheet_summary.write_formula('J19','=if(AX19="","",AX19)',format_general)
    worksheet_summary.write_formula('J20','=if(AX20="","",AX20)',format_general)
    worksheet_summary.write_formula('J21','=if(AX21="","",AX20)',format_general)
    worksheet_summary.write_formula('J22','=if(AX22="","",AX20)',format_general)
    worksheet_summary.write_formula('J23','=if(AX23="","","")',format_subtotal)

    worksheet_summary.write_formula('K13','=AY13',format_general)
    worksheet_summary.write_formula('K14','=AY14',format_general)
    worksheet_summary.write_formula('K15','=AY15',format_general)
    worksheet_summary.write_formula('K16','=AY16',format_general)
    worksheet_summary.write_formula('K17','=AY17',format_general)
    worksheet_summary.write_formula('K18','=AY18',format_general)
    worksheet_summary.write_formula('K19','=AY19',format_general)
    worksheet_summary.write_formula('K20','=AY20',format_general)
    worksheet_summary.write_formula('K21','=AY21',format_general)
    worksheet_summary.write_formula('K22','=AY22',format_general)
    worksheet_summary.write_formula('K23','=AY23',format_subtotal)

    worksheet_summary.write_formula('L13','=AZ13',format_general)
    worksheet_summary.write_formula('L14','=AZ14',format_general)
    worksheet_summary.write_formula('L15','=AZ15',format_general)
    worksheet_summary.write_formula('L16','=AZ16',format_general)
    worksheet_summary.write_formula('L17','=AZ17',format_general)
    worksheet_summary.write_formula('L18','=AZ18',format_general)
    worksheet_summary.write_formula('L19','=AZ19',format_general)
    worksheet_summary.write_formula('L20','=AZ20',format_general)
    worksheet_summary.write_formula('L21','=AZ21',format_general)
    worksheet_summary.write_formula('L22','=AZ22',format_general)
    worksheet_summary.write_formula('L23','=AZ23',format_subtotal)

    worksheet_summary.write_formula('M13','=BA13',format_general)
    worksheet_summary.write_formula('M14','=BA14',format_general)
    worksheet_summary.write_formula('M15','=BA15',format_general)
    worksheet_summary.write_formula('M16','=BA16',format_general)
    worksheet_summary.write_formula('M17','=BA17',format_general)
    worksheet_summary.write_formula('M18','=BA18',format_general)
    worksheet_summary.write_formula('M19','=BA19',format_general)
    worksheet_summary.write_formula('M20','=BA20',format_general)
    worksheet_summary.write_formula('M21','=BA21',format_general)
    worksheet_summary.write_formula('M22','=BA22',format_general)
    worksheet_summary.write_formula('M23','=BA23',format_subtotal)

    worksheet_summary.write_formula('N13','=BB13',format_general)
    worksheet_summary.write_formula('N14','=BB14',format_general)
    worksheet_summary.write_formula('N15','=BB15',format_general)
    worksheet_summary.write_formula('N16','=BB16',format_general)
    worksheet_summary.write_formula('N17','=BB17',format_general)
    worksheet_summary.write_formula('N18','=BB18',format_general)
    worksheet_summary.write_formula('N19','=BB19',format_general)
    worksheet_summary.write_formula('N20','=BB20',format_general)
    worksheet_summary.write_formula('N21','=BB21',format_general)
    worksheet_summary.write_formula('N22','=BB22',format_general)
    worksheet_summary.write_formula('N23','=BB23',format_subtotal)


    """Format Corp Section"""

    """Format N88"""
    worksheet_summary.write_formula('A24','=AM24',format_general)
    worksheet_summary.write_formula('A25','=AM25',format_general)
    worksheet_summary.write('A26','',format_general)

    worksheet_summary.write_formula('B24','=AN24',format_general)
    worksheet_summary.write_formula('B25','=AN25',format_general)
    worksheet_summary.write('B26','Total',format_mini_total)

    worksheet_summary.write_formula('C24','=AO24',format_general)
    worksheet_summary.write_formula('C25','=AO25',format_general)
    worksheet_summary.write_formula('C26','=AO26',format_mini_total)

    worksheet_summary.write_formula('D24','=AP24',format_general)
    worksheet_summary.write_formula('D25','=AP25',format_general)
    worksheet_summary.write_formula('D26','=AP26',format_mini_total)

    worksheet_summary.write_formula('E24','=AQ24',format_general)
    worksheet_summary.write_formula('E25','=AQ25',format_general)
    worksheet_summary.write_formula('E26','=AQ26',format_mini_total)

    worksheet_summary.write_formula('F24','=AR24',format_general)
    worksheet_summary.write_formula('F25','=AR25',format_general)
    worksheet_summary.write_formula('F26','=AR26',format_mini_total)

    #DAILY CHANGE
    worksheet_summary.write_formula('I24','=AW24',format_general)
    worksheet_summary.write_formula('I25','=AW25',format_general)
    worksheet_summary.write('AI26','',format_general)

    worksheet_summary.write_formula('J24','=AX24',format_general)
    worksheet_summary.write_formula('J25','=AX25',format_general)
    worksheet_summary.write('J26','Total',format_mini_total)

    worksheet_summary.write_formula('K24','=AY24',format_general)
    worksheet_summary.write_formula('K25','=AY25',format_general)
    worksheet_summary.write_formula('K26','=AY26',format_mini_total)

    worksheet_summary.write_formula('L24','=AZ24',format_general)
    worksheet_summary.write_formula('L25','=AZ25',format_general)
    worksheet_summary.write_formula('L26','=AZ26',format_mini_total)

    worksheet_summary.write_formula('M24','=BA24',format_general)
    worksheet_summary.write_formula('M25','=BA25',format_general)
    worksheet_summary.write_formula('M26','=BA26',format_mini_total)

    worksheet_summary.write_formula('N24','=BB24',format_general)
    worksheet_summary.write_formula('N25','=BB25',format_general)
    worksheet_summary.write_formula('N26','=BB26',format_mini_total)

    """Format N90"""
    worksheet_summary.write_formula('A27','=AM27',format_general)
    worksheet_summary.write_formula('A28','=AM28',format_general)
    worksheet_summary.write('A29','',format_general)

    worksheet_summary.write_formula('B27','=AN27',format_general)
    worksheet_summary.write_formula('B28','=AN28',format_general)
    worksheet_summary.write('B29','Total',format_mini_total)

    worksheet_summary.write_formula('C27','=AO27',format_general)
    worksheet_summary.write_formula('C28','=AO28',format_general)
    worksheet_summary.write_formula('C29','=AO29',format_mini_total)

    worksheet_summary.write_formula('D27','=AP27',format_general)
    worksheet_summary.write_formula('D28','=AP28',format_general)
    worksheet_summary.write_formula('D29','=AP29',format_mini_total)

    worksheet_summary.write_formula('E27','=AQ27',format_general)
    worksheet_summary.write_formula('E28','=AQ28',format_general)
    worksheet_summary.write_formula('E29','=AQ29',format_mini_total)

    worksheet_summary.write_formula('F27','=AR27',format_general)
    worksheet_summary.write_formula('F28','=AR28',format_general)
    worksheet_summary.write_formula('F29','=AR29',format_mini_total)


    #DAILY CHANGE
    worksheet_summary.write_formula('I27','=AW27',format_general)
    worksheet_summary.write_formula('I28','=AW28',format_general)
    worksheet_summary.write('I29','',format_general)

    worksheet_summary.write_formula('J27','=AX27',format_general)
    worksheet_summary.write_formula('J28','=AX28',format_general)
    worksheet_summary.write('J29','Total',format_mini_total)

    worksheet_summary.write_formula('K27','=AY27',format_general)
    worksheet_summary.write_formula('K28','=AY28',format_general)
    worksheet_summary.write_formula('K29','=AY29',format_mini_total)

    worksheet_summary.write_formula('L27','=AZ27',format_general)
    worksheet_summary.write_formula('L28','=AZ28',format_general)
    worksheet_summary.write_formula('L29','=AZ29',format_mini_total)

    worksheet_summary.write_formula('M27','=BA27',format_general)
    worksheet_summary.write_formula('M28','=BA28',format_general)
    worksheet_summary.write_formula('M29','=BA29',format_mini_total)

    worksheet_summary.write_formula('N27','=BB27',format_general)
    worksheet_summary.write_formula('N28','=BB28',format_general)
    worksheet_summary.write_formula('N29','=BB29',format_mini_total)




    """Format P01"""
    worksheet_summary.write_formula('A30','=AM30',format_general)
    worksheet_summary.write_formula('A31','=AM31',format_general)
    worksheet_summary.write('A32','',format_general)

    worksheet_summary.write_formula('B30','=AN30',format_general)
    worksheet_summary.write_formula('B31','=AN31',format_general)
    worksheet_summary.write('B32','Total',format_mini_total)

    worksheet_summary.write_formula('C30','=AO30',format_general)
    worksheet_summary.write_formula('C31','=AO31',format_general)
    worksheet_summary.write_formula('C32','=AO32',format_mini_total)

    worksheet_summary.write_formula('D30','=AP30',format_general)
    worksheet_summary.write_formula('D31','=AP31',format_general)
    worksheet_summary.write_formula('D32','=AP32',format_mini_total)

    worksheet_summary.write_formula('E30','=AQ30',format_general)
    worksheet_summary.write_formula('E31','=AQ31',format_general)
    worksheet_summary.write_formula('E32','=AQ32',format_mini_total)

    worksheet_summary.write_formula('F30','=AR30',format_general)
    worksheet_summary.write_formula('F31','=AR31',format_general)
    worksheet_summary.write_formula('F32','=AR32',format_mini_total)


    #DAILY CHANGE
    worksheet_summary.write_formula('I30','=AW30',format_general)
    worksheet_summary.write_formula('I31','=AW31',format_general)
    worksheet_summary.write('I32','',format_general)

    worksheet_summary.write_formula('J30','=AX30',format_general)
    worksheet_summary.write_formula('J31','=AX31',format_general)
    worksheet_summary.write('J32','Total',format_mini_total)

    worksheet_summary.write_formula('K30','=AY30',format_general)
    worksheet_summary.write_formula('K31','=AY31',format_general)
    worksheet_summary.write_formula('K32','=AY32',format_mini_total)

    worksheet_summary.write_formula('L30','=AZ30',format_general)
    worksheet_summary.write_formula('L31','=AZ31',format_general)
    worksheet_summary.write_formula('L32','=AZ32',format_mini_total)

    worksheet_summary.write_formula('M30','=BA30',format_general)
    worksheet_summary.write_formula('M31','=BA31',format_general)
    worksheet_summary.write_formula('M32','=BA32',format_mini_total)

    worksheet_summary.write_formula('N30','=BB30',format_general)
    worksheet_summary.write_formula('N31','=BB31',format_general)
    worksheet_summary.write_formula('N32','=BB32',format_mini_total)



    """Format P02"""
    worksheet_summary.write_formula('A33','=AM33',format_general)
    worksheet_summary.write_formula('A34','=AM34',format_general)
    worksheet_summary.write('A35','',format_general)

    worksheet_summary.write_formula('B33','=AN33',format_general)
    worksheet_summary.write_formula('B34','=AN34',format_general)
    worksheet_summary.write('B35','Total',format_mini_total)

    worksheet_summary.write_formula('C33','=AO33',format_general)
    worksheet_summary.write_formula('C34','=AO34',format_general)
    worksheet_summary.write_formula('C35','=AO35',format_mini_total)

    worksheet_summary.write_formula('D33','=AP33',format_general)
    worksheet_summary.write_formula('D34','=AP34',format_general)
    worksheet_summary.write_formula('D35','=AP35',format_mini_total)

    worksheet_summary.write_formula('E33','=AQ33',format_general)
    worksheet_summary.write_formula('E34','=AQ34',format_general)
    worksheet_summary.write_formula('E35','=AQ35',format_mini_total)

    worksheet_summary.write_formula('F33','=AR33',format_general)
    worksheet_summary.write_formula('F34','=AR34',format_general)
    worksheet_summary.write_formula('F35','=AR35',format_mini_total)


    #DAILY CHANGE
    worksheet_summary.write_formula('I33','=AW33',format_general)
    worksheet_summary.write_formula('I34','=AW34',format_general)
    worksheet_summary.write('I35','',format_general)

    worksheet_summary.write_formula('J33','=AX33',format_general)
    worksheet_summary.write_formula('J34','=AX34',format_general)
    worksheet_summary.write('J35','Total',format_mini_total)

    worksheet_summary.write_formula('K33','=AY33',format_general)
    worksheet_summary.write_formula('K34','=AY34',format_general)
    worksheet_summary.write_formula('K35','=AY35',format_mini_total)

    worksheet_summary.write_formula('L33','=AZ33',format_general)
    worksheet_summary.write_formula('L34','=AZ34',format_general)
    worksheet_summary.write_formula('L35','=AZ35',format_mini_total)

    worksheet_summary.write_formula('M33','=BA33',format_general)
    worksheet_summary.write_formula('M34','=BA34',format_general)
    worksheet_summary.write_formula('M35','=BA35',format_mini_total)

    worksheet_summary.write_formula('N33','=BB33',format_general)
    worksheet_summary.write_formula('N34','=BB34',format_general)
    worksheet_summary.write_formula('N35','=BB35',format_mini_total)




    """Format K74"""
    worksheet_summary.write_formula('A36','=AM36',format_general)
    worksheet_summary.write_formula('A37','=AM37',format_general)
    worksheet_summary.write('A38','',format_general)

    worksheet_summary.write_formula('B36','=AN36',format_general)
    worksheet_summary.write_formula('B37','=AN37',format_general)
    worksheet_summary.write('B38','Total',format_mini_total)

    worksheet_summary.write_formula('C36','=AO36',format_general)
    worksheet_summary.write_formula('C37','=AO37',format_general)
    worksheet_summary.write_formula('C38','=AO38',format_mini_total)

    worksheet_summary.write_formula('D36','=AP36',format_general)
    worksheet_summary.write_formula('D37','=AP37',format_general)
    worksheet_summary.write_formula('D38','=AP38',format_mini_total)

    worksheet_summary.write_formula('E36','=AQ36',format_general)
    worksheet_summary.write_formula('E37','=AQ37',format_general)
    worksheet_summary.write_formula('E38','=AQ38',format_mini_total)

    worksheet_summary.write_formula('F36','=AR36',format_general)
    worksheet_summary.write_formula('F37','=AR37',format_general)
    worksheet_summary.write_formula('F38','=AR38',format_mini_total)


    #DAILY CHANGE
    worksheet_summary.write_formula('I36','=AW36',format_general)
    worksheet_summary.write_formula('I37','=AW37',format_general)
    worksheet_summary.write('I38','',format_general)

    worksheet_summary.write_formula('J36','=AX36',format_general)
    worksheet_summary.write_formula('J37','=AX37',format_general)
    worksheet_summary.write('J38','Total',format_mini_total)

    worksheet_summary.write_formula('K36','=AY36',format_general)
    worksheet_summary.write_formula('K37','=AY37',format_general)
    worksheet_summary.write_formula('K38','=AY38',format_mini_total)

    worksheet_summary.write_formula('L36','=AZ36',format_general)
    worksheet_summary.write_formula('L37','=AZ37',format_general)
    worksheet_summary.write_formula('L38','=AZ38',format_mini_total)

    worksheet_summary.write_formula('M36','=BA36',format_general)
    worksheet_summary.write_formula('M37','=BA37',format_general)
    worksheet_summary.write_formula('M38','=BA38',format_mini_total)

    worksheet_summary.write_formula('N36','=BB36',format_general)
    worksheet_summary.write_formula('N37','=BB37',format_general)
    worksheet_summary.write_formula('N38','=BB38',format_mini_total)




    """Format K
    
    
    """
    worksheet_summary.write_formula('A39','=AM39',format_general)
    worksheet_summary.write_formula('A40','=AM40',format_general)
    worksheet_summary.write('A41','',format_general)

    worksheet_summary.write_formula('B39','=AN39',format_general)
    worksheet_summary.write_formula('B40','=AN40',format_general)
    worksheet_summary.write('B41','Total',format_mini_total)

    worksheet_summary.write_formula('C39','=AO39',format_general)
    worksheet_summary.write_formula('C40','=AO40',format_general)
    worksheet_summary.write_formula('C41','=AO41',format_mini_total)

    worksheet_summary.write_formula('D39','=AP39',format_general)
    worksheet_summary.write_formula('D40','=AP40',format_general)
    worksheet_summary.write_formula('D41','=AP41',format_mini_total)

    worksheet_summary.write_formula('E39','=AQ39',format_general)
    worksheet_summary.write_formula('E40','=AQ40',format_general)
    worksheet_summary.write_formula('E41','=AQ41',format_mini_total)

    worksheet_summary.write_formula('F39','=AR39',format_general)
    worksheet_summary.write_formula('F40','=AR40',format_general)
    worksheet_summary.write_formula('F41','=AR41',format_mini_total)


    #DAILY CHANGE
    worksheet_summary.write_formula('I39','=AW39',format_general)
    worksheet_summary.write_formula('I40','=AW40',format_general)
    worksheet_summary.write('I41','',format_general)

    worksheet_summary.write_formula('J39','=AX39',format_general)
    worksheet_summary.write_formula('J40','=AX40',format_general)
    worksheet_summary.write('J41','Total',format_mini_total)

    worksheet_summary.write_formula('K39','=AY39',format_general)
    worksheet_summary.write_formula('K40','=AY40',format_general)
    worksheet_summary.write_formula('K41','=AY41',format_mini_total)

    worksheet_summary.write_formula('L39','=AZ39',format_general)
    worksheet_summary.write_formula('L40','=AZ40',format_general)
    worksheet_summary.write_formula('L41','=AZ41',format_mini_total)

    worksheet_summary.write_formula('M39','=BA39',format_general)
    worksheet_summary.write_formula('M40','=BA40',format_general)
    worksheet_summary.write_formula('M41','=BA41',format_mini_total)

    worksheet_summary.write_formula('N39','=BB39',format_general)
    worksheet_summary.write_formula('N40','=BB40',format_general)
    worksheet_summary.write_formula('N41','=BB41',format_mini_total)



    """Format P03"""
    worksheet_summary.write_formula('A42','=AM42',format_general)
    worksheet_summary.write_formula('A43','=AM43',format_general)
    worksheet_summary.write('A44','',format_general)

    worksheet_summary.write_formula('B42','=AN42',format_general)
    worksheet_summary.write_formula('B43','=AN43',format_general)
    worksheet_summary.write('B44','Total',format_mini_total)

    worksheet_summary.write_formula('C42','=AO42',format_general)
    worksheet_summary.write_formula('C43','=AO43',format_general)
    worksheet_summary.write_formula('C44','=AO44',format_mini_total)

    worksheet_summary.write_formula('D42','=AP42',format_general)
    worksheet_summary.write_formula('D43','=AP43',format_general)
    worksheet_summary.write_formula('D44','=AP44',format_mini_total)

    worksheet_summary.write_formula('E42','=AQ42',format_general)
    worksheet_summary.write_formula('E43','=AQ43',format_general)
    worksheet_summary.write_formula('E44','=AQ44',format_mini_total)

    worksheet_summary.write_formula('F42','=AR42',format_general)
    worksheet_summary.write_formula('F43','=AR43',format_general)
    worksheet_summary.write_formula('F44','=AR44',format_mini_total)


    #DAILY CHANGE
    worksheet_summary.write_formula('I42','=AW42',format_general)
    worksheet_summary.write_formula('I43','=AW43',format_general)
    worksheet_summary.write('I44','',format_general)

    worksheet_summary.write_formula('J42','=AX42',format_general)
    worksheet_summary.write_formula('J43','=AX43',format_general)
    worksheet_summary.write('J44','Total',format_mini_total)

    worksheet_summary.write_formula('K42','=AY42',format_general)
    worksheet_summary.write_formula('K43','=AY43',format_general)
    worksheet_summary.write_formula('K44','=AY44',format_mini_total)

    worksheet_summary.write_formula('L42','=AZ42',format_general)
    worksheet_summary.write_formula('L43','=AZ43',format_general)
    worksheet_summary.write_formula('L44','=AZ44',format_mini_total)

    worksheet_summary.write_formula('M42','=BA42',format_general)
    worksheet_summary.write_formula('M43','=BA43',format_general)
    worksheet_summary.write_formula('M44','=BA44',format_mini_total)

    worksheet_summary.write_formula('N42','=BB42',format_general)
    worksheet_summary.write_formula('N43','=BB43',format_general)
    worksheet_summary.write_formula('N44','=BB44',format_mini_total)




    """Format N87"""
    worksheet_summary.write_formula('A45','=AM45',format_general)
    worksheet_summary.write_formula('A46','=AM46',format_general)
    worksheet_summary.write('A47','',format_general)

    worksheet_summary.write_formula('B45','=AN45',format_general)
    worksheet_summary.write_formula('B46','=AN46',format_general)
    worksheet_summary.write('B47','Total',format_mini_total)

    worksheet_summary.write_formula('C45','=AO45',format_general)
    worksheet_summary.write_formula('C46','=AO46',format_general)
    worksheet_summary.write_formula('C47','=AO47',format_mini_total)

    worksheet_summary.write_formula('D45','=AP45',format_general)
    worksheet_summary.write_formula('D46','=AP46',format_general)
    worksheet_summary.write_formula('D47','=AP47',format_mini_total)

    worksheet_summary.write_formula('E45','=AQ45',format_general)
    worksheet_summary.write_formula('E46','=AQ46',format_general)
    worksheet_summary.write_formula('E47','=AQ47',format_mini_total)

    worksheet_summary.write_formula('F45','=AR45',format_general)
    worksheet_summary.write_formula('F46','=AR46',format_general)
    worksheet_summary.write_formula('F47','=AR47',format_mini_total)


    #DAILY CHANGE
    worksheet_summary.write_formula('I45','=AW45',format_general)
    worksheet_summary.write_formula('I46','=AW46',format_general)
    worksheet_summary.write('I47','',format_general)

    worksheet_summary.write_formula('J45','=AX45',format_general)
    worksheet_summary.write_formula('J46','=AX46',format_general)
    worksheet_summary.write('J47','Total',format_mini_total)

    worksheet_summary.write_formula('K45','=AY45',format_general)
    worksheet_summary.write_formula('K46','=AY46',format_general)
    worksheet_summary.write_formula('K47','=AY47',format_mini_total)

    worksheet_summary.write_formula('L45','=AZ45',format_general)
    worksheet_summary.write_formula('L46','=AZ46',format_general)
    worksheet_summary.write_formula('L47','=AZ47',format_mini_total)

    worksheet_summary.write_formula('M45','=BA45',format_general)
    worksheet_summary.write_formula('M46','=BA46',format_general)
    worksheet_summary.write_formula('M47','=BA47',format_mini_total)

    worksheet_summary.write_formula('N45','=BB45',format_general)
    worksheet_summary.write_formula('N46','=BB46',format_general)
    worksheet_summary.write_formula('N47','=BB47',format_mini_total)


    """Corp Total"""
    worksheet_summary.write('A48', 'Corp Total',format_subtotal)
    worksheet_summary.write('B48', ' ',format_subtotal)
    worksheet_summary.write('I48', 'Corp Total',format_subtotal)
    worksheet_summary.write('J48', ' ',format_subtotal)
    worksheet_summary.write_formula('C48', '=SUM(C47,C44,C41,C38,C35,C32,C29,C26)',format_subtotal)
    worksheet_summary.write_formula('D48', '=SUM(D47,D44,D41,D38,D35,D32,D29,D26)',format_subtotal)
    worksheet_summary.write_formula('E48', '=SUM(E47,E44,E41,E38,E35,E32,E29,E26)',format_subtotal)
    worksheet_summary.write_formula('F48', '=SUM(F47,F44,F41,F38,F35,F32,F29,F26)',format_subtotal)
    worksheet_summary.write_formula('K48', '=SUM(K47,K44,K41,K38,K35,K32,K29,K26)',format_subtotal)
    worksheet_summary.write_formula('L48', '=SUM(L47,L44,L41,L38,L35,L32,L29,L26)',format_subtotal)
    worksheet_summary.write_formula('M48', '=SUM(M47,M44,M41,M38,M35,M32,M29,M26)',format_subtotal)
    worksheet_summary.write_formula('N48', '=SUM(N47,N44,N41,N38,N35,N32,N29,N26)',format_subtotal)


    """CD Formating"""
    worksheet_summary.write('A50','CD Total',format_subtotal)
    worksheet_summary.write('B50','',format_subtotal)
    worksheet_summary.write_formula('C50','=AP50',format_subtotal)
    worksheet_summary.write_formula('D50','=AQ50',format_subtotal)
    worksheet_summary.write_formula('E50','=AR50',format_subtotal)
    worksheet_summary.write_formula('F50','=AS50',format_subtotal)

    worksheet_summary.write('I50','CD Total',format_subtotal)
    worksheet_summary.write('J50','',format_subtotal)
    worksheet_summary.write_formula('K50','=AY50',format_subtotal)
    worksheet_summary.write_formula('L50','=AZ50',format_subtotal)
    worksheet_summary.write_formula('M50','=BA50',format_subtotal)
    worksheet_summary.write_formula('N50','=BB50',format_subtotal)



    """CMO Formating"""
    worksheet_summary.write_formula('A51','=AN56',format_general)
    worksheet_summary.write('B51','',format_general)
    worksheet_summary.write_formula('C51','=AP56',format_general)
    worksheet_summary.write_formula('D51','=AQ56',format_general)
    worksheet_summary.write_formula('E51','=AR56',format_general)
    worksheet_summary.write_formula('F51','=AS56',format_general)

    worksheet_summary.write_formula('I51','=AW56',format_general)
    worksheet_summary.write('J51','',format_general)
    worksheet_summary.write_formula('K51','=AY56',format_general)
    worksheet_summary.write_formula('L51','=AZ56',format_general)
    worksheet_summary.write_formula('M51','=BB56',format_general)
    worksheet_summary.write_formula('N51','=BA56',format_general)

    worksheet_summary.write_formula('A52','=AN57',format_general)
    worksheet_summary.write('B52','',format_general)
    worksheet_summary.write_formula('C52','=AP57',format_general)
    worksheet_summary.write_formula('D52','=AQ57',format_general)
    worksheet_summary.write_formula('E52','=AR57',format_general)
    worksheet_summary.write_formula('F52','=AS57',format_general)

    worksheet_summary.write_formula('I52','=AW57',format_general)
    worksheet_summary.write('J52','',format_general)
    worksheet_summary.write_formula('K52','=AY57',format_general)
    worksheet_summary.write_formula('L52','=AZ57',format_general)
    worksheet_summary.write_formula('M52','=BA57',format_general)
    worksheet_summary.write_formula('N52','=BB57',format_general)

    worksheet_summary.write_formula('A53','=AN58',format_subtotal)
    worksheet_summary.write('B53','',format_subtotal)
    worksheet_summary.write_formula('C53','=AP58',format_subtotal)
    worksheet_summary.write_formula('D53','=AQ58',format_subtotal)
    worksheet_summary.write_formula('E53','=AR58',format_subtotal)
    worksheet_summary.write_formula('F53','=AS58',format_subtotal)

    worksheet_summary.write_formula('I53','=AW58',format_subtotal)
    worksheet_summary.write('J53','',format_subtotal)
    worksheet_summary.write_formula('K53','=AY58',format_subtotal)
    worksheet_summary.write_formula('L53','=AZ58',format_subtotal)
    worksheet_summary.write_formula('M53','=BA58',format_subtotal)
    worksheet_summary.write_formula('N53','=BB58',format_subtotal)

    """FIRM TOTAL SUMS"""
    worksheet_summary.write('A55','Firm Total',format_subtotal)
    worksheet_summary.write('B55','',format_subtotal)
    worksheet_summary.write_formula('C55','=SUM(C53,C50,C48,C23)',format_subtotal)
    worksheet_summary.write_formula('D55','=SUM(D53,D50,D48,D23)',format_subtotal)
    worksheet_summary.write_formula('E55','=SUM(E53,E50,E48,E23)',format_subtotal)
    worksheet_summary.write_formula('F55','=SUM(F53,F50,F48,F23)',format_subtotal)

    worksheet_summary.write('I55','Firm Total',format_subtotal)
    worksheet_summary.write('J55','',format_subtotal)
    worksheet_summary.write_formula('K55','=SUM(K53,K50,K48,K23)',format_subtotal)
    worksheet_summary.write_formula('L55','=SUM(L53,L50,L48,L23)',format_subtotal)
    worksheet_summary.write_formula('M55','=SUM(M53,M50,M48,M23)',format_subtotal)
    worksheet_summary.write_formula('N55','=SUM(N53,N50,N48,N23)',format_subtotal)


    """Conditional Formating"""
    worksheet_summary.conditional_format('L13:L65', {'type':'cell',
                                            'criteria': '<',
                                            'value':    0,
                                            'format':   format_general_row_green})
    worksheet_summary.conditional_format('L13:L65', {'type':'cell',
                                            'criteria': '>',
                                            'value':    0,
                                            'format':   format_general_row_red})
    worksheet_summary.conditional_format('M13:N65', {'type':'cell',
                                            'criteria': '<',
                                            'value':    0,
                                            'format':   format_general_row_red})
    worksheet_summary.conditional_format('M13:N65', {'type':'cell',
                                            'criteria': '>',
                                            'value':    0,
                                            'format':   format_general_row_green})

    """Set Row & Column height/width"""

    worksheet_summary.set_row(48,3)
    worksheet_summary.set_row(53,3)

    worksheet_summary.set_column('A:A',15,None)
    worksheet_summary.set_column('B:B',13,None)
    worksheet_summary.set_column('C:C',13,None)
    worksheet_summary.set_column('D:D',13,None)
    worksheet_summary.set_column('E:E',13,None)
    worksheet_summary.set_column('F:F',13,None)
    worksheet_summary.set_column('G:G',3,None)
    worksheet_summary.set_column('H:H',3,None)
    worksheet_summary.set_column('I:I',15,None)
    worksheet_summary.set_column('J:J',13,None)
    worksheet_summary.set_column('K:K',13,None)
    worksheet_summary.set_column('L:L',13,None)
    worksheet_summary.set_column('M:M',13,None)
    worksheet_summary.set_column('N:N',13,None)
    worksheet_summary.hide_gridlines(2)
    worksheet_summary.insert_image('L1', 'P:/1. Individual Folders/Chad/Python Scripts/PNL Report/Logo.png',{'x_scale':.7,'y_scale':.7})

    worksheet_summary.write_url('A2',"internal:'Position DSP'!A1",format_url_links,string = '1. Position DSP')
    worksheet_summary.write_url('A3',"internal:'Real PnL DSP'!A1",format_url_links,string = '2. Real PnL DSP')
    worksheet_summary.write_url('A4',"internal:'Adj Unrealized PnL'!A1",format_url_links,string = '3. Adj Unrealized PnL')
    worksheet_summary.write_url('A5',"internal:'Requirement Change'!A1",format_url_links,string = '4. Requirement Change')
    worksheet_summary.write_url('A6',"internal:'HT Detail'!A1",format_url_links,string = '5. HT Detail')
    worksheet_summary.write_url('A7',"internal:'BBG Detail'!A1",format_url_links ,string = '6. BBG Detail')
    

    """New Muni Short Check"""
    worksheet_summary.write_formula('J13','=IF(AY101=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J14','=IF(AY102=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J15','=IF(AY103=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J16','=IF(AY104=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J17','=IF(AY105=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J18','=IF(AY106=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J19','=IF(AY107=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J20','=IF(AY108=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J21','=IF(AY109=TRUE,"Short","`")',format_general)
    worksheet_summary.write_formula('J22','=IF(AY110=TRUE,"Short","`")',format_general)





    """
    Position DSP Formatting
    
    """
    Position_DSP = position_dsp.Position_DSP()
    Position_DSP[0].to_excel(writer,sheet_name = 'Position DSP', index = False)
    Cleared_Position = Position_DSP[1]
    
    Cleared_Position.to_excel(writer,sheet_name = 'Position DSP',index = False, startcol = 8)

    worksheet_Position_DSP = writer.sheets['Position DSP']

    worksheet_Position_DSP.set_column('A:A',28,format_general_row)
    worksheet_Position_DSP.set_column('B:B',12,format_general_row)
    worksheet_Position_DSP.set_column('C:C',10,format_general_row)
    worksheet_Position_DSP.set_column('D:D',11,format_general_row)
    worksheet_Position_DSP.set_column('E:E',15,format_general_row)
    worksheet_Position_DSP.set_column('F:F',7,format_general_row)
    worksheet_Position_DSP.set_column('G:G',12,format_general_row)


    worksheet_Position_DSP.write('A1','Security',format_top_summary)
    worksheet_Position_DSP.write('B1','Account Name',format_top_summary)
    worksheet_Position_DSP.write('C1','CUSIP',format_top_summary)
    worksheet_Position_DSP.write('D1','QTY DSP',format_top_summary)
    worksheet_Position_DSP.write('E1','Position Notes',format_top_summary)


    worksheet_Position_DSP.set_column('I:I',28,format_general_row)
    worksheet_Position_DSP.set_column('J:J',12,format_general_row)
    worksheet_Position_DSP.set_column('K:K',10,format_general_row)
    worksheet_Position_DSP.set_column('L:L',11,format_general_row)
    worksheet_Position_DSP.set_column('M:M',15,format_general_row)
    worksheet_Position_DSP.set_column('N:N',7,format_general_row)
    worksheet_Position_DSP.set_column('O:O',12,format_general_row)

    worksheet_Position_DSP.write('I1','Security',format_top_summary)
    worksheet_Position_DSP.write('J1','Account Name',format_top_summary)
    worksheet_Position_DSP.write('K1','CUSIP',format_top_summary)
    worksheet_Position_DSP.write('L1','QTY DSP',format_top_summary)
    worksheet_Position_DSP.write('M1','Position Notes',format_top_summary)
    worksheet_Position_DSP.autofilter('A1:E20000')
    worksheet_Position_DSP.freeze_panes(1, 0)
    """
    PnL DSP XLSX Formatting
    
    """

    PnL_DSP_Tables = real_pnl_dsp.PnL_DSP()
    Real_PnL_DSP = PnL_DSP_Tables[0]
    Real_PnL_DSP.to_excel(writer,sheet_name = 'Real PnL DSP',index = False)
    worksheet_PnL_DSP = writer.sheets['Real PnL DSP']
    worksheet_PnL_DSP.set_column('A:A',8,format_general_row)
    worksheet_PnL_DSP.set_column('B:B',27,format_general_row)
    worksheet_PnL_DSP.set_column('C:C',12,format_general_row)
    worksheet_PnL_DSP.set_column('D:D',12,format_general_row)
    worksheet_PnL_DSP.set_column('E:E',12,format_general_row)
    worksheet_PnL_DSP.set_column('F:F',12,format_general_row)

 
    worksheet_PnL_DSP.write('A1','Date',format_top_summary_textwrap)
    worksheet_PnL_DSP.write('B1','Security',format_top_summary_textwrap)
    worksheet_PnL_DSP.write('C1','Account Name',format_top_summary_textwrap)
    worksheet_PnL_DSP.write('D1','CUSIP',format_top_summary_textwrap)
    worksheet_PnL_DSP.write('E1','Real PnL DSP',format_top_summary_textwrap)
    worksheet_PnL_DSP.write('F1','Notes',format_top_summary_textwrap)
    worksheet_PnL_DSP.freeze_panes(1, 0)


    Running_PnL_DSP = PnL_DSP_Tables[1]
    Running_PnL_DSP.to_excel(writer,sheet_name = 'Real PnL DSP',startrow = 1,startcol = 7, index = False)

    worksheet_PnL_DSP.merge_range('H1:O1', 'Unresolved PnL DSP', merge_format)
    worksheet_PnL_DSP.set_column('H:H',8,format_general_row)
    worksheet_PnL_DSP.set_column('I:I',27,format_general_row)
    worksheet_PnL_DSP.set_column('J:J',12,format_general_row)
    worksheet_PnL_DSP.set_column('K:K',14,format_general_row)
    worksheet_PnL_DSP.set_column('L:L',14,format_general_row)
    worksheet_PnL_DSP.set_column('M:M',13,format_general_row)
    worksheet_PnL_DSP.set_column('N:N',10,format_general_row)
    worksheet_PnL_DSP.set_column('O:O',12,format_general_row)

    worksheet_PnL_DSP.write('H2','Date',format_top_summary)
    worksheet_PnL_DSP.write('I2','Security',format_top_summary)
    worksheet_PnL_DSP.write('J2','Account Name',format_top_summary)
    worksheet_PnL_DSP.write('K2','CUSIP',format_top_summary)
    worksheet_PnL_DSP.write('L2','Previous PnL DSP',format_top_summary)
    worksheet_PnL_DSP.write('M2','Current PnL DSP',format_top_summary)
    worksheet_PnL_DSP.write('N2','Net PnL DSP',format_top_summary)
    worksheet_PnL_DSP.write('O2','Notes',format_top_summary)
    worksheet_PnL_DSP.autofilter('A1:F20000')
    """
    Adj Unrealized PnL Change

    """
    HT_Detail = ht_detail.HT_Detail_Generate()
    Adj_Unrealized_Change = HT_Detail[1]
    Adj_Unrealized_Change.to_excel(writer,sheet_name = 'Adj Unrealized PnL',index = False)
    worksheet_Adj_Unrealized_Change = writer.sheets['Adj Unrealized PnL']

    worksheet_Adj_Unrealized_Change.set_column('A:A',30,format_general_row)
    worksheet_Adj_Unrealized_Change.set_column('B:B',10,format_general_row)
    worksheet_Adj_Unrealized_Change.set_column('C:C',10,format_general_row)
    worksheet_Adj_Unrealized_Change.set_column('D:D',19,format_general_row)
    worksheet_Adj_Unrealized_Change.set_column('G:G',10,format_general_row)
    worksheet_Adj_Unrealized_Change.set_column('H:H',19,format_general_row)

    worksheet_Adj_Unrealized_Change.write('A1','Security',format_top_summary)
    worksheet_Adj_Unrealized_Change.write('B1','CUSIP',format_top_summary)
    worksheet_Adj_Unrealized_Change.write('C1','Account',format_top_summary)
    worksheet_Adj_Unrealized_Change.write('D1','Adj Unreal PnL Change',format_top_summary)
    worksheet_Adj_Unrealized_Change.write('H1','Adj Unrealized Pnl',format_top_summary)
    worksheet_Adj_Unrealized_Change.write('G1',' ',format_top_summary)
    worksheet_Adj_Unrealized_Change.write('G2','Total',format_top_summary)
    worksheet_Adj_Unrealized_Change.write_formula('H2','=SUM(D2:D1048576)',format_general_row)
    worksheet_Adj_Unrealized_Change.freeze_panes(1, 0)
    worksheet_Adj_Unrealized_Change.autofilter('A1:D20000')

    """
    Requirement Change

    """
    Requirement_Change = HT_Detail[2]
    Requirement_Change.to_excel(writer,sheet_name = 'Requirement Change',index = False)
    worksheet_Requirement_Change = writer.sheets['Requirement Change']

    worksheet_Requirement_Change.set_column('A:A',30,format_general_row)
    worksheet_Requirement_Change.set_column('B:B',10,format_general_row)
    worksheet_Requirement_Change.set_column('C:C',10,format_general_row)
    worksheet_Requirement_Change.set_column('D:D',19,format_general_row)
    worksheet_Requirement_Change.set_column('H:H',17,format_general_row)

    worksheet_Requirement_Change.write('A1','Security',format_top_summary)
    worksheet_Requirement_Change.write('B1','CUSIP',format_top_summary)
    worksheet_Requirement_Change.write('C1','Account',format_top_summary)
    worksheet_Requirement_Change.write('D1','Requirement Change',format_top_summary)
    worksheet_Requirement_Change.write('H1','Requirement Change',format_top_summary)
    worksheet_Requirement_Change.write('G1',' ',format_top_summary)
    worksheet_Requirement_Change.write('G2','Total',format_top_summary)
    worksheet_Requirement_Change.write_formula('H2','=SUM(D2:D1048576)',format_general_row)
    worksheet_Requirement_Change.freeze_panes(1, 0)
    worksheet_Requirement_Change.autofilter('A1:D20000')


    """
    HT Detail Formating

    """
    HT_Detail[0].to_excel(writer,sheet_name = 'HT Detail',startrow = 4,index = False)
    worksheet_HT_Detail = writer.sheets['HT Detail']

    worksheet_HT_Detail.write('A2','Column Totals',format_top_summary)

    worksheet_HT_Detail.write_formula('E2','=SUM(E6:E1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('F2','=SUM(F6:F1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('G2','=SUM(G6:G1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('H2','=SUM(H6:H1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('I2','=SUM(I6:I1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('J2','=SUM(J6:J1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('K2','=SUM(K6:K1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('L2','=SUM(L6:L1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('M2','=SUM(M6:M1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('N2','=SUM(N6:N1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('O2','=SUM(O6:O1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('P2','=SUM(P6:P1048576)',format_general_row)
    worksheet_HT_Detail.write_formula('Q2','=SUM(Q6:Q1048576)',format_general_row)

    worksheet_HT_Detail.set_column('A:A',27,format_general)
    worksheet_HT_Detail.set_column('B:B',10,format_general)
    worksheet_HT_Detail.set_column('C:C',9,format_general)
    worksheet_HT_Detail.set_column('D:D',7,format_general_decimal)
    worksheet_HT_Detail.set_column('E:E',9,format_general)
    worksheet_HT_Detail.set_column('F:F',9,format_general)
    worksheet_HT_Detail.set_column('G:G',9,format_general)
    worksheet_HT_Detail.set_column('H:H',14,format_general)
    worksheet_HT_Detail.set_column('I:I',20,format_general)
    worksheet_HT_Detail.set_column('J:J',21,format_general)
    worksheet_HT_Detail.set_column('K:K',20,format_general)
    worksheet_HT_Detail.set_column('L:L',16,format_general)
    worksheet_HT_Detail.set_column('M:M',9,format_general)
    worksheet_HT_Detail.set_column('N:N',15,format_general)
    worksheet_HT_Detail.set_column('O:O',19,format_general)
    worksheet_HT_Detail.set_column('P:P',19,format_general)
    worksheet_HT_Detail.set_column('Q:Q',19,format_general)

    worksheet_HT_Detail.write('A1',' ',format_top_summary)
    worksheet_HT_Detail.write('B1','CUSIP',format_top_summary)
    worksheet_HT_Detail.write('C1','Account',format_top_summary)
    worksheet_HT_Detail.write('D1','Price',format_top_summary)
    worksheet_HT_Detail.write('E1','BBG QTY',format_top_summary)
    worksheet_HT_Detail.write('F1','HT QTY',format_top_summary)
    worksheet_HT_Detail.write('G1','QTY DSP',format_top_summary)
    worksheet_HT_Detail.write('H1','HT QTY Change',format_top_summary)
    worksheet_HT_Detail.write('I1','HT Current Unreal PnL',format_top_summary)
    worksheet_HT_Detail.write('J1','HT Prevoius Unreal PnL',format_top_summary)
    worksheet_HT_Detail.write('K1','Adj Unreal PnL Change',format_top_summary)
    worksheet_HT_Detail.write('L1','Real PnL Change',format_top_summary)
    worksheet_HT_Detail.write('M1','BBG PnL',format_top_summary)
    worksheet_HT_Detail.write('N1','HT-BBG PnL DSP',format_top_summary)
    worksheet_HT_Detail.write('O1','Requirement Change',format_top_summary)
    worksheet_HT_Detail.write('P1','Current Requirement',format_top_summary)
    worksheet_HT_Detail.write('Q1','Req %',format_top_summary)

    worksheet_HT_Detail.write('A5','Security',format_top_summary)
    worksheet_HT_Detail.write('B5','CUSIP',format_top_summary)
    worksheet_HT_Detail.write('C5','Account',format_top_summary)
    worksheet_HT_Detail.write('D5','Price',format_top_summary)
    worksheet_HT_Detail.write('E5','BBG QTY',format_top_summary)
    worksheet_HT_Detail.write('F5','HT QTY',format_top_summary)
    worksheet_HT_Detail.write('G5','QTY DSP',format_top_summary)
    worksheet_HT_Detail.write('H5','HT QTY Change',format_top_summary)
    worksheet_HT_Detail.write('I5','HT Current Unreal PnL',format_top_summary)
    worksheet_HT_Detail.write('J5','HT Prevoius Unreal PnL',format_top_summary)
    worksheet_HT_Detail.write('K5','Adj Unreal PnL Change',format_top_summary)
    worksheet_HT_Detail.write('L5','Real PnL Change',format_top_summary)
    worksheet_HT_Detail.write('M5','BBG PnL',format_top_summary)
    worksheet_HT_Detail.write('N5','HT-BBG PnL DSP',format_top_summary)
    worksheet_HT_Detail.write('O5','Requirement Change',format_top_summary)
    worksheet_HT_Detail.write('P5','Current Requirement',format_top_summary)
    worksheet_HT_Detail.write('Q5','Req %',format_top_summary)
    worksheet_HT_Detail.freeze_panes(1, 0)
    worksheet_HT_Detail.autofilter('A5:Q20000')

    """
    BBG Detail Formating

    """
    BBG_Detail = HT_Detail[3]
    BBG_Detail.to_excel(writer,sheet_name = 'BBG Detail',index = False)
    worksheet_BBG_Detail = writer.sheets['BBG Detail']
    
    worksheet_BBG_Detail.set_column('A:A',30,format_general)
    worksheet_BBG_Detail.set_column('B:B',13,format_general)
    worksheet_BBG_Detail.set_column('C:C',13,format_general)
    worksheet_BBG_Detail.set_column('D:D',13,format_general)
    worksheet_BBG_Detail.set_column('E:E',13,format_general)
    worksheet_BBG_Detail.set_column('F:F',13,format_general)
    worksheet_BBG_Detail.set_column('G:G',13,format_general_decimal)

    worksheet_BBG_Detail.write('A1','Security',format_top_summary)
    worksheet_BBG_Detail.write('B1','CUSIP',format_top_summary)
    worksheet_BBG_Detail.write('C1','PnL',format_top_summary)
    worksheet_BBG_Detail.write('D1','Account',format_top_summary)
    worksheet_BBG_Detail.write('E1','Position',format_top_summary)
    worksheet_BBG_Detail.write('F1','MTG Position',format_top_summary)
    worksheet_BBG_Detail.write('G1','Average Cost',format_top_summary)
    worksheet_BBG_Detail.freeze_panes(1, 0)
    worksheet_BBG_Detail.autofilter('A1:G20000')

    worksheet_BBG_Detail = writer.sheets['BBG Detail']




    writer.save()

