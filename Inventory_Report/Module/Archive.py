import config as config
import datetime as datetime
import xlsxwriter
import pandas as pd
import Module.Date_File_Path as date_file_path
import shutil

def Write_to_Drive(write_to_file_path,copy_from_file_path):
    recent_date = date_file_path.Date_Recent()
    file_path = write_to_file_path + recent_date + '.xlsx'
    shutil.copy(copy_from_file_path,file_path)
