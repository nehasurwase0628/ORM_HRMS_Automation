import os
from configparser import ConfigParser

import pandas as pd

print(os.path.dirname(__file__))
print(os.pardir)
files_location = os.path.dirname(__file__) + os.sep + os.pardir + 'Configuration' + os.sep
print(files_location)
filepath = files_location + 'ORMApplication.properties'
print("File path for config", filepath)
con = ConfigParser()
con.read(filepath)

username_elk =  con.get('CREDENTIALS', 'username')
password_elk =  con.get('CREDENTIALS', 'password')
indexname_elk =  con.get('INDEXNAME', 'index_name')
elkirl =  con.get('ELKURL', 'elk_url')


def convert_file_to_dataframe(file_path, sheet, unique_col, sep):
    """Convert excel or csv file to dataframe"""
    if file_path.endswith('.xls') or file_path.endswith('.xlsx') or file_path.endswith('.xlsb') or file_path.endswith('.XLS'):
        dataframe = ''
        if file_path.endswith('.xls') or file_path.endswith('.xlsx') or file_path.endswith('.XLS'):
            dataframe = pd.read_excel(file_path, sheet_name=sheet, dtype=str, keep_default_na=False)


