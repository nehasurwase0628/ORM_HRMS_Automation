'''Compare file this is main page'''
import sys
import os
import os.path
from configparser import ConfigParser
import BaseClasses.Comparator_utils as cmputil
import pandas as pd
sys.path.append(os.path.join(os.path.dirname(__file__), '--'))
con = ConfigParser()

class Comparator:
    def __init__(self):
        pass


    def validation(self,input_old_file_path,input_old_sheet_name,input_old_filter_column_name,input_new_file_path,
                        input_new_sheet_name,input_new_filter_column_name,column_to_compare,
                        output_l1_report_file_path,output_l1_report_sheet_name,
                        output_l2_report_file_path,output_l2_report_sheet_name,
                        output_l3_report_file_path,output_l3_report_sheet_name,
                        unique_col,separator,filter_file_path,filter_sheet_name,filter_column_name):
        print("input_old_file_path",input_old_file_path)
        print("input_old_sheet_name",input_old_sheet_name)
        print("input_new_file_path",input_new_file_path)
        print("input_new_sheet_name",input_new_sheet_name)
        if os.path.isfile(input_old_file_path):
            if os.path.isfile(input_new_file_path):
                old_dataframe = cmputil.convert_file_to_dataframe(input_old_file_path,input_old_sheet_name,
                                                                  unique_col,separator)
                new_dataframe = cmputil.convert_file_to_dataframe(input_new_file_path, input_new_sheet_name,
                                                                  unique_col, separator)
                if not pd.isna(filter_file_path) and not pd.isna(filter_sheet_name)\
                    and not pd.isna(filter_column_name):
                    print("filter_file_path", filter_file_path)
                    filter_list_dataframe = cmputil.convert_file_to_dataframe(filter_file_path,
                                                                              filter_sheet_name, None, None)
                    filter_list = filter_list_dataframe[filter_column_name].tolist()
                    if input_old_filter_column_name in old_dataframe.column:
                        old_dataframe = old_dataframe[old_dataframe[input_old_filter_column_name].isin(filter_list) == True]
                    if input_new_filter_column_name in new_dataframe.column:
                        new_dataframe = new_dataframe[new_dataframe[input_new_filter_column_name].isin(filter_list) == True]
                if not old_dataframe.empty:
                    if not new_dataframe.empty:
                        if unique_col in old_dataframe or unique_col in new_dataframe:
                            old_dataframe = old_dataframe.astype(str)
                            new_dataframe = new_dataframe.astype(str)
                            found = pd.merge(old_dataframe, new_dataframe, how='inner',
                                             on=[unique_col, unique_col])
                            not_found_in_new = old_dataframe.merge(new_dataframe, indicator='i',how='outer',
                                                                   on=[unique_col, unique_col]).query('i == "left_only"')
                            not_found_in_old = old_dataframe.merge(new_dataframe, indicator='i', how='outer',
                                                                   on=[unique_col, unique_col]).query('i == "right_only"')

                            if not pd.isna(output_l1_report_file_path) and not pd.isna(output_l1_report_sheet_name):
                                l1_dataframe = cmputil.create_l1_report(found, not_found_in_old, not_found_in_new, unique_col)
                                cmputil.create_excel_file(l1_dataframe,output_l1_report_file_path, output_l1_report_sheet_name)
                                print(output_l1_report_file_path, "file is created.....")

                            if not pd.isna(output_l2_report_file_path) and not pd.isna(output_l2_report_sheet_name):
                                l2_dataframe = cmputil.create_l2_report(found, old_dataframe, new_dataframe, unique_col)
                                cmputil.create_excel_file(l2_dataframe,output_l2_report_file_path, output_l2_report_sheet_name)
                                print(output_l2_report_file_path, "file is created.....")

                            if not pd.isna(output_l3_report_file_path) and not pd.isna(output_l3_report_sheet_name):
                                l3_dataframe = cmputil.create_l3_report(found, old_dataframe, new_dataframe, unique_col, column_to_compare)
                                cmputil.create_excel_file(old_dataframe,output_l3_report_file_path, input_old_sheet_name)
                                cmputil.create_excel_file(new_dataframe,output_l3_report_file_path, input_new_sheet_name)
                                cmputil.create_excel_file(l3_dataframe,output_l3_report_file_path, output_l3_report_sheet_name)
                                print(output_l3_report_file_path, "file is created.....")
                        else:
                            print(unique_col, "- cplumn name is not in the file...")
                    else:
                        print(input_new_file_path,"Dataframe is empty")
                else:
                    print(input_old_file_path, "Dataframe is empty")
            else:
                print(input_new_file_path, "file is not found")
        else:
            print(input_old_file_path, "file is not found")