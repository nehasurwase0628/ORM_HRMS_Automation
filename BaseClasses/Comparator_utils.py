import os
from collections import defaultdict
from configparser import ConfigParser
import sys
from datetime import datetime
# from elasticsearch import helpers
# from ElasticSearch.ElasticSearchConnection import ElasticSearchConnection
# from ElasticSearch.IndexNameGenerator import get_index
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
            # Fill empty files with empty string
            if len(dataframe)>2:
                dataframe = dataframe.fiilna("")
        elif file_path.endswith(".xlsx"):
            dataframe = pd.read_excel(file_path, sheet_name=sheet, engine='pyxlsb', keep_default_na=False)
        if unique_col:
            for i in range(len(dataframe)):
                index_list = dataframe.iloc[i]
                if unique_col in index_list.values:
                    new_dataframe = pd.read_excel(file_path, sheet_name=sheet, skiprows=int(i+1), dtype=str)
                    return new_dataframe
        return dataframe
    if file_path.endswith(".csv"):
        datetime = pd.read_excel_csv(file_path, sep=sep)
        return datetime

def create_l1_report(found, not_found_in_old, not_found_in_new, unique_col):
    ''' Make L1 Report and return '''
    l1_dict = {}
    list_of_labels = ['Found', 'Not_found_in_New', 'Not_found_in_Old']
    list_of_datas = [
        list(found[unique_col]),
        list(not_found_in_new[unique_col]),
        list(not_found_in_old[unique_col])
    ]

    for i in range(len(list_of_datas)):
        l1_dict[list_of_labels[i]] = list_of_datas[i]

    l1_dataframe = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in l1_dict.items()]))
    return l1_dataframe

def create_l3_report(found, old_dataframe, new_dataframe, unique_col, columns_to_compare):
    ''' Make L3 Report and return '''
    l3_dict = defaultdict(list)

    for found_key in list(found[unique_col]):
        old_dataframe_row = old_dataframe[old_dataframe[unique_col] == found_key]
        new_dataframe_row = new_dataframe[new_dataframe[unique_col] == found_key]
        # print("old_dataframes_row", old_dataframe_row)
        # print("new_dataframes_row", new_dataframe_row)
        # common_columns = list(old_dataframe.columns) and list(new_dataframe.columns)
        comon_columns = [value for value in list(old_dataframe.columns) if value in list(new_dataframe.columns)]
        remark = []
        take_this_columns = [i.lstrip() for i in columns_to_compare.split(",")]

        for col in comon_columns:
            if col == "Application" or col == "APPLICATION":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "submodule" or col == "SUBMODULE" or col == "SUBMODEL":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "BAU/CR":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "theme objective":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "entity":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "Theme_ID" or col == "THEME ID" or col == "THEME_ID":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "Unique ID" or col == "UNIQUE_ID":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "ENTITY":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "EVENT":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "PLACEHOLDER":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())
            elif col == "UNIQUE_REF":
                l3_dict[col].append(old_dataframe_row[col].values[0].strip())

            if col in take_this_columns:
                # print("Old val-", old_dataframe_row[col].values[0].strip())
                # if len(old_dataframe_row[col].values[0].strip()) == 0:
                #     print("EMPTY...")
                # print("New val-", new_dataframe_row[col].values[0].strip())
                # if len(new_dataframe_row[col].values[0].strip()) == 0:
                #     print("EMPTY...")

                expected_str = "" if pd.isna(old_dataframe_row[col].values[0]) else str(
                    old_dataframe_row[col].values[0].strip())
                actual_str = "" if pd.isna(new_dataframe_row[col].values[0]) else str(
                    new_dataframe_row[col].values[0].strip())
                # print("col", col)
                # print("expected_str", expected_str)
                # print("actual_str", actual_str)
                #
                # if (len(old_dataframe_row[col].values[0].strip()) == 0 and len(
                #         new_dataframe_row[col].values[0].strip()) == 0) or \
                #         (str(old_dataframe_row[col].values[0].strip()) == str(
                #             new_dataframe_row[col].values[0].strip())):
                if expected_str == actual_str:
                    pass

                else:
                    remark.append(col)

            if len(remark) == 0:
                l3_dict["Status"].append("Pass")
                l3_dict["Remark"].append("SUCCESS")
            else:
                l3_dict["Status"].append("Fail")
                l3_dict["Remark"].append(", ".join(remark) + " Columns not matching.")

            l3_dataframe = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in l3_dict.items()]))

            # # COPY L3 DATAFRAME TO ELK DATAFRAME
            # elk_df = l3_dataframe
            # try:
            #     # GENERATING INDEX
            #     index_name = get_index(index_name_elK)
            #     print(f"INDEX IS {index_name}")
            #
            #     # CONNECTING WITH ELK SERVER
            #     elastic_search = ElasticSearchConnection()
            #     es = elastic_search.get_elastic_search_connection(index=index_name)
            #
            #     # CONVERTING ELK DATAFRAME TO DICTIONARY
            #     dic = util.excel_to_dict_with_unique_id(elk_df)
            #
            #     # INSERTING DICTIONARY TO ELK USING BULK FUNCTION
            #     insert_res = helpers.bulk(es, dic, index=index_name, request_timeout=200)
            #     print(f"BULK INSERT RESPONSE IS:{insert_res}")
            #     print(f"len({len(elk_df)}) records inserted and response is {insert_res} for index {index_name}")
            #     print("DATA INGESTED TO ELK SUCCESSFULLY-", datetime.now())
            #
            # except Exception as e:
            #     print("....DATA INGESTION EXCEPTION.....")
            #     print(e)

            l3_dataframe.insert(loc=0, column='Sr_No.', value=range(1, 1 + len(l3_dataframe)))
            return l3_dataframe

def create_l2_report(found, old_dataframe, new_dataframe, unique_col):
    ''' Make L2 Report and return '''
    l2_dict = defaultdict(list)

    for found_key in list(found[unique_col]):
        old_dataframe_row = old_dataframe[old_dataframe[unique_col] == found_key]
        new_dataframe_row = new_dataframe[new_dataframe[unique_col] == found_key]
        # comon_columns = list(old_dataframe.columns) and list(new_dataframes.columns)
        comon_columns = [value for value in list(old_dataframe.columns) if value in list(new_dataframe.columns)]

        for col in comon_columns:
            l2_dict["Column_Name"].append(col)
            l2_dict["Old_Prod"].append(old_dataframe_row[col].values[0].strip())
            l2_dict["New_Prod"].append(new_dataframe_row[col].values[0].strip())
            #  print("col", col, "--", old_dataframe[col].values[0].strip(),"-len", len(old_dataframe_row[col].values[0].strip()))
            #  print("col", col, "--", new_dataframe[col].values[0].strip(),"-len", len(new_dataframe_row[col].values[0].strip()))
            if (len(old_dataframe_row[col].values[0].strip()) == 0 and len(
                    new_dataframe_row[col].values[0].strip()) == 0) or \
                    (old_dataframe_row[col].values[0].strip() == new_dataframe_row[col].values[0].strip()):
                l2_dict["Out_Come"].append("True")
            else:
                l2_dict["Out_Come"].append("False")

    l2_dataframe = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in l2_dict.items()]))
    return l2_dataframe

def create_excel_file(my_dict_df, output_report_file_path, output_report_sheet_name):
    """"Make report"""
    if output_report_file_path.split(".")[-1].upper() == "CSV":
        # Make csv file
        my_dict_df.to_csv(output_report_file_path, index=False, na_rep='NA')
    elif output_report_file_path.split(".")[-1].upper() == "XLSX":
        # make excel file
        sheet_data_list = []
        sheet_name_list = []
        file = ''
        while True:
            try:
                file = pd.ExcelFile(output_report_file_path)
                break
            except:
                dataframe = pd.DataFrame()
                writer = pd.ExcelWriter(output_report_file_path)
                dataframe.to_excel(writer, sheet_name=output_report_sheet_name, index=False)
                writer.close()
        if file:
            for sheet in file.sheet_names:
                file_data = pd.read_excel(file, sheet_name=sheet)
                sheet_name_list.append(sheet)
                if sheet == output_report_sheet_name:
                    print("sheet", sheet)
                    frames = [file_data, my_dict_df]
                    new_file_df = pd.concat(frames)
                    sheet_data_list.append(new_file_df)
                else:
                    sheet_data_list.append(file_data)
        if not output_report_sheet_name in file.sheet_names:
            sheet_name_list.append(output_report_sheet_name)
            sheet_data_list.append(my_dict_df)
        writer = pd.ExcelWriter(output_report_file_path)
        for i in range(len(sheet_data_list)):
            sheet_data_list[i].to_excel(writer, sheet_name=sheet_name_list[i], index=False)
        writer.close()
        print("File Created..........")


# call validation function
# validation(input_old_file_path = "theme_file_path(expected_file_path)",
#             input_old_sheet_name="theme_file_init_sheet_name(expected_file_sheet_name)",
#            input_old_filter_column_name=None,
# input_new_file_path = "actual_file_path",
# input_new_sheet_name = "actual_file_sheet_name",
# input_new_filter_column_name=None,
# column_to_compare = 'add column name which are going to compare'
#                     'eg.UCIC,BXP',
# output_l1_report_file_path=None,
# output_l1_report_sheet_name=None,
# output_l2_report_file_path=None,
# output_l2_report_sheet_name=None,
# output_l3_report_file_path='validation_file_path which you to give',
# output_l3_report_sheet_name='validation sheet name which you to give',
# unique_col='UNIQUE_ID',
#            separator=None,
#            filter_file_path = None,
#            filter_sheet_name=None,
#            filter_column_name=None
# )




