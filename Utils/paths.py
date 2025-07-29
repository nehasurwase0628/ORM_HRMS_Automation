import os
import sys
from configparser import ConfigParser



class Paths():
    def __init__(self, logger, driver, screenshot_path, flag):
        self.logger = logger
        self.driver = driver
        self.screenshot_path = screenshot_path
        self.flag = flag

    print(os.path.dirname(__file__))
    print(os.pardir)
    files_location = os.path.dirname(__file__) + os.sep + os.pardir + 'Configuration' + os.sep
    print(files_location)
    filepath = files_location + 'ORMApplication.properties'
    print("File path for config", filepath)
    con = ConfigParser()
    con.read(filepath)

    DRIVER_PATH =  con.get('drivers', 'chrome_path')
    URL_PATH =  con.get('url', 'url')
    THEMEDASHOBARD_PATH =  con.get('excel_path', 'theme_dashboard_xlsx')

    PDF_PATH =  con.get('excel_path', 'pdf_file_path')
    TXT_PATH =  con.get('excel_path', 'txt_file_path')
    XLSX_PATH =  con.get('excel_path', 'xlsx_file_path')
    DOC_PATH =  con.get('excel_path', 'doc_file_path')
    DOCX_PATH =  con.get('excel_path', 'docx_file_path')
    XLS_PATH =  con.get('excel_path', 'xls_file_path')

    ACTUAL_FILE_PATH =  con.get('excel_path', 'orm_actual_file_path')

    VALIDATION_L1_PATH =  con.get('excel_path', 'validation_l1_xlsx')
    VALIDATION_L2_PATH =  con.get('excel_path', 'validation_l2_xlsx')
    VALIDATION_L3_PATH =  con.get('excel_path', 'validation_l3_xlsx')

    IMAGE_STATIC_FOLDER_PATH =  con.get('static_folder_path', 'image_folder_path')
    SCREENSHOT_FOLDER_PATH =  con.get('static_folder_path', 'screenshot_folder_path')
    LOGGER_PATH =  con.get('static_folder_path', 'logger')
    IMAGE_PROCESSING_OUTPUT_PATH =  con.get('static_folder_path', 'image_processing_output_folder')




