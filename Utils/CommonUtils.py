import datetime
import os.path
import re
import time
from pathlib import Path
from telnetlib import EC

import openpyxl
from docx import Document
from docx.shared import Inches
import numpy as np
import pandas as pd
from Tools.scripts.patchcheck import status
from selenium.common import TimeoutException, StaleElementReferenceException, InvalidElementStateException, \
    ElementClickInterceptedException
from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

curr_file_path = Path(__file__)
root_dir = curr_file_path.parent.parent.absolute()
# from Utils.logger import root_dir


class CommonUtils():

    def __init__(self, logger, driver, screenshot_path, flag):
        self.logger = logger
        self.driver = driver
        self.screenshot_path = screenshot_path
        self.thanos_flag = flag
        self.wait = WebDriverWait(driver, 10)

    def create_update_excel_row_wise(self, file_path_n_name, flag, my_dict_df, rewrite_df_flag=None):
        # lock_file_path = file_path_n_name + ".lock"
        # lock = FileLock(lock_file_path, timeout=60)
        # file_path_n_name = Validation_file_path
        # flag = Sheet name or actual_sheet_name eg.EP_CMS_init
        # my_dict_df =  expected_df or actual_df
        try:
            # with lock
            if True:
                print("Lock acquired, proceeding with the file operations")
                sheet_data_list = []
                sheet_name_list = []
                file = ''

                while True:
                    try:
                        file = pd.ExcelFile(file_path_n_name)
                        break
                    except:
                        dataframe = pd.DataFrame()
                        writer = pd.ExcelFile(file_path_n_name)
                        if pd.isna(flag):
                            dataframe.to_excel(writer, sheet_name="Sheet1", index=False)
                        else:
                            dataframe.to_excel(writer, sheet_name=flag, index=False)
                        writer.close()

                for sheet in file.sheet_names:
                    file_data = pd.read_excel(file, sheet_name=sheet)
                    sheet_name_list.append(sheet)
                    if sheet == flag:
                        print("sheet", sheet)
                        if not rewrite_df_flag:
                            frames = [file_data, my_dict_df]
                            new_file_df = pd.concat(frames)
                        else:
                            new_file_df = my_dict_df
                        sheet_data_list.append(new_file_df)
                    else:
                        sheet_data_list.append(file_data)

                if not flag in file.sheet_names:
                    sheet_name_list.append(flag)
                    sheet_data_list.append(my_dict_df)

                writer = pd.ExcelWriter(file_path_n_name)
                for i in range(len(sheet_data_list)):
                    sheet_data_list[i].to_excel(writer, sheet_name=sheet_name_list[i], index=False)
                writer.close()
                print("File updated Successfully")
        except Exception as e:
            print(f"Exception in create_update_excel_row_wise line no :{e.__traceback__.tb_lineno}:{type(e).__name__}")

    def compare_dict(self, dict1, dict2, ket_list, logger=None):
        try:
            # dict1 is actual, dict2 is expected
            status_flag = True
            reason_string = ""
            value_list = []
            for key, value in dict2.items():
                if key in ket_list:
                    if key in dict1:
                        actual_val = dict1[key]
                        expected_val = dict2[key]

                        # handle Nan and empty values properly
                        if pd.isna(actual_val) and pd.isna(expected_val):
                            print(f"Match -> Key: {key}, Both values are NaN")
                            if logger:
                                logger.info(f"Match -> Key: {key}, Both values are NaN")
                            continue

                        # Convert everything to string for uniform comparison
                        actual_str = "" if pd.isna(actual_val) else str(actual_val).strip()
                        expected_str = "" if pd.isna(expected_val) else str(expected_val).strip()

                        if actual_str == expected_str:
                            print(f"Match -> Key: {key}, Actual:{actual_str}, Expected:{expected_str}")
                            if logger:
                                logger.info(f"Match -> Key: {key}, Actual:{actual_str}, Expected:{expected_str}")
                            continue
                        else:
                            status_flag = False
                            reason_string+=key + ", "
                            value_list.append("_".join([expected_str, actual_str]))
                            print(f"Mismatch -> Key: {key}, Actual:{actual_str}, Expected:{expected_str}")
                            if logger:
                                logger.info(f"Mismatch -> Key: {key}, Actual:{actual_str}, Expected:{expected_str}")
                    else:
                        status_flag = False
                        reason_string+=key + "/key missing in actual" + ", "
                        value_list.append(key + "/key missing in actual" + ", ")
                        print(f"key missing ->{key}")
                        if logger:
                            logger.error(f"key missing->{key}")
                else:
                    print(f"Ignoring {key} for comparison")
                    if logger:
                        logger.info(f"Ignoring {key} for comparison")
            value_str = ",".join(value_list) if value_list else ""
            status_flag = "PASS" if status_flag else "FAIL"
            return status_flag,reason_string,value_str
        except Exception as e:
            print(f"compare_dict Exception Line no:{e.__traceback__.tb_lineno}:{type(e).__name__}")
            logger.error(f"compare_dict Exception Line no:{e.__traceback__.tb_lineno}:{type(e).__name__}")

    def convert_float_to_int_if_no_decimal(self, value):
        # if ".0" in str(value):
        print("Before value in convert_float_to_int_if_no_decimal", value)
        if str(value).endswith(".0"):
            value = float(value)
            print(type(value))
            if isinstance(value, (np.float64, float)):
                value = int(value)
            else:
                value = str(value)
        else:
            value = str(value)
        print("AFTER value in convert_float_to_int_if_no_decimal", value)
        return  value

    def validate_single_record(self, actual_dict, expected_dict, valid_path, exp_sheet_name, actual_sheet_name, flag, compare_key_list):
        status_flag = ""
        reason_string = ""
        value_str = ""
        try:
            print(f"validate_single_record {exp_sheet_name}, {actual_sheet_name}")
            self.logger.info(f"validate_single_record {exp_sheet_name}, {actual_sheet_name}")
            self.logger.info(f"actual_dict, {actual_sheet_name}")

            actual_dict_df = pd.DataFrame(actual_dict, columns=list(actual_dict.keys()), index=[0])
            self.logger.info(f"actual_dict_df, {actual_dict_df}")
            # valid_path = validation file path

            self.create_update_excel_row_wise(valid_path, actual_sheet_name, actual_dict_df)

            validation_df = {}
            status_flag, reason_string, value_str = self.compare_dict(actual_dict, expected_dict, compare_key_list)
            print(f"status_flag {status_flag}")
            print(f"reason_string {reason_string}")
            print(f"value_str {value_str}")
            validation_dict={
                "THEME_ID": expected_dict["ThemE_ID"],
                "UNIQUE_ID": expected_dict["UNIQUE_ID"],
                "ENTITY": expected_dict["ENTITY"],
                "VALIDATION_RESULT": status_flag,
                "REASON": f"{reason_string}-!-!-{value_str}"
            }
            validation_df = pd.DataFrame(validation_dict, index=[0])
            self.create_update_excel_row_wise(valid_path, "FINAL_RESULT", validation_df)

        except Exception as e:
            self.logger.error(f"validate_single_record Exception == Line no:{e.__traceback__.tb_lineno}:{type(e).__name__}")
            print(f"validate_single_record Exception == Line no:{e.__traceback__.tb_lineno}:{type(e).__name__}")

    def Take_screenshot(self, unique_id, flag=None, paragraph=None):
        # senf flag as folder name eg. "EP/
        flag = flag.replace("/", "") # it removes /
        today_date = datetime.date.today().strftime("%d%m")
        curr_exec_time = datetime.datetime.now().strftime("%H%M")
        try:
            time.sleep(1)
            if not self.thanos_flag:
                if flag:
                    screenshot_filename = "image.png"
                    image_path = f"{root_dir}/Screenshot/{flag}/Image/{screenshot_filename}"
                else:
                    screenshot_filename = "image.png"
                    image_path = f"{root_dir}/Screenshot/Image/{screenshot_filename}"
            else:
                screenshot_filename = f"{unique_id}_image.png"
                image_path = os.path.join(self.screenshot_path, "Image", screenshot_filename)

            img_dir_path = os.path.dirname(image_path)
            if not os.path.exists(img_dir_path):
                os.makedirs(img_dir_path)

            self.driver.save_screenshot(image_path)

            if not self.thanos_flag:
                if flag:
                    docx_path = f"{root_dir}/Screenshot/{flag}/{flag}_{unique_id}.docx"
                    print("docx_path", docx_path)
                else:
                    docx_path = f"{root_dir}/Screenshot/COMMON_{unique_id}.docx"
                    print("docx_path", docx_path)
            else:
                docx_path = os.path.join(self.screenshot_path, f"{unique_id}.docx")

            dir_path = os.path.dirname(docx_path)
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)

            if not os.path.exists(docx_path):
                doc = Document()
                doc.add_heading(f"Screenshot for {unique_id}", 1)
            else:
                doc = Document(docx_path)

            if paragraph:
                doc.add_paragraph(f"{paragraph}")
            doc.add_picture(image_path, width=Inches(5))

            doc.save(docx_path)

            print("Screenshot saved and added to word document")
            print(f"Screenshot saved and added to word document {docx_path}")
            self.logger.info(f"Screenshot saved and added to word document {docx_path}")

        except Exception as e:
            self.logger.error(f"screenshot unsuccessful --> Line no.{e.__traceback__.tb_lineno}:{type(e).__name__}")
            print(e)
            print(f"screenshot unsuccessful --> Line no.{e.__traceback__.tb_lineno}:{type(e).__name__}")
            if type(e).__name__ == "WebDriverException":
                raise Exception(f"{type(e).__name__} ---- {repr(e)}")

    def take_screenshot_visual_regression(self, path, image_name):
        image_path = "../TestData/Screenshot/" + path + image_name
        if os.path.exists(image_path):
            os.remove(image_path)
            print("image file cleared successfully")
        self.driver.save_screenshot(image_path)

    def visual_regression_screenshot(self, screenshot_flag, screenshot_folder, ss_file_path, ss_flow_name, pagename):
        if screenshot_flag == "baseline":
            time.sleep(1)
            self.take_screenshot_visual_regression(f"BaseLineImage/{screenshot_folder}", f"{screenshot_folder[:-1]}{pagename}.png")

        elif screenshot_flag == 'currentImage':
            time.sleep(1)
            self.take_screenshot_visual_regression(f"currentImage/{screenshot_folder}",
                                                   f"{screenshot_folder[:-1]}{pagename}.png")
            ss_file_path.append(f"{screenshot_folder[:-1]}{pagename}.png")
            ss_flow_name.append(f"{screenshot_folder[:-1]}{pagename}.png")

        return ss_file_path, ss_flow_name

    def waitForVisibilityOfElement(self, xpath):
        try:
            wait = WebDriverWait(self.driver, timeout=2, ignored_exceptions=[TimeoutException,
                                                                             StaleElementReferenceException])
            wait.until(expected_conditions.visibility_of_element_located((By.XPATH, xpath)))
            print("Element is found wth xpath", xpath)
            return True
        except:
            print("Element is not found wth xpath", xpath)
            return False

    def load_workbook(self, wb_path):
        if os.path.exists(wb_path):
            return openpyxl.load_workbook((wb_path))
        return openpyxl.Workbook()

    def create_excel_file(self, file_path_n_name, flag, my_dict_df):
        sheet_data_list = []
        sheet_name_list = []
        file = ''
        while True:
            try:
                file = pd.ExcelFile(file_path_n_name)
                break
            except:
                dataframe = pd.DataFrame()
                writer = pd.ExcelWriter(file_path_n_name)
                if pd.isna(flag):
                    dataframe.to_excel(writer, sheet_name="Sheet1", index=False)
                else:
                    dataframe.to_excel(writer, sheet_name=flag, index=False)
                writer.close()

        for sheet in file.sheet_names:
            file_data = pd.read_excel(file, sheet_name=sheet)
            sheet_name_list.append(sheet)
            if sheet == flag:
                print("sheet", sheet)
                frames = [file_data, my_dict_df]
                new_file_df = pd.concat(frames)
                sheet_data_list.append(new_file_df)
            else:
                sheet_data_list.append(file_data)
        if not flag in file.sheet_names:
            sheet_name_list.append(flag)
            sheet_data_list.append(my_dict_df)
        writer = pd.ExcelWriter(file_path_n_name)
        for i in range(len(sheet_data_list)):
            sheet_data_list[i].to_excel(writer, sheet_name=sheet_name_list[i], index=False)
        writer.close()
        print("File Created..........")

    def converttostring(self, input_list):
        result = ''
        for item in input_list:
            result += str(item) + ';'
        result = result[:-1]
        return result

    def input_field(self, xpath, value):
        try:
            input = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            input.clear()
            input.send_keys(value)
        except:
            print("in except block")
            raise InvalidElementStateException

    def click_button(self, xpath):
        try:
            if self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath))):
                self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
        except ElementClickInterceptedException as e:
            print(f"click button exception line no: {e.__traceback__}:{type(e).__name__}")

    def find_element(self, xpath):
        try:
            if self.wait.until(EC.presence_of_element_located((By.XPATH, xpath))):
                element = self.driver.find_element(By.XPATH, xpath)
                return  element
            else:
                return False
        except:
            print(xpath,"....exception")

    def new_click_button(self, xpath):
        while True:
            if self.waitForVisibilityOfElement(xpath):
                try:
                    if self.waitForVisibilityOfElement(xpath):
                        time.sleep(1)
                        element = self.find_element(xpath)
                        if element:
                            self.click_button(xpath)
                            break
                        else:
                            if self.waitForVisibilityOfElement(xpath):
                                time.sleep(1)
                                element = self.find_element(xpath)
                                if element:
                                    self.click_button(xpath)
                                    break
                except:
                    print("new_click_button exception...")
            else:
                try:
                    if self.waitForVisibilityOfElement(xpath):
                        time.sleep(1)
                        element = self.find_element(xpath)
                        if element:
                            self.click_button(xpath)
                            break
                except:
                    print("new_click_button exception...")

    def find_element_5(self, xpath):
        for _ in range(5):
            try:
                if self.wait.until(EC.presence_of_element_located((By.XPATH, xpath))):
                    element = self.driver.find_element(By.XPATH, xpath)
                    return  element
                else:
                    return False
            except:
                print(xpath,"....exception")

    def clear_xpath(self, xpath):
        while True:
            try:
                element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                if element:
                    self.driver.find_element(By.XPATH, xpath).clear()
                    break
            except:
                print("clear exception", xpath)

    def click_xpath(self, xpath):
        while True:
            try:
                if self.wait.until(EC.presence_of_element_located((By.XPATH, xpath))):
                    element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                    if element:
                        self.driver.find_element(By.XPATH, xpath).clear()
                        break
            except:
                print("clear exception", xpath)

    def send_keys_xpath(self, xpath, text):
        while True:
            try:
                element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                if element:
                    self.driver.find_element(By.XPATH, xpath).send_keys(text)
                    break
            except:
                print("clear exception", xpath)

    def single_get_and_find_element_xpath(self, xpath):
        while True:
            try:
                element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                if element:
                    element = self.driver.find_element(By.XPATH, xpath)
                    return element
            except:
                print("clear exception", xpath)

    def multiple_get_and_find_element_xpath(self, xpath):
        while True:
            try:
                element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                if element:
                    elements = self.driver.find_elements(By.XPATH, xpath)
                    print("length of element", len(elements))
                    return element
            except:
                print("clear exception", xpath)
#------------------------ Number code logic ------------
    def data_extract(self, data, input_list):
        # data_extrac(510, [2,4,8,16,32,64,128,256])
        print(data)
        regex = re.compile('[@_!#$%*&{}<>/\}{~:]')
        if regex.search(str(data)) == None:
            lst = []
            d = str(data)
            lst.append(d)
            for i in lst:
                li = int(i)
                d = self.Calresutl(input_list, li)
            return d
        else:
            d =data.split('&')
            for i in d:
                li = int(i)
                d =self.Calresutl(input_list, li)
            return d

    def Calresutl(self, lst,target):
        for r in range(1, len(lst+1)):
            for i in self.combinatination(lst,r):
                if sum(i) == target:
                    return i

    def combinatination(self, mylist,r):
        if r ==0:
            return [[]]
        l = []
        for i in range(0,len(mylist)):
            first = mylist[i]
            rem = mylist[i+1:]
            comlst = self.combinatination(rem, r -1)
            for x in comlst:
                l.append([first]+x)
        return l

#------------------------ Number code logic End ------------

    def click_tab(self):
        actions=ActionChains(self.driver)
        actions.send_keys(Keys.TAB)
        actions.perform()

    def click_arrow_down(self):
        actions=ActionChains(self.driver)
        actions.send_keys(Keys.ARROW_UP)
        actions.perform()

    def click_arrow_up(self):
        actions=ActionChains(self.driver)
        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()

    def click_arrow_left(self):
        actions=ActionChains(self.driver)
        actions.send_keys(Keys.ARROW_LEFT)
        actions.perform()

    def click_arrow_right(self):
        actions=ActionChains(self.driver)
        actions.send_keys(Keys.ARROW_RIGHT)
        actions.perform()

    def click_space(self):
        actions=ActionChains(self.driver)
        actions.send_keys(Keys.SPACE)
        actions.perform()

    def click_enter(self):
        actions=ActionChains(self.driver)
        actions.send_keys(Keys.ENTER)
        actions.perform()

    def scroll_additional_pixels(self, additional_pixel):
        # Get the current scroll position
        current_scroll_position = self.driver.execute_script("return window.pageYOffset;")

        #Calcuate the new scroll position
        new_scroll_position = current_scroll_position + additional_pixel

        # Use javascript to scroll to new position
        scroll_script = "window.scroll(0, {});".format(new_scroll_position)
        self.driver.execute_script(scroll_script)

    def find_element_by_presence(self, xpath):
        try:
            if self.wait.until(EC.presence_of_element_located((By.XPATH, xpath))):
                return True
            return False
        except:
            print("Exception occur while finding element", xpath)

    def has_special_character(self, input_str):
        special_chars = ['@','"', '(', '}', ']', '#','$','%','^','&','*','<','>','/', '{', '~', ':', ' ', '' ]
        flag = False
        string =str(input_str)
        for c in string:
            if c in special_chars:
                flag=True
                break
        return flag






