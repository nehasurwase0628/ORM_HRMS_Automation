from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common import keys

from Utils.paths import Paths


class WebDriverFactory():

    def __init__(self, browser,logger, driver, screenshot_path, flag):
        self.browser = browser
        self.driver = driver
        self.path = Paths(logger, driver, screenshot_path, flag)


    def get_webdriver_instance(self):
        """ Get webdriver instance based on the browser
        Returns:
            'Wbedriver Instance'"""
        try:
            logger = self.logger
        except:
            logger = ""  # foe thanos

        if self.browser == "iexplorer":
            # set dirver
            driver = webdriver.Ie()
        elif self.browser == "firefox":
            driver = webdriver.Firefox()
        elif self.browser == "chrome":
            driver = None
            try:
                # Set chrome driver
                """This path is to execute in jenkins"""
                driver_path = '/uer/bin/chromedriver'
                chrome_options = Options()
                chrome_options.add_argument('--no-sandbox')
                chrome_options.add_argument('headless')
                chrome_options.add_argument('--disable-blink-feature=AutomationControlled')
                chrome_options.add_argument('--disable-web-security')
                chrome_options.add_argument('--disable-extensions')
                chrome_options.add_argument('--ignore-certificate-errors')
                chrome_options.add_argument('--disable-dev-shm-usage')
                chrome_options.add_argument('--disable-features=VizDisplayCompositor')
                chrome_options.add_argument('--force-device-scale-factor=1')
                chrome_options.add_argument('--window-size=1680,1020')
                chrome_options.add_argument('force-device-scale-factor=0.8')
                chrome_options.add_argument('high-dpi-support=0.8')
                chrome_options.add_argument('--remote-allow-origins=* ')
                chrome_options.add_argument("--incognito")
                service = Service(driver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
            except:
                """This oTH IS TO EXECUTE IN LOCAL MACHINE"""
                driver_path = self.path.DRIVER_PATH
                chrome_options_local = Options()
                chrome_options_local.add_argument("--disable-dev-shm-usage")
                chrome_options_local.add_argument('--disable-gpu')
                chrome_options_local.add_argument('--enable-feature=NetworkServiceInProcess')
                chrome_options_local.add_argument('--disable-features=NetworkService')
                chrome_options_local.add_argument('--remote-allow-origins=*')
                chrome_options_local.add_argument("--incognito")
                service = Service(driver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options_local)
            if self.driver.service.is_connectable():
                print(f"=========== Chrome Driver Launched Successfully-----{self.driver}")
            else:
                print("======== Driver not invoked===========")

        else:
            driver =None
            try:
                # Set chrome driver
                """This path is to execute in jenkins"""
                driver_path = '/uer/bin/chromedriver'
                chrome_options = Options()
                chrome_options.add_argument('--remote-allow-origins=*')
                chrome_options.add_argument('--no-sandbox')
                chrome_options.add_argument('headless')
                chrome_options.add_argument('--disable-blink-feature=AutomationControlled')
                chrome_options.add_argument('--disable-web-security')
                chrome_options.add_argument('--disable-extensions')
                chrome_options.add_argument('--ignore-certificate-errors')
                chrome_options.add_argument('--disable-dev-shm-usage')
                chrome_options.add_argument('--disable-features=VizDisplayCompositor')
                chrome_options.add_argument('--force-device-scale-factor=1')
                chrome_options.add_argument('--window-size=1680,1020')
                chrome_options.add_argument('force-device-scale-factor=0.8')
                chrome_options.add_argument('high-dpi-support=0.8')
                chrome_options.add_argument("--incognito")
                service = Service(driver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
            except:
                """This oTH IS TO EXECUTE IN LOCAL MACHINE"""
                driver_path = self.path.DRIVER_PATH
                chrome_options_local = Options()
                chrome_options.add_argument('--remote-allow-origins=*')
                chrome_options_local.add_argument("--disable-dev-shm-usage")
                chrome_options_local.add_argument('--disable-gpu')
                chrome_options_local.add_argument('--enable-feature=NetworkServiceInProcess')
                chrome_options_local.add_argument('--disable-features=NetworkService')
                chrome_options_local.add_argument('--remote-allow-origins=*')
                chrome_options_local.add_argument("--incognito")
                service = Service(driver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options_local)
            if self.driver.service.is_connectable():
                print(f"=========== Chrome Driver Launched Successfully-----{self.driver}")
            else:
                print("======== Driver not invoked===========")
        # setting Driver Implicit time out for an element
        driver.implicitly_wait(20)
        return driver

    def setup_driver(self):
        driver = webdriver.Chrome()
        return driver

    def clear_cookie_cache(self):
        self.service= Service(self.path.DRIVER_PATH)
        self,options = Options()
        self.options.add_argument("--start-maximize") # maximise window
        self.driver = webdriver.Chrome(service=self.service, options=self.options)
        self.driver.delete_all_cookies()
        self.driver.get("chrome://settings/clearBrowserData") # clear cookies
        self.driver.find_element(By.XPATH, "//settings-ui").send_keys(keys.ENTER) #clear cache