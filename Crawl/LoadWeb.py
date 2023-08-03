from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from chromedriver_py import binary_path
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Constant.Constant import Constant as ct
from selenium.common.exceptions import JavascriptException, NoSuchElementException
import pandas as pd
import time
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.expand_frame_repr', False)  # Ngăn các cột bị cắt và xuống dòng

# Khởi tạo trình duyệt (ở đây sử dụng Chrome)

# Mở trang web
class loadWeb():
    def __init__(self):
        self.url = ct.url
        self.url_nkcv = ct.url_nkcv
        options = Options()
        options.add_experimental_option("detach", True)  # khong cho tat chorme
        service_object = Service(binary_path)
        self.driver = webdriver.Chrome(options=options, service=service_object)
        self.wait = WebDriverWait(self.driver, 10)
    def login(self):
        self.driver.get(self.url)
        UserName = self.driver.find_element(By.ID, "LoginExtender_txtUserName")
        UserName.send_keys(ct.user_name)
        password = self.driver.find_element(By.ID, "LoginExtender_txtPassword")
        password.send_keys(ct.password)
        button_login = self.driver.find_element(By.ID, "LoginExtender_Ok")
        button_login.click()
        try:
            self.driver.find_element(By.ID, "LoginExtender_Attention").find_element(By.CLASS_NAME, "Attention")
            js_script = "$find('LoginExtender')._login(true);"
            self.driver.execute_script(js_script)
        except NoSuchElementException:
            pass
    def enter_info_nkcv(self):
        attention_div = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "MenuExtenderGroup")))
        self.driver.get(self.url_nkcv)
        self.wait.until(EC.presence_of_element_located((By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_ma_ltql")))

        input_ltql = self.driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_ma_ltql")
        input_ltql.send_keys(ct.user_name)
        input_trang_thai = self.driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_trang_thai")

        input_trang_thai.send_keys('HT,OK,UP')
        input_ngay_ht_tu = self.driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_ngay_ht_tu")

        self.driver.execute_script("arguments[0].value = arguments[1];", input_ngay_ht_tu, '01/07/2023')
        input_ngay_ht_den = self.driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_ngay_ht_den")
        self.driver.execute_script("arguments[0].value = arguments[1];", input_ngay_ht_den, '31/08/2023')

        button_ok = self.driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_updateDlgOk")
        button_ok.click()
    def crawl_info_nkcv(self):
        self.wait.until(EC.presence_of_element_located((By.ID, "ctl00_FastBusiness_MainReport_ToolbarButton_JobLog")))
        button_nkcv_hidden = self.driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_ToolbarButton_JobLog')
        button_nkcv_hidden.click()
        scripts = (
            "var g = $find('ctl00_FastBusiness_MainReport');"
            "g.set_gridPageSize(250);"
            "var stt_rec_list = [], noi_dung_list = [], ngay_ht_list = [];"
            "for (var i = 1; i <= g._rowCount; i++) {"
            "    stt_rec_list.push(g._getItemValue(i, g._getColumnOrder('stt_rec')));"
            "    noi_dung_list.push(g._getItemValue(i, g._getColumnOrder('noi_dung')));"
            "    ngay_ht_list.push(g._getItemValue(i, g._getColumnOrder('ngay_ht')));"
            "}"
            "return [stt_rec_list, noi_dung_list, ngay_ht_list];"
        )
        print(scripts)
        result = self.driver.execute_script(scripts)
        dict = {
            "stt_rec": result[0],
            "noi_dung": result[1],
            "ngay_ht": result[2],
        }
        df = pd.DataFrame(dict)
        print(df)
    def run(self):
        self.login()
        self.enter_info_nkcv()
        self.crawl_info_nkcv()
crawl = loadWeb()
crawl.run()