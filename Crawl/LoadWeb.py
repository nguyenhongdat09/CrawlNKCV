import threading
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
import threading

class loadWeb():
    def __init__(self):
        self.url = ct.url
        self.url_nkcv = ct.url_nkcv
        self.url_qlyc = ct.url_qlyc
    def init_driver(self):
        options = Options()
        options.add_experimental_option("detach", True)  # khong cho tat chorme
        service_object = Service(binary_path)
        options.add_argument("--start-maximized")  # Thêm tùy chọn để mở cửa sổ lớn nhất

        driver = webdriver.Chrome(options=options, service=service_object)
        wait = WebDriverWait(driver, 10)
        return driver, wait

    def login(self):
        driver, wait = self.init_driver()
        driver.get(self.url)
        UserName = driver.find_element(By.ID, "LoginExtender_txtUserName")
        UserName.send_keys(ct.user_name)
        password = driver.find_element(By.ID, "LoginExtender_txtPassword")
        password.send_keys(ct.password)
        button_login = driver.find_element(By.ID, "LoginExtender_Ok")
        button_login.click()
        wait.until(EC.presence_of_element_located((By.ID, "LoginExtender_Attention")) )
        WebDriverWait(driver.find_element(By.ID, "LoginExtender_Attention"), 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "Attention"))
        )
        js_script = "$find('LoginExtender')._login(true);"
        driver.execute_script(js_script)
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "MenuExtenderGroup")))
        # Lấy danh sách cookie từ tab thứ nhất
        cookies = driver.get_cookies()
        return driver, wait, cookies
    def enter_info_qlyc(self, driver, wait):
        driver.get(self.url_nkcv)

        wait.until(EC.presence_of_element_located((By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_ma_ltql")))

        input_ltql = driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_ma_ltql")
        input_ltql.send_keys(ct.user_name)

        input_trang_thai = driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_trang_thai")
        input_trang_thai.send_keys(ct.trang_thai)

        input_ngay_ht_tu = driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_ngay_ht_tu")
        driver.execute_script("arguments[0].value = arguments[1];", input_ngay_ht_tu, ct.tu_ngay)

        input_ngay_ht_den = driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_form_ngay_ht_den")
        driver.execute_script("arguments[0].value = arguments[1];", input_ngay_ht_den, ct.den_ngay)

        button_ok = driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_updateDlgOk")
        button_ok.click()
        self.crawl_info_qlyc(driver, wait)
    def crawl_info_qlyc(self, driver, wait):
        wait.until(EC.presence_of_element_located((By.ID, "ctl00_FastBusiness_MainReport_ToolbarButton_JobLog")))
        button_nkcv_hidden = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_ToolbarButton_JobLog')
        button_nkcv_hidden.click()
        scripts = (
            "var g = $find('ctl00_FastBusiness_MainReport');"
            "g.set_gridPageSize(250);"
        )
        page_count = driver.execute_script(scripts)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'td#ctl00_FastBusiness_MainReport_Wait a.PaddedLink')))
        scripts = (
            "var g = $find('ctl00_FastBusiness_MainReport');"
            "return g._gridPageCount "
        )
        page_count = driver.execute_script(scripts)
        result = []
        print(page_count)
        for i in range(0, page_count):
            # ctl00_FastBusiness_MainReport_Wait
            scripts = (
                "var g = $find('ctl00_FastBusiness_MainReport');"
                f"g.goToPage('{i}'); "
            )

            if i != 0: driver.execute_script(scripts)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'td#ctl00_FastBusiness_MainReport_Wait a.PaddedLink')))

            scripts = (
                "var g = $find('ctl00_FastBusiness_MainReport');"
                "var stt_rec_list = [], noi_dung_list = [], ngay_ht_list = [];"
                "for (var i = 1; i <= g._rowCount; i++) {"
                "    stt_rec_list.push(g._getItemValue(i, g._getColumnOrder('stt_rec')));"
                "    noi_dung_list.push(g._getItemValue(i, g._getColumnOrder('noi_dung')));"
                "    ngay_ht_list.push(g._getItem(i, g._getColumnOrder('ngay_ht')).value);"
                "}"
                "return [stt_rec_list, noi_dung_list, ngay_ht_list];"
            )
            result.append(driver.execute_script(scripts))
        df = pd.DataFrame(columns=["stt_rec", "noi_dung", "ngay_ht"])
        for item in result:
            dict = {
                "stt_rec": item[0],
                "noi_dung": item[1],
                "ngay_ht": item[2],
            }
            # Tạo DataFrame tạm thời từ dict data
            temp_df = pd.DataFrame(dict)
            df = pd.concat([df, temp_df], ignore_index=True)
        df['so_gio'] = None
        df['diem'] = None
        print(df)
    def enter_info_nkcv(self, driver, wait):
        driver.get(self.url_qlyc)

        wait.until(EC.presence_of_element_located((By.ID, "ctl00_FastBusiness_MainReport_searchExtender_updateDlgPanel")))
        input_ngay_tu = driver.find_element(By.ID,
                                                    "ctl00_FastBusiness_MainReport_searchExtender_form_ngay_tu")
        driver.execute_script("arguments[0].value = arguments[1];", input_ngay_tu, ct.tu_ngay)

        input_ngay_den = driver.find_element(By.ID,
                                                     "ctl00_FastBusiness_MainReport_searchExtender_form_ngay_den")
        driver.execute_script("arguments[0].value = arguments[1];", input_ngay_den, ct.den_ngay)
        button_ok = driver.find_element(By.ID, "ctl00_FastBusiness_MainReport_searchExtender_updateDlgOk")
        button_ok.click()
    def run(self):
        driver, wait, cookies = self.login()
        self.enter_info_qlyc(driver, wait)
        self.enter_info_nkcv(driver, wait)



crawl = loadWeb()
crawl.run()