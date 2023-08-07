import pandas as pd
import selenium
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from chromedriver_py import binary_path
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Constant.Constant import Constant as ct
from selenium.common.exceptions import JavascriptException, NoSuchElementException, TimeoutException
from Crawl.LoadWeb import loadWeb
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from Crawl.ReadExcel import readExcel
from selenium.webdriver.common.keys import Keys
class entertoweb:
    def __init__(self):
        self.url_nkcv = ct.url_nkcv
        self.url = ct.url
    def enter_nkcv(self, driver: WebDriver, wait: WebDriverWait, df: pd.DataFrame):
        wait.until(EC.presence_of_element_located((By.ID, "ctl00_FastBusiness_MainReport_ToolbarButton_New")) )
        button_new = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_ToolbarButton_New')
        button_new.click()

        for index, row in df.iterrows():
            stt_rec, ngay, gio, diem, ma_diem, trang_thai, nh_cv1, nh_cv2, ma_nv = row['stt_rec'], row['ngay'], row['gio_pb_arr'], int(float(row['diem'])), row['ma_diem'], row['trang_thai'], row['nh_cv1'], row['nh_cv2'], row['ma_nv']
            input_ma_yc = wait.until(
                EC.presence_of_element_located((By.ID, "ctl00_FastBusiness_MainReport_dirExtender_form_stt_rec_yc")))
            print(stt_rec)
            input_ngay_ct = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_form_ngay_ct')
            input_ngay_ht = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_form_ngay_ht')
            input_ma_nv = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_form_ma_nhv')
            input_so_gio = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_form_so_gio')
            input_ma_diem = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_form_ma_plyc')
            input_diem = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_form_diem')
            input_nh_cv1 = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_form_nh_cv1')
            input_nh_cv2 = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_form_nh_cv2')
            input_ma_yc.send_keys(stt_rec)
            input_ma_yc.send_keys(Keys.ENTER)
            driver.execute_script("arguments[0].value = arguments[1];", input_ngay_ct, ngay)
            driver.execute_script("arguments[0].value = arguments[1];", input_ngay_ht, ngay)

            input_ma_diem.send_keys(Keys.ENTER)
            input_ma_diem.send_keys(ma_diem)
            input_ma_diem.send_keys(Keys.ENTER)
            input_so_gio.send_keys(gio)
            input_nh_cv1.send_keys(nh_cv1)
            input_nh_cv2.send_keys(nh_cv2)
            input_diem.clear()
            input_diem.send_keys(diem)
            button_save = driver.find_element(By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_updateDlgOk')
            button_save.click()
            button_re_new = wait.until(
                EC.visibility_of_element_located((By.ID, 'ctl00_FastBusiness_MainReport_dirExtender_updateDlgNew')))
            button_re_new.click()

    def run(self):

        reader = readExcel()
        df = reader.get_df_nkcv_finish()
        df['ngay'] = pd.to_datetime(df['ngay']).dt.strftime('%d/%m/%Y')
        print(df)
        confirm = input('Hãy xem lại bảng nhật ký công việc trên(1 - nhập đi, 0 - Không nhập): ' )
        while confirm not in ('0', '1'):
            confirm = input('Hãy xem lại bảng nhật ký công việc trên (1 - OK nhập đi, 0 - Không nhập): ')
        if confirm == '1':
            load = loadWeb()
            driver, wait, cookie = load.login()
            load.enter_info_nkcv(driver, wait)
            self.enter_nkcv(driver, wait, df)
        else:
            return None

# enter = entertoweb()
# enter.run()