from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import logging
from excel_input import excel_input
from valuate import value_test 
from valuate_1 import value_test_1
from valuate_45 import value_test_45
from excel_output import excel_output
from edit import edit
from edit_demo2 import edit_demo2
from edit_Earth import edit_Earth
from edit_ghg import edit_ghg
from delete import delete
from create_data_prod import source_input_prod
from create_data_demo2 import source_input_demo2
from create_data_Earth import source_input_Earth
from create_data_people_English import source_input_English
from upstream_emissions_verification import upstreamEmissions_add,upstreamEmissions_edit
from significance_edit import significance_edit
from salience_reset import salience_reset
from test import salience_calculate_test
from pass_response import excel_input_test
import os
import allure



logging.shutdown()
def login_page(driver,user_id,user_pass,site_num):
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.visibility_of_element_located((By.ID, "login_username")))   #輸入帳號
        element.send_keys(user_id)
        element = wait.until(EC.visibility_of_element_located((By.ID, "login_password")))   #輸入密碼
        element.send_keys(user_pass)
        time.sleep(5)
        
        checkbox = driver.find_element(By.CSS_SELECTOR, "input[type='checkbox']")  #勾選同意使用條款
        checkbox.click()
    
        button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))    #登入submit
        button.click()
        if site_num == 3:
            autotest_td = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//td[text()='亞O源']"))  # 要測試的盤查計畫名稱
            )
        elif site_num == 5:
            autotest_td = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//td[text()='A廠區測試']"))  # 要測試的盤查計畫名稱
            )
        elif site_num == 7:
            autotest_td = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//td[text()='顧問係數測試']"))  # 要測試的盤查計畫名稱
            )
        else:
            autotest_td = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//td[text()='測試盤查']"))   #要測試的盤查計畫名稱
        )
    
    # 找到父 <tr> 元素
        parent_tr = autotest_td.find_element(By.XPATH, './parent::tr')
    
        button = parent_tr.find_element(By.XPATH, './/button[@type="button" and contains(@class, "ant-btn ant-btn-round ant-btn-link ant-btn-icon-only text-primary  border-0")]')
    #JavaScript寫法
        driver.execute_script("arguments[0].click()", button)
        time.sleep(1)
        
        link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "排放資料輸入")))   #點擊排放資料輸入
        link.click()

def main():
    log_file_path = 'autotest.log'
    if os.path.exists(log_file_path):
        os.remove(log_file_path)
    logging.basicConfig(
        filename='autotest.log',  # 指定日誌輸出的檔案名稱
        level=logging.INFO,  # 設定日誌級別（這裡設定為 INFO 級別，可以自行調整）
        format='%(asctime)s - %(levelname)s - %(message)s',  # 設定日誌格式
        datefmt='%Y-%m-%d %H:%M:%S'  # 設定日期格式
    )

    # current_dir = os.path.dirname(os.path.abspath(__file__))
    # chromedriver_path = os.path.join(current_dir, 'chromedriver-win64', 'chromedriver.exe')
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    #options.add_argument('-headless')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument("--start-maximized")
    options.add_argument('--disable-web-security')
    options.add_argument('--ignore-certificate-errors')
    # service = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    print("================================================================\n")
    print("請問要測試的網站是 ??? 1.prod 2.demo2 3.地球醫生 4.雲科大 5.黑丸 6.888 7.999")
    print("請輸入 : \n")
    site_num = int(input())
    site_dict = {1: "https://prod.netzero.com.tw", 2: "https://demo2.netzero.com.tw",
                 3: "https://14064-1.com", 4: "https://ghg.netzeroyun.com",
                 5: "https://220.132.206.5:666", 6: "https://220.132.206.5:8888",
                 7: "https://220.132.206.5:9999", 8: "https://naturesbank.production.mjmtech.com.tw"}
    url = site_dict[site_num]


    driver.get(url)
    if site_num == 3:
        user_id = 'GHGINVCLASS'
        user_pass = '0277289341'
    elif site_num == 4 or site_num == 5:
        user_id = 'testCompany'
        user_pass = '@Test123'
    # elif site_num == 6:
    #     user_id = '51802691'
    #     user_pass = 'J62U6AU4A83'
    else:
        user_id = 'testCompany'
        user_pass = 'testCompany'


    login_page(driver,user_id,user_pass,site_num)   #自動登入至排放源輸入
    print("請問要自動測試何種功能? 1:刪除 2:新增(多個) 3:顯著性評分 4:總值驗證 5:編輯 6:Excel模板下載 7:Excel匯出 8:Excel匯入 9:新增(英文語系)要先轉換語系 ,10:一鍵測試(不包含多語系)")

    func = int(input())
    start_time = time.time()
    start_date = datetime.now()
    if func ==1 :
        print("刪除 自動化測試")
        logging.info('==================刪除 開始==================')
        delete(driver, url)  # 刪除
        time.sleep(1)
        logging.info('=================刪除 結束===================')
    elif func ==2 :
        print("新增(3筆) 自動化測試")
        logging.info('===============新增(3筆) 開始=================')
        # if site_num == 2 or site_num == 4:
        #     source_input_demo2(driver,url) #demo2新增
        # elif site_num == 3:
        #     source_input_Earth(driver,url) #地球醫生新增
        # else:
        source_input_prod(driver,url,site_num)  #其他新增
        time.sleep(1)
        logging.info('================新增(3筆) 結束================')
    elif func ==3 :
        print("顯著性評分 自動化測試")
        logging.info('================顯著性評分 開始================')
        significance_edit(driver,url, site_num)
        time.sleep(1)
        logging.info('================顯著性評分 結束================')
    elif func ==4 :
        print("總值計算與驗證")
        logging.info('================總值計算 開始==================')
        # value_test(driver, url,site_num)
        # value_test_1(driver,url,site_num)
        value_test_45(driver, url,site_num)
        time.sleep(1)
        logging.info('================總值計算 結束==================')
    elif func ==5 :
        print("編輯 自動化測試")
        logging.info('===================編輯 開始===================')
        if site_num == 2 :
            edit_demo2(driver,url)
        elif site_num == 4:
            edit_ghg(driver,url)
        elif site_num == 3:
            edit_Earth(driver,url)
        else:
            edit(driver, url)
        time.sleep(1)
        logging.info('==================編輯 結束=====================')
    elif func == 6 :
        print("excel 自動模板下載")
        logging.info('===============excel模板下載 開始================')
        temp_func = 'download'
        excel_output(driver,url,temp_func)
        logging.info('===============excel模板下載 結束================')
    elif func == 7 :
        print("excel 匯出")
        logging.info('================excel匯出 開始==================')
        temp_func = 'output'
        excel_output(driver,url,temp_func)
        logging.info('================excel匯出 結束==================')
    elif func == 8 :
        print("excel 匯入")
        logging.info('================excel匯入 開始==================')
        excel_input(driver,url)
        #value_test(driver, url, site_num)
        logging.info('================excel匯入 結束==================')
    else:
        print("刪除 自動化測試")
        logging.info('================刪除 開始=======================')
        delete(driver,url) #刪除
        time.sleep(1)
        logging.info('================刪除 結束=======================')

        print("新增(3筆) 自動化測試 \n")
        logging.info('================新增(3筆) 開始===================')
        # if site_num == 2 or site_num == 4:
        #     source_input_demo2(driver,url) #demo2新增
        # elif site_num == 3:
        #     source_input_Earth(driver,url) #地球醫生新增
        # else:
        source_input_prod(driver,url,site_num)  #prod與其他新增
        time.sleep(1)
        print("顯著性分析評分 \n")
        significance_edit(driver, url,site_num) #顯著性評分
        time.sleep(1)
        logging.info('================新增(3筆) 結束===================')

        print("新增後計算總值與驗證 \n")
        logging.info('================新增後總值計算 開始================')
        value_test(driver,url,site_num) #總值計算
        time.sleep(1)
        logging.info('================新增後總值計算 結束================')
        print("新增後的上游運輸驗證 \n")
        logging.info('=============新增後的上游能源驗證 開始==============')
        upstreamEmissions_add(driver, url, site_num) # 新增後的上游運輸驗證
        time.sleep(1)
        logging.info('=============新增後的上游能源驗證 結束==============')
        # salience_reset(driver,url) #顯著性重置
        print("編輯 自動化測試 \n")
        logging.info('==================編輯 開始======================')
        if site_num == 2 :
            edit_demo2(driver, url) #demo2編輯
        elif site_num == 4:
            edit_ghg(driver,url)
        elif site_num == 3:
            edit_Earth(driver, url) #地球醫生編輯
        else:
            edit(driver, url) #prod與其他編輯
        time.sleep(1)
        logging.info('==================編輯 結束======================')
        time.sleep(1)
        print("編輯後計算總值與驗證 \n")
        logging.info('================編輯後總值計算 開始================')
        value_test(driver,url,site_num) #編輯後總值驗證
        time.sleep(1)
        logging.info('================編輯後總值計算 結束================')
        print("編輯後的上游能源驗證 \n")
        logging.info('=============編輯後的上游能源驗證 開始==============')
        upstreamEmissions_edit(driver,url, site_num) #編輯後的上游能源驗證
        time.sleep(1)
        logging.info('=============編輯後的上游能源驗證 結束==============')
        print("excel 自動模板下載")
        logging.info('================Excel模板下載 開始================')
        temp_func = 'download'
        excel_output(driver,url,temp_func) #excel模板下載
        logging.info('================Excel模板下載 結束================')
        print("Excel 匯出")
        logging.info('================Excel匯出 開始===================')
        temp_func = 'output'
        excel_output(driver, url, temp_func) #excel匯出
        logging.info('================Excel匯出 結束===================')
        print("Excel 匯入")
        logging.info('=====================Excel匯入 開始==============')
        excel_input(driver, url)
        time.sleep(1)
        logging.info('==================Excel匯入 結束=================')
        print("匯入後總值驗證")
        logging.info('================匯入後總值驗證 開始================')
        value_test(driver,url,site_num) #匯入後總值驗證
        time.sleep(1)
        print("匯入後的上游能源驗證")
        upstreamEmissions_edit(driver,url, site_num) #匯入後的上游能源驗證
        time.sleep(1)
        logging.info('================匯入後總值驗證 結束================')



    #計算時間
    end_time = time.time()
    total_time = end_time - start_time
    hours = total_time // 3600
    minutes = (total_time % 3600) // 60
    seconds = total_time % 60
    print(f"測試日期: {start_date.strftime("%Y-%m-%d %H:%M")}，總耗時: {hours} 小時 {minutes} 分 {seconds:.2f} 秒")
main()
