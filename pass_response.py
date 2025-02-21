from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import random
import logging
from response import catch_response
import os
import sys


def resource_path(relative_path):
    # 取得資源檔案的絕對路徑，無論是開發環境或打包後的環境
    try:
        # PyInstaller 建立臨時資料夾，並將路徑儲存在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def excel_input_test(driver, url, site_num):
    source_type = ['移動源 C1', '固定燃燒源 C1', '工業製程 C1', '人為逸散 C1', '其他關注類物質 C1', '輸入電力 C2',
                   '輸入蒸汽 C2', '客戶和訪客運輸 C3', '員工差旅 C3', '員工通勤 C3', '營運產生之廢棄物 C4',
                   '資本設備 C4', '輸入能源上游排放 C4', '用水(水資源管理) B.5.2.a', '購買商品 C4', '上游租賃資產 C4',
                   '顧問諮詢、清潔、維護等 C4', '上游運輸及配送 C3',
                   '下游運輸及配送 C3', '廢棄物運輸 C3', '下游租賃資產 C5', '產品壽命終止階段 C5', '產品使用階段 C5',
                   '投資 C5', '其它間接排放 C6']

    file_path = resource_path('正確資料')

    upload_path = ['Vehicle.xlsx', 'StationaryCombustionSource.xlsx', 'IndustrialProcess.xlsx', 'workhour.xlsx',
                   'RFG.xlsx', 'Fire.xlsx', 'other.xlsx', 'Elec.xlsx', 'steam_plus.xlsx', 'EmployeeCommuting.xlsx',
                   'EmployeeCommuting.xlsx', 'EmployeeCommuting.xlsx', 'Disposal waste.xlsx',
                   'LifecycleAssmt.xlsx', 'LifecycleAssmt.xlsx', 'LifecycleAssmt.xlsx', 'PurchasedProduct.xlsx',
                   'LifecycleAssmt.xlsx', 'LifecycleAssmt.xlsx', 'MaterialTransport.xlsx',
                   'MaterialTransport.xlsx', 'MaterialTransport.xlsx', 'LifecycleAssmt.xlsx', 'LifecycleAssmt.xlsx',
                   'LifecycleAssmt.xlsx', 'LifecycleAssmt.xlsx', 'LifecycleAssmt.xlsx']

    for_response = ['移動源 C1', '固定燃燒源 C1', '工業製程 C1', '人為逸散 C1(工時計算)', '人為逸散 C1(冷媒設備)',
                    '人為逸散 C1(消防設備)', '其他關注類物質 C1', '輸入電力 C2(一般用電)', '輸入蒸汽 C2(加項)',
                    '客戶和訪客運輸 C3', '員工差旅 C3', '員工通勤 C3', '營運產生之廢棄物 C4',
                    '資本設備 C4', '輸入能源上游排放 C4(輸入電力上游)', '用水(水資源管理) B.5.2.a', '購買商品 C4',
                    '上游租賃資產 C4', '顧問諮詢、清潔、維護等 C4', '上游運輸及配送 C3',
                    '下游運輸及配送 C3', '廢棄物運輸 C3', '下游租賃資產 C5', '產品壽命終止階段 C5', '產品使用階段 C5',
                    '投資 C5', '其它間接排放 C6']
    # for_response 單純拿來寫response
    url = url + "/project/calculate/"
    now = -1  # 記第幾個upload_path，因為有冷媒、消防設備
    for i in range(len(source_type)):
        now += 1
        if source_type[i] != '人為逸散 C1' and source_type[i] != '輸入電力 C2' and source_type[i] != '輸入蒸汽 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source_type[i])))
            link.click()
            time.sleep(1)
            excel_import(driver, source_type, file_path, upload_path, url, now, for_response)
            time.sleep(1)
        elif source_type[i] == '人為逸散 C1':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source_type[i])))
            link.click()
            time.sleep(1)
            excel_import(driver, source_type, file_path, upload_path, url, now, for_response)
            driver.get(url + 'directFugitiveEmission')  # 導回 人為逸散介面
            now += 1
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(2)
            excel_import(driver, source_type, file_path, upload_path, url, now, for_response)
            driver.get(url + 'directFugitiveEmission')
            now += 1
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            excel_import(driver, source_type, file_path, upload_path, url, now, for_response)
            time.sleep(2)
        elif source_type[i] == '輸入電力 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source_type[i])))
            link.click()
            excel_import(driver, source_type, file_path, upload_path, url, now, for_response)

        elif source_type[i] == '輸入蒸汽 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source_type[i])))
            link.click()
            excel_import(driver, source_type, file_path, upload_path, url, now, for_response)

    driver.quit()


def excel_import(driver, source_type, file_path, upload_path, url, now, for_response):
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '匯入 & 匯出')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    import_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//ul//button[span[text()="匯入"]]')))
    driver.execute_script("arguments[0].click()", import_element)
    time.sleep(1)

    file_input = driver.find_element(By.XPATH, '//input[@id="Excel"]')  # 上傳檔案(路徑要自己改)
    temp = os.path.join(file_path, upload_path[now])
    file_input.send_keys(temp)

    time.sleep(1)
    upload_element = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[class='ant-btn ant-btn-primary']")))
    driver.execute_script("arguments[0].click()", upload_element)  # 篩選上傳檔案有儲存按鈕的
    files_without_save_button = {
        'Vehicle.xlsx', 'StationaryCombustionSource.xlsx', 'IndustrialProcess.xlsx',
        'workhour.xlsx', 'other.xlsx', 'Elec.xlsx', 'GreenElec.xlsx', 'steam_plus.xlsx'
    }
    try:
        if upload_path[now] not in files_without_save_button:
            print(upload_path[now])
            time.sleep(1)
            save_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "確 定")]')))
            driver.execute_script("arguments[0].click()", save_element)

            time.sleep(1)
            catch_response(driver, for_response[now])


        else:  # upload_path[now]=='RFG.xlsx' or upload_path[now]=='Fire.xlsx':
            print(upload_path[now])
            # time.sleep(3)
            save_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@class="ant-btn css-dts6b9 ant-btn-primary"]')))
            driver.execute_script("arguments[0].click()", save_element)
            print("成功")
            time.sleep(1)
            catch_response(driver, for_response[now])
    except:
        print("失敗")
    driver.get(url)
    time.sleep(1)
